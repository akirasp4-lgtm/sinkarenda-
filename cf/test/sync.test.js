import { describe, it, expect, vi } from 'vitest';
import { validateGasPayload, sanitizeForStorage, fetchWithRetry, syncAll } from '../src/sync.js';

const HEADERS = ['登録日時','作業日','元請名','現場名','氏名','役割','出勤','退勤','人工',
                 'メモ','夜勤','会社','ID','更新者','色','事業部','工番','作業区分','車両'];

function makeCompactPayload(overrides = {}) {
  return {
    status: 'ok', compact: 1, headers: HEADERS,
    rows: [], members: [], genbaMaster: [], jobsites: [],
    ...overrides
  };
}

// ============================================================
// validateGasPayload（修正3: 応答の妥当性検証）
// ============================================================
describe('validateGasPayload', () => {
  it('現行の19列・並びと一致するcompact応答は通る', () => {
    const out = validateGasPayload(makeCompactPayload());
    expect(out.ok).toBe(true);
  });

  it('compactでない応答は拒否する', () => {
    const out = validateGasPayload({ status: 'ok', rows: [] });
    expect(out.ok).toBe(false);
    expect(out.message).toMatch(/compact/);
  });

  it('headersが1列欠けていたら拒否する（将来GAS側で項目が変わる想定）', () => {
    const out = validateGasPayload(makeCompactPayload({ headers: HEADERS.slice(0, 18) }));
    expect(out.ok).toBe(false);
    expect(out.message).toMatch(/headers/);
  });

  it('headersの並び順が違うだけでも拒否する', () => {
    const swapped = [...HEADERS];
    [swapped[0], swapped[1]] = [swapped[1], swapped[0]];
    const out = validateGasPayload(makeCompactPayload({ headers: swapped }));
    expect(out.ok).toBe(false);
  });

  it('membersが配列でない（欠落してオブジェクト等になっている）と拒否する', () => {
    const out = validateGasPayload(makeCompactPayload({ members: undefined }));
    expect(out.ok).toBe(false);
    expect(out.message).toMatch(/members|配列/);
  });

  it('rows/genbaMaster/jobsitesのいずれかが配列でなくても拒否する', () => {
    expect(validateGasPayload(makeCompactPayload({ rows: {} })).ok).toBe(false);
    expect(validateGasPayload(makeCompactPayload({ genbaMaster: null })).ok).toBe(false);
    expect(validateGasPayload(makeCompactPayload({ jobsites: 'x' })).ok).toBe(false);
  });
});

// ============================================================
// sanitizeForStorage（単価除去 + 忠実な写しの維持）
// ============================================================
describe('sanitizeForStorage', () => {
  it('職人マスタから単価(rate)を落とす', () => {
    const out = sanitizeForStorage(makeCompactPayload({
      members: [{ name: '森', company: 'GRHD', division: 'ICT', rate: 18000 }]
    }));
    expect(out.members[0]).toEqual({ name: '森', company: 'GRHD', division: 'ICT' });
    expect(JSON.stringify(out.members)).not.toContain('18000');
  });

  it('同名・同会社の職人が複数あっても畳まない（元データの重複をそのまま保つ）', () => {
    const out = sanitizeForStorage(makeCompactPayload({
      members: [
        { name: '林電工(りんたろう)', company: '和信カインド', division: '電気', rate: 15000 },
        { name: '林電工(りんたろう)', company: '和信カインド', division: '電気', rate: 15000 },
        { name: '林電工(りんたろう)', company: '和信カインド', division: '電気', rate: 15000 }
      ]
    }));
    expect(out.members).toHaveLength(3);
  });

  it('rowsは中身・順序をそのまま保つ（氏名が空の行=車検期限リマインダー行も捨てない）', () => {
    const row = new Array(19).fill('');
    row[1] = '2026-10-16'; row[3] = 'ハイエース白 車検期限'; row[4] = ''; // 氏名は空
    row[12] = 'VKEN_なにわ432そ8800';
    const out = sanitizeForStorage(makeCompactPayload({ rows: [row] }));
    expect(out.rows).toHaveLength(1);
    expect(out.rows[0][4]).toBe('');
    expect(out.rows[0][12]).toBe('VKEN_なにわ432そ8800');
  });

  it('人工(#NUM!のような非数値)や並び順もそのまま保つ（強制変換しない）', () => {
    const rows = ['c-3', 'a-1', 'b-2', 'a-1'].map(id => {
      const row = new Array(19).fill('');
      row[1] = '2026-05-02'; row[4] = '森'; row[8] = '#NUM!'; row[12] = id;
      return row;
    });
    const out = sanitizeForStorage(makeCompactPayload({ rows }));
    expect(out.rows.map(r => r[12])).toEqual(['c-3', 'a-1', 'b-2', 'a-1']);
    expect(out.rows.every(r => r[8] === '#NUM!')).toBe(true);
  });

  it('genbaMaster/jobsitesは名前が空の行もそのまま保つ（読み取り側でGASと同じ条件で弾く設計のため）', () => {
    const out = sanitizeForStorage(makeCompactPayload({
      genbaMaster: [{ name: '', company: '' }],
      jobsites: [{ genba: '', loc: '', jobNo: '', completed: false, billingMethod: '' }]
    }));
    expect(out.genbaMaster).toHaveLength(1);
    expect(out.jobsites).toHaveLength(1);
  });
});

// ============================================================
// fetchWithRetry（修正3: タイムアウト付き）
// ============================================================
describe('fetchWithRetry', () => {
  it('1回目が404でも2回目で成功すれば結果を返す（GASは5回に1回404を返す）', async () => {
    const calls = [];
    global.fetch = vi.fn(async (u) => {
      calls.push(u);
      if (calls.length === 1) return { ok: false, status: 404, text: async () => '<html>' };
      return { ok: true, status: 200, json: async () => ({ status: 'ok' }) };
    });
    const out = await fetchWithRetry('https://example.test/', 3);
    expect(out).toEqual({ status: 'ok' });
    expect(calls).toHaveLength(2);
  });

  it('回数を使い切ったら投げる', async () => {
    global.fetch = vi.fn(async () => ({ ok: false, status: 404, text: async () => '<html>' }));
    await expect(fetchWithRetry('https://example.test/', 2)).rejects.toThrow(/404/);
  });

  it('毎回のfetchにタイムアウト用のAbortSignalを渡している（無応答で待ち続けない対策）', async () => {
    let seenSignal = null;
    global.fetch = vi.fn(async (u, init) => {
      seenSignal = init && init.signal;
      return { ok: true, status: 200, json: async () => ({ status: 'ok' }) };
    });
    await fetchWithRetry('https://example.test/', 3);
    expect(seenSignal).toBeInstanceOf(AbortSignal);
  });
});

// ============================================================
// syncAll（修正1: スナップショット方式・修正2: 同時実行の抑止）
// ============================================================

// D1のprepare/bind/run/all()を模した簡易・状態保持モック。
// snapshot / sync_lock / sync_log の3テーブルだけを実装する
// （sync.jsが実際に発行するクエリはこの3つだけのため）。
//
// ★再レビュー対応でsync.jsのSQLが以下のように変わった点を反映している:
//   - ロック取得: 「SELECT→INSERT」の2文 → 単一の
//     「INSERT ... ON CONFLICT ... WHERE」。D1同様、meta.changesで成否を返す。
//   - snapshot書き込み: 「INSERT OR REPLACE」（無条件） → 単一の
//     「INSERT ... ON CONFLICT ... WHERE fetch_started_at比較」。
//     古い取得時刻での書き込みは無視され、meta.changesが0になる。
function makeMockDB({ initialLockedAt = null, throwOnSnapshotWrite = false, throwOnLockRead = false } = {}) {
  const state = { snapshot: null, lockedAt: initialLockedAt, syncLog: [] };
  const calls = { snapshotWrites: 0, snapshotWriteAttempts: 0, lockWrites: [], lockAttempts: 0, syncLogWrites: [] };

  function respond(sql, args) {
    return {
      async all() {
        if (throwOnLockRead && /sync_lock/.test(sql)) throw new Error('mock lock read failure');
        if (/SELECT rows, hash, members_count, genba_count, jobsites_count FROM snapshot/.test(sql)) {
          return {
            results: state.snapshot
              ? [{
                  rows: state.snapshot.rows, hash: state.snapshot.hash,
                  members_count: state.snapshot.membersCount, genba_count: state.snapshot.genbaCount,
                  jobsites_count: state.snapshot.jobsitesCount
                }]
              : []
          };
        }
        if (/SELECT payload FROM snapshot/.test(sql)) {
          return { results: state.snapshot ? [{ payload: state.snapshot.payload }] : [] };
        }
        if (/SELECT rows, bytes, at FROM snapshot/.test(sql)) {
          return { results: state.snapshot ? [{ rows: state.snapshot.rows, bytes: state.snapshot.bytes, at: state.snapshot.at }] : [] };
        }
        if (/SELECT ok, message FROM sync_log/.test(sql)) {
          return { results: [...state.syncLog].sort((a, b) => b.at.localeCompare(a.at)) };
        }
        if (/FROM sync_log/.test(sql)) {
          return { results: [...state.syncLog].sort((a, b) => b.at.localeCompare(a.at)) };
        }
        return { results: [] };
      },
      async run() {
        if (throwOnLockRead && /sync_lock/.test(sql) && !/VALUES \(1, NULL\)/.test(sql)) {
          throw new Error('mock lock read failure');
        }
        if (/VALUES \(1, NULL\)/.test(sql) && /sync_lock/.test(sql)) {
          state.lockedAt = null;
          return { success: true, meta: { changes: 1 } };
        }
        if (/INSERT INTO sync_lock/.test(sql) && /ON CONFLICT/.test(sql)) {
          calls.lockAttempts++;
          const [newLockedAt, staleCutoff] = args;
          const cutoff = Number(staleCutoff);
          const isFree = state.lockedAt == null || Number(state.lockedAt) < cutoff;
          if (!isFree) return { success: true, meta: { changes: 0 } };
          state.lockedAt = newLockedAt;
          calls.lockWrites.push(newLockedAt);
          return { success: true, meta: { changes: 1 } };
        }
        if (/INSERT INTO snapshot/.test(sql) && /ON CONFLICT/.test(sql)) {
          calls.snapshotWriteAttempts++;
          if (throwOnSnapshotWrite) throw new Error('mock snapshot write failure');
          const [payload, hash, rows, membersCount, genbaCount, jobsitesCount, bytes, fetchStartedAt, at] = args;
          const isNewer = !state.snapshot || Number(fetchStartedAt) >= Number(state.snapshot.fetchStartedAt);
          if (!isNewer) return { success: true, meta: { changes: 0 } };
          state.snapshot = { payload, hash, rows, membersCount, genbaCount, jobsitesCount, bytes, fetchStartedAt, at };
          calls.snapshotWrites++;
          return { success: true, meta: { changes: 1 } };
        }
        if (/INSERT OR REPLACE INTO sync_log/.test(sql)) {
          const [at, rows, ok, message] = args;
          // ★実際のschema.sqlではat(TEXT)がPRIMARY KEYのため、同じatでの書き込みは
          // 新しい行としてpushされず、既存行を置き換える（INSERT OR REPLACEの本来の意味）。
          // これを素朴なpush()だけにすると、テストの高速な連続呼び出しで同一ミリ秒の
          // at が発生した際に「本来は1行のはずが2行できてしまう」というモック特有の
          // 不具合でテストがまれに揺れる（本番のD1では起こらない）。
          const idx = state.syncLog.findIndex(l => l.at === at);
          const entry = { at, rows, ok, message };
          if (idx >= 0) state.syncLog[idx] = entry; else state.syncLog.push(entry);
          calls.syncLogWrites.push(entry);
          return { success: true };
        }
        return { success: true };
      }
    };
  }

  const db = {
    prepare(sql) {
      return {
        bind: (...args) => respond(sql, args),
        all: () => respond(sql, []).all(),
        run: () => respond(sql, []).run()
      };
    }
  };

  return { db, state, calls };
}

function mockFetchOk(payload) {
  global.fetch = vi.fn(async () => ({
    ok: true, status: 200, json: async () => payload
  }));
}

function makeRows(n, extra = {}) {
  return Array.from({ length: n }, (_, i) => {
    const row = new Array(19).fill('');
    row[1] = '2026-05-02'; row[4] = '作業員' + i; row[8] = 1; row[11] = 'グローライズ'; row[12] = 'id-' + i;
    return row;
  });
}

describe('syncAll（正常系）', () => {
  it('初回同期はsnapshotへの単一のINSERT OR REPLACEだけで完結する（原子性の直接的な証拠）', async () => {
    mockFetchOk(makeCompactPayload({ rows: makeRows(600) }));
    const { db, state, calls } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    const out = await syncAll(env);

    expect(out.ok).toBe(true);
    expect(out.rows).toBe(600);
    // ★重大1（原子性）の直接的な証拠: snapshotへの書き込みは1回のrun()だけ。
    // 旧設計（DELETE+500文ずつのbatch）のように複数の文にまたがらない。
    expect(calls.snapshotWrites).toBe(1);
    expect(state.snapshot).toBeTruthy();
    expect(state.snapshot.rows).toBe(600);
  });

  it('★重大2（費用）の直接的な証拠: 600件の日報でもD1への書き込みはsnapshot 1行+sync_log 1行だけ（旧設計なら600+α行）', async () => {
    mockFetchOk(makeCompactPayload({ rows: makeRows(600) }));
    const { db, calls } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };
    await syncAll(env);
    expect(calls.snapshotWrites).toBe(1);
    expect(calls.syncLogWrites).toHaveLength(1);
  });

  it('ロックは同期完了後に解放される（連続呼び出しをブロックしたままにしない）', async () => {
    mockFetchOk(makeCompactPayload({ rows: makeRows(5) }));
    const { db, state } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };
    await syncAll(env);
    expect(state.lockedAt).toBeNull();
  });
});

describe('syncAll（修正1: 変更が無ければ書かない）', () => {
  it('2回連続で同じ内容を同期すると、2回目は書き込みをスキップしてok:trueで返す', async () => {
    const payload = makeCompactPayload({ rows: makeRows(300) });
    mockFetchOk(payload);
    const { db, calls } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    const first = await syncAll(env);
    expect(first.ok).toBe(true);
    expect(calls.snapshotWrites).toBe(1);

    const second = await syncAll(env);
    expect(second.ok).toBe(true);
    expect(second.skipped).toBe(true);
    expect(second.message).toMatch(/変更なし/);
    // 2回目はsnapshotへの書き込みが増えていない（スキップされた）
    expect(calls.snapshotWrites).toBe(1);
  });

  it('内容が変わっていれば2回目も書き込む', async () => {
    const { db, calls } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    mockFetchOk(makeCompactPayload({ rows: makeRows(300) }));
    await syncAll(env);
    expect(calls.snapshotWrites).toBe(1);

    mockFetchOk(makeCompactPayload({ rows: makeRows(301) }));
    const second = await syncAll(env);
    expect(second.ok).toBe(true);
    expect(second.skipped).toBeFalsy();
    expect(calls.snapshotWrites).toBe(2);
  });
});

describe('syncAll（修正1: サイズガード）', () => {
  it('payloadが1,500,000バイトを超えたら書き込まず失敗として記録する', async () => {
    // 1行あたり十分に長いmemoを持たせてサイズ上限を超えさせる
    const bigRow = () => {
      const row = new Array(19).fill('');
      row[1] = '2026-05-02'; row[4] = '森'; row[9] = 'x'.repeat(2000); row[12] = 'id';
      return row;
    };
    const rows = Array.from({ length: 1000 }, bigRow); // 概算 1000 * (~2000+α) > 1,500,000
    mockFetchOk(makeCompactPayload({ rows }));
    const { db, state, calls } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    const out = await syncAll(env);

    expect(out.ok).toBe(false);
    expect(out.message).toMatch(/上限|バイト/);
    expect(calls.snapshotWrites).toBe(0);
    expect(state.snapshot).toBeNull();
    expect(calls.syncLogWrites).toHaveLength(1);
    expect(calls.syncLogWrites[0].ok).toBe(0);
  });

  it('サイズ超過で失敗しても、既存の（前回成功した）snapshotは変わらず残る', async () => {
    const { db, state } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    mockFetchOk(makeCompactPayload({ rows: makeRows(300) }));
    await syncAll(env);
    const before = { ...state.snapshot };

    const bigRow = () => {
      const row = new Array(19).fill('');
      row[1] = '2026-05-02'; row[4] = '森'; row[9] = 'x'.repeat(2000); row[12] = 'id';
      return row;
    };
    mockFetchOk(makeCompactPayload({ rows: Array.from({ length: 1000 }, bigRow) }));
    const out = await syncAll(env);

    expect(out.ok).toBe(false);
    expect(state.snapshot).toEqual(before);
  });
});

describe('syncAll（修正3: 急激な件数減少ガード）', () => {
  it('保存済みスナップショットがあり、新しいrowsが半分未満なら書き込まず失敗として記録する', async () => {
    const { db, state, calls } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    mockFetchOk(makeCompactPayload({ rows: makeRows(600) }));
    await syncAll(env);
    expect(state.snapshot.rows).toBe(600);

    mockFetchOk(makeCompactPayload({ rows: makeRows(299) })); // 600の半分(300)未満
    const out = await syncAll(env);

    expect(out.ok).toBe(false);
    expect(out.message).toMatch(/急減|半分|異常/);
    // 既存のsnapshotは変わらない（前回の600件のまま）
    expect(state.snapshot.rows).toBe(600);
    expect(calls.syncLogWrites.at(-1).ok).toBe(0);
  });

  it('ちょうど半分は「半分未満」ではないので通す', async () => {
    const { db, state } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    mockFetchOk(makeCompactPayload({ rows: makeRows(600) }));
    await syncAll(env);

    mockFetchOk(makeCompactPayload({ rows: makeRows(300) })); // ちょうど半分
    const out = await syncAll(env);
    expect(out.ok).toBe(true);
    expect(state.snapshot.rows).toBe(300);
  });

  it('初回（保存済みが無い）ときはこの検査を飛ばす（0件近くでも通る）', async () => {
    mockFetchOk(makeCompactPayload({ rows: makeRows(1) }));
    const { db, state } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    const out = await syncAll(env);
    expect(out.ok).toBe(true);
    expect(state.snapshot.rows).toBe(1);
  });
});

describe('syncAll（修正3: 応答の妥当性検証との結線）', () => {
  it('headersが想定と違う応答は書き込まず失敗として記録する（GAS側の項目名変更を想定）', async () => {
    mockFetchOk({ status: 'ok', compact: 1, headers: HEADERS.slice(0, 18), rows: [], members: [], genbaMaster: [], jobsites: [] });
    const { db, state, calls } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    const out = await syncAll(env);

    expect(out.ok).toBe(false);
    expect(state.snapshot).toBeNull();
    expect(calls.syncLogWrites[0].ok).toBe(0);
  });

  it('membersが配列でない（欠落）応答は「0件として保存され成功する」ことなく失敗として記録する', async () => {
    mockFetchOk({ status: 'ok', compact: 1, headers: HEADERS, rows: [], members: undefined, genbaMaster: [], jobsites: [] });
    const { db, state } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    const out = await syncAll(env);

    expect(out.ok).toBe(false);
    expect(state.snapshot).toBeNull(); // 職人マスタが0件のまま保存されて成功記録されることが無い
  });
});

describe('syncAll（修正2: 同時実行の抑止）', () => {
  it('直近ロックが取得されている（進行中とみなせる）間は、fetchすら行わずスキップする', async () => {
    const fetchMock = vi.fn();
    global.fetch = fetchMock;
    const { db } = makeMockDB({ initialLockedAt: String(Date.now() - 1000) }); // 1秒前＝進行中とみなす
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    const out = await syncAll(env);

    expect(out.ok).toBe(true);
    expect(out.skipped).toBe(true);
    expect(out.message).toMatch(/進行中/);
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it('古いロック（前回が異常終了して解放されなかった想定）は上書きして実行する', async () => {
    mockFetchOk(makeCompactPayload({ rows: makeRows(5) }));
    const { db, state } = makeMockDB({ initialLockedAt: String(Date.now() - 999999) }); // 十分に古い
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    const out = await syncAll(env);

    expect(out.ok).toBe(true);
    expect(out.skipped).toBeFalsy();
    expect(state.snapshot).toBeTruthy();
  });

  it('ロック機構自体が読めなくても同期は続行する（フェイルオープン。正しさは書き込みの原子性が担保）', async () => {
    mockFetchOk(makeCompactPayload({ rows: makeRows(5) }));
    const { db, state } = makeMockDB({ throwOnLockRead: true });
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    const out = await syncAll(env);

    expect(out.ok).toBe(true);
    expect(state.snapshot).toBeTruthy();
  });
});

describe('syncAll（修正3・再レビュー: members/genbaMaster/jobsitesの半減・全消えガード）', () => {
  it('職人マスタが全消え（保存済み2件→今回0件）なら書き込みを拒否する（レビューのCodex再現ケース）', async () => {
    const { db, state, calls } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    mockFetchOk(makeCompactPayload({
      rows: makeRows(5),
      members: [{ name: '森', company: 'グローライズ', division: '' }, { name: '田中', company: 'グローライズ', division: '' }]
    }));
    await syncAll(env);
    expect(state.snapshot.membersCount).toBe(2);

    mockFetchOk(makeCompactPayload({ rows: makeRows(5), members: [] })); // 職人マスタが空配列で返る
    const out = await syncAll(env);

    expect(out.ok).toBe(false);
    expect(out.message).toMatch(/職人マスタ/);
    // 既存の職人マスタ件数は変わらない（書き込まれていない）
    expect(state.snapshot.membersCount).toBe(2);
    expect(calls.syncLogWrites.at(-1).ok).toBe(0);
  });

  it('元請マスタが半分未満に減ったら書き込みを拒否する', async () => {
    const { db, state } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    mockFetchOk(makeCompactPayload({
      rows: makeRows(5),
      genbaMaster: [{ name: 'A', company: '' }, { name: 'B', company: '' }, { name: 'C', company: '' }, { name: 'D', company: '' }]
    }));
    await syncAll(env);
    expect(state.snapshot.genbaCount).toBe(4);

    mockFetchOk(makeCompactPayload({ rows: makeRows(5), genbaMaster: [{ name: 'A', company: '' }] })); // 4→1（半分未満）
    const out = await syncAll(env);

    expect(out.ok).toBe(false);
    expect(out.message).toMatch(/元請マスタ/);
    expect(state.snapshot.genbaCount).toBe(4);
  });

  it('現場マスタが半分未満に減ったら書き込みを拒否する', async () => {
    const { db, state } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    const jobsites4 = ['a', 'b', 'c', 'd'].map(loc => ({ genba: 'G', loc, jobNo: '', completed: false, billingMethod: '応援' }));
    mockFetchOk(makeCompactPayload({ rows: makeRows(5), jobsites: jobsites4 }));
    await syncAll(env);
    expect(state.snapshot.jobsitesCount).toBe(4);

    mockFetchOk(makeCompactPayload({ rows: makeRows(5), jobsites: [jobsites4[0]] })); // 4→1
    const out = await syncAll(env);

    expect(out.ok).toBe(false);
    expect(out.message).toMatch(/現場マスタ/);
    expect(state.snapshot.jobsitesCount).toBe(4);
  });

  it('日報が変わらずマスタだけ増えるのは半減ではないので通す（回帰確認）', async () => {
    const { db, state } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    mockFetchOk(makeCompactPayload({ rows: makeRows(5), members: [{ name: '森', company: 'G', division: '' }] }));
    await syncAll(env);

    mockFetchOk(makeCompactPayload({
      rows: makeRows(5),
      members: [{ name: '森', company: 'G', division: '' }, { name: '田中', company: 'G', division: '' }]
    }));
    const out = await syncAll(env);
    expect(out.ok).toBe(true);
    expect(state.snapshot.membersCount).toBe(2);
  });
});

describe('syncAll（修正7: 急減ガードが自己回復しない失敗ループにならないこと）', () => {
  it('連続3回拒否された後の4回目は自動的に受け入れる（アーカイブ等の正当な大幅減が続いた場合の自己回復）', async () => {
    const { db, state, calls } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    // ★sync_log.at はミリ秒分解能のISO文字列（かつ実schemaではPRIMARY KEY）のため、
    // テストのように短時間に何度も呼ぶと同一ミリ秒に当たりうる。連続拒否の判定は
    // 「直近の行が何個連続で拒否ログか」を見るため、時刻を明示的に進めて
    // 各呼び出しのatを確実に一意にする（本番はCronが5分間隔なので衝突しない）。
    let t = 1_700_000_000_000;
    vi.useFakeTimers();
    try {
      vi.setSystemTime(t);
      mockFetchOk(makeCompactPayload({ rows: makeRows(600) }));
      await syncAll(env);
      expect(state.snapshot.rows).toBe(600);

      mockFetchOk(makeCompactPayload({ rows: makeRows(100) })); // 600の半分未満
      t += 1000; vi.setSystemTime(t);
      const r1 = await syncAll(env);
      t += 1000; vi.setSystemTime(t);
      const r2 = await syncAll(env);
      t += 1000; vi.setSystemTime(t);
      const r3 = await syncAll(env);
      expect([r1, r2, r3].every(r => r.ok === false)).toBe(true);
      expect(state.snapshot.rows).toBe(600); // 3回とも拒否され、既存のまま

      t += 1000; vi.setSystemTime(t);
      const r4 = await syncAll(env);
      expect(r4.ok).toBe(true);
      expect(r4.message).toMatch(/自動的に受け入れ/);
      expect(state.snapshot.rows).toBe(100); // 4回目で自己回復し受け入れられる
      expect(calls.syncLogWrites.filter(l => l.ok === 0)).toHaveLength(3);
    } finally {
      vi.useRealTimers();
    }
  });

  it('force:trueを指定すると、1回目の拒否条件でも即座に受け入れる（利用者が今すぐ反映したい場合の脱出口）', async () => {
    const { db, state } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    mockFetchOk(makeCompactPayload({ rows: makeRows(600) }));
    await syncAll(env);

    mockFetchOk(makeCompactPayload({ rows: makeRows(100) }));
    const out = await syncAll(env, { force: true });

    expect(out.ok).toBe(true);
    expect(out.message).toMatch(/force=1/);
    expect(state.snapshot.rows).toBe(100);
  });

  it('force:trueでも応答形式検証・サイズ上限は無条件のまま維持される（forceは急減ガードのみの脱出口）', async () => {
    const { db, state } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };
    mockFetchOk({ status: 'ok', compact: 1, headers: HEADERS.slice(0, 18), rows: [], members: [], genbaMaster: [], jobsites: [] });

    const out = await syncAll(env, { force: true });
    expect(out.ok).toBe(false);
    expect(state.snapshot).toBeNull();
  });
});

describe('syncAll（修正2・再レビュー: ロック取得の単一SQL化）', () => {
  it('同時に2回呼んでも、ロックを取得できるのは1回だけ（SELECT→INSERTの2文に分かれていた旧設計では両方成功しえた）', async () => {
    mockFetchOk(makeCompactPayload({ rows: makeRows(5) }));
    const { db, calls } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    const [a, b] = await Promise.all([syncAll(env), syncAll(env)]);
    const progressing = [a, b].filter(r => r.skipped && /進行中/.test(r.message));
    const executed = [a, b].filter(r => !(r.skipped && /進行中/.test(r.message)));

    // どちらか一方だけがロックを取得して実行され、もう一方は「進行中」でスキップされる。
    // 両方が実行される（＝両方がロックを取得できてしまう）ことは無い。
    expect(progressing).toHaveLength(1);
    expect(executed).toHaveLength(1);
    expect(calls.snapshotWrites).toBe(1);
  });
});

describe('syncAll（修正2・再レビュー: 世代の逆転防止＝古い取得結果は新しい結果を上書きしない）', () => {
  it('先に始まったが後から完了する取得（古いfetch_started_at）は、既に保存済みのより新しい取得結果を上書きしない（Codexが再現した競合の再現テスト）', async () => {
    const { db, state, calls } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };
    const nowSpy = vi.spyOn(Date, 'now');

    try {
      // 「B」＝後から始まったが先に完了する取得（新しいfetch_started_at）
      nowSpy.mockReturnValue(2_000_000);
      mockFetchOk(makeCompactPayload({ rows: makeRows(3).map((r, i) => { r[12] = 'NEW-' + i; return r; }) }));
      const first = await syncAll(env);
      expect(first.ok).toBe(true);
      expect(calls.snapshotWrites).toBe(1);
      const afterB = state.snapshot.payload;
      expect(afterB).toContain('NEW-0');

      // 「A」＝先に始まったが後から完了する取得（Bより古いfetch_started_at）。
      // 現実のWorkerでは、Aのfetchが遅延しBの完了後に届くとこの状態になる。
      nowSpy.mockReturnValue(1_000_000); // Bより古い取得開始時刻
      mockFetchOk(makeCompactPayload({ rows: makeRows(3).map((r, i) => { r[12] = 'OLD-' + i; return r; }) }));
      const second = await syncAll(env);

      // ★世代の逆転防止：古い取得時刻の内容は書き込まれない（失敗ではなくskipped:trueで正常終了）
      expect(second.ok).toBe(true);
      expect(second.skipped).toBe(true);
      expect(second.message).toMatch(/より新しい取得結果/);
      // snapshotはBの内容のまま（Aで上書きされていない）＝古いデータが新しいデータを上書きしないことの証拠
      expect(state.snapshot.payload).toBe(afterB);
      expect(state.snapshot.payload).toContain('NEW-0');
      expect(state.snapshot.payload).not.toContain('OLD-0');
      expect(calls.snapshotWrites).toBe(1); // Aの書き込みは実際には行われていない
    } finally {
      nowSpy.mockRestore();
    }
  });
});

describe('syncAll（修正1・再レビュー: ハッシュ一致スキップでも鮮度ログは更新される）', () => {
  it('2回目が変更なしでスキップされても、sync_logには新しいok=1の行が追加される（読み取り側の鮮度ガードが「変更が無いだけ」を「古い」と誤判定しないための前提）', async () => {
    const payload = makeCompactPayload({ rows: makeRows(10) });
    mockFetchOk(payload);
    const { db, calls } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    await syncAll(env);
    expect(calls.syncLogWrites).toHaveLength(1);
    expect(calls.syncLogWrites[0].ok).toBe(1);

    mockFetchOk(payload); // 同じ内容
    const second = await syncAll(env);
    expect(second.skipped).toBe(true);
    expect(calls.syncLogWrites).toHaveLength(2);
    expect(calls.syncLogWrites[1].ok).toBe(1); // 変更なしも「成功」として記録される
    expect(calls.syncLogWrites[1].message).toMatch(/変更なし/);
  });
});

describe('syncAll（例外を投げない契約の維持）', () => {
  it('fetchが例外を投げても、syncAllは例外を投げずok:falseを返し、sync_logにok=0が記録される', async () => {
    global.fetch = vi.fn(async () => { throw new Error('mock network failure'); });
    const { db, calls, state } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    const out = await syncAll(env);

    expect(out.ok).toBe(false);
    expect(out.rows).toBe(0);
    expect(out.message).toMatch(/mock network failure/);
    expect(state.snapshot).toBeNull();
    expect(calls.syncLogWrites).toHaveLength(1);
    expect(calls.syncLogWrites[0].ok).toBe(0);
  });

  it('snapshotへの書き込み自体が例外を投げても、syncAllは例外を投げずok:falseを返す', async () => {
    mockFetchOk(makeCompactPayload({ rows: makeRows(5) }));
    const { db, calls } = makeMockDB({ throwOnSnapshotWrite: true });
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    const out = await syncAll(env);

    expect(out.ok).toBe(false);
    expect(out.message).toMatch(/mock snapshot write failure/);
    expect(calls.syncLogWrites).toHaveLength(1);
    expect(calls.syncLogWrites[0].ok).toBe(0);
  });

  it('ロックは例外発生時でも必ず解放される（finally）', async () => {
    mockFetchOk(makeCompactPayload({ rows: makeRows(5) }));
    const { db, state } = makeMockDB({ throwOnSnapshotWrite: true });
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    await syncAll(env);
    expect(state.lockedAt).toBeNull();
  });
});
