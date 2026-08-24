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
function makeMockDB({ initialLockedAt = null, throwOnSnapshotWrite = false, throwOnLockRead = false } = {}) {
  const state = { snapshot: null, lockedAt: initialLockedAt, syncLog: [] };
  const calls = { snapshotWrites: 0, lockWrites: [], syncLogWrites: [] };

  function respond(sql, args) {
    return {
      async all() {
        if (throwOnLockRead && /FROM sync_lock/.test(sql)) throw new Error('mock lock read failure');
        if (/SELECT locked_at FROM sync_lock/.test(sql)) {
          return { results: state.lockedAt != null ? [{ locked_at: state.lockedAt }] : [] };
        }
        if (/SELECT rows, hash FROM snapshot/.test(sql)) {
          return { results: state.snapshot ? [{ rows: state.snapshot.rows, hash: state.snapshot.hash }] : [] };
        }
        if (/SELECT payload FROM snapshot/.test(sql)) {
          return { results: state.snapshot ? [{ payload: state.snapshot.payload }] : [] };
        }
        if (/SELECT rows, bytes, at FROM snapshot/.test(sql)) {
          return { results: state.snapshot ? [{ rows: state.snapshot.rows, bytes: state.snapshot.bytes, at: state.snapshot.at }] : [] };
        }
        if (/FROM sync_log/.test(sql)) {
          return { results: [...state.syncLog].sort((a, b) => b.at.localeCompare(a.at)) };
        }
        return { results: [] };
      },
      async run() {
        if (/VALUES \(1, NULL\)/.test(sql) && /sync_lock/.test(sql)) {
          state.lockedAt = null;
          return { success: true };
        }
        if (/INSERT OR REPLACE INTO sync_lock/.test(sql)) {
          state.lockedAt = args[0];
          calls.lockWrites.push(args[0]);
          return { success: true };
        }
        if (/INSERT OR REPLACE INTO snapshot/.test(sql)) {
          if (throwOnSnapshotWrite) throw new Error('mock snapshot write failure');
          const [payload, hash, rows, bytes, at] = args;
          state.snapshot = { payload, hash, rows, bytes, at };
          calls.snapshotWrites++;
          return { success: true };
        }
        if (/INSERT OR REPLACE INTO sync_log/.test(sql)) {
          const [at, rows, ok, message] = args;
          state.syncLog.push({ at, rows, ok, message });
          calls.syncLogWrites.push({ at, rows, ok, message });
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
