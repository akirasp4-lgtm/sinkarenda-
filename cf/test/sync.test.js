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
        // ★3回目レビュー修正3: 急減ガードの自己回復（sameHashShrinkRejectStreak）が
        // payload_hash列も含めて取得するようになった。
        if (/SELECT at, ok, message, payload_hash FROM sync_log/.test(sql)) {
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
          // ★3回目レビュー修正4: 本番のWHERE条件が `>=` から `>` に変わったことに合わせる。
          // 初回（保存済みが無い）はON CONFLICTに入らず素直にINSERTされるのでtrue。
          // 2回目以降は「保存済みより厳密に新しい」ときだけ上書きできる（同着は不可）。
          const isNewer = !state.snapshot || Number(fetchStartedAt) > Number(state.snapshot.fetchStartedAt);
          if (!isNewer) return { success: true, meta: { changes: 0 } };
          state.snapshot = { payload, hash, rows, membersCount, genbaCount, jobsitesCount, bytes, fetchStartedAt, at };
          calls.snapshotWrites++;
          return { success: true, meta: { changes: 1 } };
        }
        if (/INSERT OR REPLACE INTO sync_log/.test(sql)) {
          const [at, rows, ok, message, payloadHash] = args;
          // ★実際のschema.sqlではat(TEXT)がPRIMARY KEYのため、同じatでの書き込みは
          // 新しい行としてpushされず、既存行を置き換える（INSERT OR REPLACEの本来の意味）。
          // これを素朴なpush()だけにすると、テストの高速な連続呼び出しで同一ミリ秒の
          // at が発生した際に「本来は1行のはずが2行できてしまう」というモック特有の
          // 不具合でテストがまれに揺れる（本番のD1では起こらない）。
          const idx = state.syncLog.findIndex(l => l.at === at);
          const entry = { at, rows, ok, message, payload_hash: payloadHash ?? null };
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

// ★3回目レビュー修正4（世代ガードの同着対策）でsnapshot書き込みのWHERE条件が
// `>=` から `>`（同一ミリ秒は不可）に変わったため、同じテスト内で複数回 syncAll() を
// 呼んで「両方とも書き込まれる」ことを期待するテストは、Date.now() が実際に進む
// ことを保証しないと、実機のD1（fetchのたびに本物の通信時間が経つので実質衝突しない）
// と違ってテスト環境（fetchは即座に解決する）では同一ミリ秒に衝突しうる
// （実行速度に依存するため、直さないと「たまに落ちるテスト」になる）。
// 使い終わったら必ず stop() でspyを元に戻すこと。
function stubIncreasingClock(startAt = 1_000_000, stepMs = 1000) {
  let t = startAt;
  const spy = vi.spyOn(Date, 'now').mockImplementation(() => { const v = t; t += stepMs; return v; });
  return { stop: () => spy.mockRestore() };
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
    // ★6回目レビュー修正1: 「変更なしスキップ」にはskipReason:'unchanged'が付く。
    // sync-guard.jsのdecideSyncOutcomeがこれを確実成功として扱うための合図。
    expect(second.skipReason).toBe('unchanged');
    // 2回目はsnapshotへの書き込みが増えていない（スキップされた）
    expect(calls.snapshotWrites).toBe(1);
  });


  it('内容が変わっていれば2回目も書き込む', async () => {
    const { db, calls } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };
    // ★修正4でfetch_started_atの比較が`>`（同着不可）になったため、2回とも書き込まれる
    // ことを期待するこのテストはDate.now()を明示的に単調増加させる（stubIncreasingClock参照）。
    const clock = stubIncreasingClock();

    try {
      mockFetchOk(makeCompactPayload({ rows: makeRows(300) }));
      await syncAll(env);
      expect(calls.snapshotWrites).toBe(1);

      mockFetchOk(makeCompactPayload({ rows: makeRows(301) }));
      const second = await syncAll(env);
      expect(second.ok).toBe(true);
      expect(second.skipped).toBeFalsy();
      expect(calls.snapshotWrites).toBe(2);
    } finally {
      clock.stop();
    }
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
    // ★5回目レビュー修正作業中に発見（既存の潜在的な不具合。今回のfix範囲外だが
    // ついでに直した）: このテストはDate.now()を進めずに2回syncAll()を呼んでいた
    // ため、実行が速い環境では2回とも同一ミリ秒のfetch_started_atになり、世代逆転
    // 防止のWHERE条件（`>`。3回目レビュー修正4）に阻まれて2回目の書き込みが
    // まれに失敗する（テストがまれに揺れる）不具合があった。他の同種テストと
    // 同じくstubIncreasingClock()で時計を単調増加させる。
    const clock = stubIncreasingClock();
    try {
      mockFetchOk(makeCompactPayload({ rows: makeRows(600) }));
      await syncAll(env);

      mockFetchOk(makeCompactPayload({ rows: makeRows(300) })); // ちょうど半分
      const out = await syncAll(env);
      expect(out.ok).toBe(true);
      expect(state.snapshot.rows).toBe(300);
    } finally {
      clock.stop();
    }
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
    // ★6回目レビュー修正1: 「進行中のためスキップ」はGASへ一度も取得しに行って
    // いないため、skipReasonを付けない（decideSyncOutcomeが確実成功として
    // 誤って扱わないようにするための区別）。
    expect(out.skipReason).toBeUndefined();
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
    // ★修正4の同着対策により、2回とも書き込まれることを期待するのでDate.now()を進める。
    const clock = stubIncreasingClock();

    try {
      mockFetchOk(makeCompactPayload({ rows: makeRows(5), members: [{ name: '森', company: 'G', division: '' }] }));
      await syncAll(env);

      mockFetchOk(makeCompactPayload({
        rows: makeRows(5),
        members: [{ name: '森', company: 'G', division: '' }, { name: '田中', company: 'G', division: '' }]
      }));
      const out = await syncAll(env);
      expect(out.ok).toBe(true);
      expect(state.snapshot.membersCount).toBe(2);
    } finally {
      clock.stop();
    }
  });
});

describe('syncAll（3回目レビュー修正3: 急減ガードの自己回復は「同一内容が30分」でなければ進まない・作り直し）', () => {
  // ★旧実装（回数だけを見る版）はCodexにより「日報→職人→元請→現場と毎回まったく
  // 別の欠損を起こしても、4回目が“3回連続拒否”の条件を満たして自動受入されてしまう」
  // ことを再現された。ここでは新実装（ハッシュ一致＋経過時間）がその脆弱性を
  // 閉じていることと、正当なケース（同一内容が続く）では従来どおり自己回復することの
  // 両方を確認する。
  it('同一内容（ハッシュ一致）の拒否がCronの間隔(5分)で続いても、最初の拒否から30分経つまでは受け入れない', async () => {
    const { db, state } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };
    const CRON_INTERVAL_MS = 5 * 60 * 1000;

    let t = 1_700_000_000_000;
    vi.useFakeTimers();
    try {
      vi.setSystemTime(t);
      mockFetchOk(makeCompactPayload({ rows: makeRows(600) }));
      await syncAll(env);
      expect(state.snapshot.rows).toBe(600);

      // 以後ずっと同じ内容（＝同じハッシュ）の急減応答を返す（実際にそういう欠損が
      // 起きて、GAS側の状態が変わらないまま続いている状況を模す）。
      mockFetchOk(makeCompactPayload({ rows: makeRows(100) })); // 600の半分未満

      // Cronの間隔(5分)どおり6回（=30分ぶん）呼ぶ。最初の拒否からちょうど30分に
      // 達するのは7回目なので、この6回はすべて拒否されるはず。
      const results = [];
      for (let i = 0; i < 6; i++) {
        t += CRON_INTERVAL_MS; vi.setSystemTime(t);
        results.push(await syncAll(env));
      }
      expect(results.every(r => r.ok === false)).toBe(true);
      expect(state.snapshot.rows).toBe(600); // 既存のまま

      // 7回目＝最初の拒否からちょうど30分後。ここで初めて自動的に受け入れる。
      t += CRON_INTERVAL_MS; vi.setSystemTime(t);
      const accepted = await syncAll(env);
      expect(accepted.ok).toBe(true);
      expect(accepted.message).toMatch(/自動的に受け入れ/);
      expect(state.snapshot.rows).toBe(100);
    } finally {
      vi.useRealTimers();
    }
  });

  it('30分経っていなければ、何度呼んでも（＝/api/syncを連打しても）自動受入は早まらない', async () => {
    const { db, state } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    let t = 1_700_000_000_000;
    vi.useFakeTimers();
    try {
      vi.setSystemTime(t);
      mockFetchOk(makeCompactPayload({ rows: makeRows(600) }));
      await syncAll(env);

      mockFetchOk(makeCompactPayload({ rows: makeRows(100) }));
      // 5分間隔ではなく1秒間隔で20回連打しても、経過時間そのものは変わらない
      // （＝回数を稼いでも早く受け入れられることはない）。
      let last = null;
      for (let i = 0; i < 20; i++) {
        t += 1000; vi.setSystemTime(t);
        last = await syncAll(env);
      }
      expect(last.ok).toBe(false);
      expect(state.snapshot.rows).toBe(600);
    } finally {
      vi.useRealTimers();
    }
  });

  it('日報→職人→元請→現場と毎回まったく別の欠損を送りつけても自動受入されない（Codexが旧実装で再現した脆弱性の再現テスト）', async () => {
    const { db, state } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };
    const members = (n) => Array.from({ length: n }, (_, i) => ({ name: '職人' + i, company: 'G', division: '' }));
    const genba = (n) => Array.from({ length: n }, (_, i) => ({ name: '現場' + i, company: '' }));
    const jobsites = (n) => genba(n).map(g => ({ genba: g.name, loc: 'x', jobNo: '', completed: false, billingMethod: '応援' }));

    let t = 1_700_000_000_000;
    vi.useFakeTimers();
    try {
      vi.setSystemTime(t);
      mockFetchOk(makeCompactPayload({ rows: makeRows(100), members: members(10), genbaMaster: genba(10), jobsites: jobsites(10) }));
      await syncAll(env);
      expect(state.snapshot.rows).toBe(100);
      expect(state.snapshot.genbaCount).toBe(10);

      // Codexの再現手順どおり：日報だけ100→10、職人だけ10→1、元請だけ10→1、
      // 現場だけ10→1、と毎回別のテーブルだけを半減させる。/api/syncを短時間で
      // 連打できることを模して1秒おきに送る。
      const attempts = [
        makeCompactPayload({ rows: makeRows(10), members: members(10), genbaMaster: genba(10), jobsites: jobsites(10) }),
        makeCompactPayload({ rows: makeRows(100), members: members(1), genbaMaster: genba(10), jobsites: jobsites(10) }),
        makeCompactPayload({ rows: makeRows(100), members: members(10), genbaMaster: genba(1), jobsites: jobsites(10) }),
        makeCompactPayload({ rows: makeRows(100), members: members(10), genbaMaster: genba(10), jobsites: jobsites(1) })
      ];
      for (const attempt of attempts) {
        t += 1000; vi.setSystemTime(t);
        mockFetchOk(attempt);
        const out = await syncAll(env);
        expect(out.ok).toBe(false);
      }
      // 4回とも拒否され、最初のスナップショットのまま（現場が1件になる等の事故は起きない）
      expect(state.snapshot.rows).toBe(100);
      expect(state.snapshot.genbaCount).toBe(10);
      expect(state.snapshot.jobsitesCount).toBe(10);
    } finally {
      vi.useRealTimers();
    }
  });

  it('force:trueを指定すると、1回目の拒否条件でも即座に受け入れる（利用者が今すぐ反映したい場合の脱出口）', async () => {
    const { db, state } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };
    // ★修正4の同着対策により、2回とも書き込まれることを期待するのでDate.now()を進める。
    const clock = stubIncreasingClock();

    try {
      mockFetchOk(makeCompactPayload({ rows: makeRows(600) }));
      await syncAll(env);

      mockFetchOk(makeCompactPayload({ rows: makeRows(100) }));
      const out = await syncAll(env, { force: true });

      expect(out.ok).toBe(true);
      expect(out.message).toMatch(/force=1/);
      expect(state.snapshot.rows).toBe(100);
    } finally {
      clock.stop();
    }
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

describe('syncAll（4回目レビュー修正5: 自己回復は「同一内容が30分」に加えて最低観測回数・直近観測の鮮度も要求する）', () => {
  // ★指摘: 旧実装（3回目レビュー修正3）は「最初の拒否」からの経過時間だけを見ており、
  // その間の観測が実質0回（＝最初の拒否と今回の“2点”だけ）でも成立してしまっていた。
  // 例: 0分に1回拒否→Cronが何らかの理由で止まる→31分後に同じ欠損が再発、の2点だけで
  // 自動受理される（「30分間ずっと同じだった」ではなく「31分離れた2点で同じだった」
  // しか確認できていない）。ここでは、その穴が実際に塞がれていること（最低観測回数
  // 未満では受け入れない）と、直近観測が古い（監視に空白期間があった）場合も受け入れ
  // ないこと、そして両方の追加条件を満たしたときは従来どおり自己回復することを確認する。
  it('「0分の1回」と「31分後の1回」の2点だけでは自動受入されない（最低観測回数（3回）未満）', async () => {
    const { db, state } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    let t = 1_700_000_000_000;
    vi.useFakeTimers();
    try {
      vi.setSystemTime(t);
      mockFetchOk(makeCompactPayload({ rows: makeRows(600) }));
      await syncAll(env);
      expect(state.snapshot.rows).toBe(600);

      // 0分後: 1回目の拒否（観測1件目）。
      t += 1000; vi.setSystemTime(t);
      mockFetchOk(makeCompactPayload({ rows: makeRows(100) })); // 600の半分未満・以後同一内容
      const first = await syncAll(env);
      expect(first.ok).toBe(false);
      expect(state.snapshot.rows).toBe(600);

      // Cronが31分ほど止まっていた想定（この間、syncAllは一切呼ばれない＝観測が無い）。
      // 31分後、同じ内容がもう一度拒否される。旧実装（経過時間のみ）なら
      // 「最初の拒否から30分経過」を満たして自動受入されてしまっていた。
      t += 31 * 60 * 1000; vi.setSystemTime(t);
      const second = await syncAll(env);

      // 新実装: 直近の拒否ログは1件（＝上の1回目）しか無く、最低観測回数3回に満たない
      // ため、経過時間は満たしていても自動受入されない。
      expect(second.ok).toBe(false);
      expect(state.snapshot.rows).toBe(600); // 事故（現場が1件になる等）は起きない
    } finally {
      vi.useRealTimers();
    }
  });

  it('観測に空白期間があると、その前の観測実績は今回の判定に持ち越さない（空白の後から改めて最低観測回数・経過時間を満たす必要がある）', async () => {
    // ★5回目レビュー修正3（Codexの再指摘）: 旧実装は「直前の1件と今回の間」しか
    // 見ておらず、この直後のテスト（旧版）は「最初の拒否から30分・観測3件」を
    // 満たしていれば、途中に25分の空白があっても「直近観測が新しくなった1回」だけで
    // 自動受理されてしまうことを許していた（＝Codexが指摘した穴そのもの）。
    // 新実装は「隣接する観測間隔が10分を超えた時点で、それより古い観測を今回の
    // 判定に含めない」ため、空白の後は最低観測回数・経過時間をゼロから積み直す
    // 必要があることをここで確認する。
    const { db, state } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };
    const CRON_INTERVAL_MS = 5 * 60 * 1000;

    let t = 1_700_000_000_000;
    vi.useFakeTimers();
    try {
      vi.setSystemTime(t);
      mockFetchOk(makeCompactPayload({ rows: makeRows(600) }));
      await syncAll(env);

      mockFetchOk(makeCompactPayload({ rows: makeRows(100) })); // 以後ずっと同一内容
      // 0分・5分・10分に拒否（観測3件）。
      for (let i = 0; i < 3; i++) {
        t += CRON_INTERVAL_MS; vi.setSystemTime(t);
        const r = await syncAll(env);
        expect(r.ok).toBe(false);
      }
      expect(state.snapshot.rows).toBe(600);

      // ここでCronが25分止まる（10分時点の観測から25分後＝35分時点で再開）。
      // 「最初の拒否(5分時点)から30分」という経過時間だけを見れば満たしてしまう
      // 局面だが、直前の観測(10分時点)からの間隔が25分あり規定(10分)を超えるため、
      // 新実装はここで遡るのを打ち切る＝空白より前の3件の実績は一切カウントされず、
      // count=0からやり直しになる。
      t += 25 * 60 * 1000; vi.setSystemTime(t);
      const afterGap = await syncAll(env);
      expect(afterGap.ok).toBe(false);
      expect(state.snapshot.rows).toBe(600);

      // 空白の後、あらためて5分間隔で観測を続ける。空白直後の1回（afterGap）を
      // 起点として、そこからさらに5回（合計6回・25分）ではまだ30分に届かないため
      // 拒否が続くことを確認する（＝空白前の3件の実績を使い回して早く自動受理される
      // ことは無い）。
      const results = [];
      for (let i = 0; i < 5; i++) {
        t += CRON_INTERVAL_MS; vi.setSystemTime(t);
        results.push(await syncAll(env));
      }
      expect(results.every(r => r.ok === false)).toBe(true);
      expect(state.snapshot.rows).toBe(600);

      // 空白直後の観測（afterGap）からちょうど30分後。ここで初めて自動的に受け入れる。
      t += CRON_INTERVAL_MS; vi.setSystemTime(t);
      const accepted = await syncAll(env);
      expect(accepted.ok).toBe(true);
      expect(accepted.message).toMatch(/自動的に受け入れ/);
      expect(state.snapshot.rows).toBe(100);
    } finally {
      vi.useRealTimers();
    }
  });

  it('Codexが5回目レビューで再現した具体例（0分・1分・2分に拒否→27分の空白→29分・30分に再拒否）では自動受理されない', async () => {
    const { db, state } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    let t = 1_700_000_000_000;
    vi.useFakeTimers();
    try {
      vi.setSystemTime(t);
      mockFetchOk(makeCompactPayload({ rows: makeRows(600) }));
      await syncAll(env);

      mockFetchOk(makeCompactPayload({ rows: makeRows(100) })); // 以後ずっと同一内容

      // 0分・1分・2分に拒否（観測3件・最低観測回数は満たす）。
      for (const deltaMin of [1, 1, 1]) {
        t += deltaMin * 60 * 1000; vi.setSystemTime(t);
        const r = await syncAll(env);
        expect(r.ok).toBe(false);
      }

      // 27分の空白（Cronが止まっていた想定）。
      t += 27 * 60 * 1000; vi.setSystemTime(t);
      const at29 = await syncAll(env); // 最初の拒否(0分)から29分後
      expect(at29.ok).toBe(false);

      // さらに1分後（最初の拒否から30分後）。旧実装（直前1件とのみ比較）だと
      // 「最古から30分・件数5件・直近1分」を満たして自動受理されてしまっていた。
      t += 1 * 60 * 1000; vi.setSystemTime(t);
      const at30 = await syncAll(env);
      expect(at30.ok).toBe(false); // 27分の空白で実績が打ち切られているため、まだ受理されない
      expect(state.snapshot.rows).toBe(600); // 事故（不完全なデータの受理）は起きない
    } finally {
      vi.useRealTimers();
    }
  });
});

describe('syncAll（5回目レビュー修正5: forceはマスタ半減には効かない・force自体にも頻度制限）', () => {
  it('日報のみが急減している場合はforce:trueで即座に受け入れる（マスタは急減していないので従来どおり）', async () => {
    const { db, state } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };
    const clock = stubIncreasingClock();
    try {
      mockFetchOk(makeCompactPayload({
        rows: makeRows(600),
        members: [{ name: '森', company: 'G', division: '' }],
        genbaMaster: [{ name: 'A', company: '' }],
        jobsites: [{ genba: 'A', loc: 'x', jobNo: '', completed: false, billingMethod: '応援' }]
      }));
      await syncAll(env);

      mockFetchOk(makeCompactPayload({
        rows: makeRows(100), // 日報だけ急減
        members: [{ name: '森', company: 'G', division: '' }],
        genbaMaster: [{ name: 'A', company: '' }],
        jobsites: [{ genba: 'A', loc: 'x', jobNo: '', completed: false, billingMethod: '応援' }]
      }));
      const out = await syncAll(env, { force: true });

      expect(out.ok).toBe(true);
      expect(out.message).toMatch(/force=1/);
      expect(state.snapshot.rows).toBe(100);
    } finally {
      clock.stop();
    }
  });

  it('マスタ（職人・元請・現場）が1つでも急減している場合はforce:trueを指定しても即座には受け入れない', async () => {
    const { db, state } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    mockFetchOk(makeCompactPayload({
      rows: makeRows(600),
      members: [
        { name: '森', company: 'G', division: '' }, { name: '田中', company: 'G', division: '' },
        { name: '佐藤', company: 'G', division: '' }, { name: '鈴木', company: 'G', division: '' }
      ]
    }));
    await syncAll(env);

    // 日報は変わらないが職人マスタだけ半減（4→1。force=1が指定されていても即時受理しない）。
    mockFetchOk(makeCompactPayload({
      rows: makeRows(600),
      members: [{ name: '森', company: 'G', division: '' }]
    }));
    const out = await syncAll(env, { force: true });

    expect(out.ok).toBe(false);
    expect(out.message).toMatch(/職人マスタ/);
    expect(out.message).toMatch(/force/); // forceが効かない旨の案内を含む
    expect(state.snapshot.membersCount).toBe(4); // 変わっていない
  });

  it('日報・マスタが同時に急減している場合もforce:trueは無効（マスタ急減が1つでも含まれれば全体としてforceは効かない）', async () => {
    const { db, state } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    mockFetchOk(makeCompactPayload({
      rows: makeRows(600),
      members: [
        { name: '森', company: 'G', division: '' }, { name: '田中', company: 'G', division: '' },
        { name: '佐藤', company: 'G', division: '' }, { name: '鈴木', company: 'G', division: '' }
      ]
    }));
    await syncAll(env);

    mockFetchOk(makeCompactPayload({
      rows: makeRows(100), // 日報も急減
      members: [{ name: '森', company: 'G', division: '' }] // 職人マスタも急減（4→1）
    }));
    const out = await syncAll(env, { force: true });

    expect(out.ok).toBe(false);
    expect(state.snapshot.rows).toBe(600);
    expect(state.snapshot.membersCount).toBe(4);
  });

  it('force:trueでも応答形式検証・サイズ上限は無条件のまま維持される（マスタ半減が絡まないケースでの回帰確認）', async () => {
    const { db, state } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };
    mockFetchOk({ status: 'ok', compact: 1, headers: HEADERS.slice(0, 18), rows: [], members: [], genbaMaster: [], jobsites: [] });

    const out = await syncAll(env, { force: true });
    expect(out.ok).toBe(false);
    expect(state.snapshot).toBeNull();
  });

  it('forceによる即時受理は直近30分に1回まで（force連打・毎回異なるタイミングを狙った悪用の緩和）', async () => {
    const { db, state } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    let t = 1_700_000_000_000;
    vi.useFakeTimers();
    try {
      vi.setSystemTime(t);
      mockFetchOk(makeCompactPayload({ rows: makeRows(600) }));
      await syncAll(env);

      // ★世代逆転防止のWHERE条件（fetch_started_at比較）が同一ミリ秒での上書きを
      // 許さないため、次の呼び出し前に時計を進める（他のテストのstubIncreasingClock
      // と同じ理由。sync.test.jsのコメント参照）。
      t += 1000; vi.setSystemTime(t);
      // 1回目のforce（日報のみ急減）: 受理される。
      mockFetchOk(makeCompactPayload({ rows: makeRows(100) }));
      const first = await syncAll(env, { force: true });
      expect(first.ok).toBe(true);
      expect(state.snapshot.rows).toBe(100);

      // 5分後、再び日報が急減した別の状況でforce=1を使おうとしても、
      // 直近30分以内に既にforceで受理済みのため無効化される。
      t += 5 * 60 * 1000; vi.setSystemTime(t);
      mockFetchOk(makeCompactPayload({ rows: makeRows(10) }));
      const second = await syncAll(env, { force: true });
      expect(second.ok).toBe(false);
      expect(second.message).toMatch(/頻度制限|force/);
      expect(state.snapshot.rows).toBe(100); // 変わっていない

      // 30分経過後は、再びforceが使えるようになる。
      t += 26 * 60 * 1000; vi.setSystemTime(t); // 1回目のforceからちょうど31分後
      mockFetchOk(makeCompactPayload({ rows: makeRows(10) }));
      const third = await syncAll(env, { force: true });
      expect(third.ok).toBe(true);
      expect(state.snapshot.rows).toBe(10);
    } finally {
      vi.useRealTimers();
    }
  });
});

describe('syncAll（5回目レビュー修正6: Worker側のGAS取得タイムアウト）', () => {
  it('GASへの取得が常に失敗するとき、fetchは2回だけ試行される（60秒×2に短縮。従来は20秒×3）', async () => {
    let calls = 0;
    global.fetch = vi.fn(async () => { calls++; throw new Error('network down'); });
    const { db } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    const out = await syncAll(env);

    expect(out.ok).toBe(false);
    expect(calls).toBe(2);
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
      expect(second.message).toMatch(/より新しい.*取得結果/);
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

describe('syncAll（3回目レビュー修正4: 世代ガードの同着対策＝同一ミリ秒の競合で旧内容が新内容を上書きしない）', () => {
  it('ロックがフェイルオープンし、2つの取得が同一のfetch_started_atで完了した場合、先に書き込めた内容を後続が上書きしない（Codexが`>=`条件で再現した競合の再現テスト）', async () => {
    // ★Codexの再現方法を模す: ロック機構自体が使えない（フェイルオープン）状態で、
    // 2つの同期の取得開始時刻(Date.now())がまったく同一のミリ秒になったケース。
    const { db, state, calls } = makeMockDB({ throwOnLockRead: true });
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };
    const nowSpy = vi.spyOn(Date, 'now');

    try {
      nowSpy.mockReturnValue(5_000_000); // 2つの取得が同じミリ秒に開始したことにする

      // 「新」＝先に完了して書き込まれる
      mockFetchOk(makeCompactPayload({ rows: makeRows(3).map((r, i) => { r[12] = 'NEW-' + i; return r; }) }));
      const first = await syncAll(env);
      expect(first.ok).toBe(true);
      expect(calls.snapshotWrites).toBe(1);
      const afterNew = state.snapshot.payload;
      expect(afterNew).toContain('NEW-0');

      // 「旧」＝同じミリ秒に始まったが、後から完了して書き込もうとする
      mockFetchOk(makeCompactPayload({ rows: makeRows(3).map((r, i) => { r[12] = 'OLD-' + i; return r; }) }));
      const second = await syncAll(env);

      // ★修正4の直接的な証拠：同一のfetch_started_atでの2回目の書き込みは
      // `>`条件により無条件で弾かれる（`>=`だった旧条件では上書きできてしまっていた）。
      expect(second.ok).toBe(true);
      expect(second.skipped).toBe(true);
      expect(state.snapshot.payload).toBe(afterNew);
      expect(state.snapshot.payload).toContain('NEW-0');
      expect(state.snapshot.payload).not.toContain('OLD-0');
      expect(calls.snapshotWrites).toBe(1); // 旧の書き込みは実際には行われていない
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
