import { describe, it, expect, vi } from 'vitest';
import { parseGasPayload, fetchWithRetry, syncAll } from '../src/sync.js';

const HEADERS = ['登録日時','作業日','元請名','現場名','氏名','役割','出勤','退勤','人工',
                 'メモ','夜勤','会社','ID','更新者','色','事業部','工番','作業区分','車両'];

describe('parseGasPayload', () => {
  it('compact形式の1行をD1の列名へ移し替える', () => {
    const json = {
      status: 'ok', compact: 1, headers: HEADERS,
      rows: [['2026-05-01T04:23:04.000Z','2026-05-02','NGS','大阪','川端（達）','代表',
              '09:00','18:00',1,'','','グローライズ','abc-1','森','#1D9E75','ICT',
              'INF-26-041','現場作業','']],
      members: [], genbaMaster: [], jobsites: []
    };
    const out = parseGasPayload(json);
    expect(out.nippo).toHaveLength(1);
    expect(out.nippo[0]).toMatchObject({
      id: 'abc-1', sagyoubi: '2026-05-02', motoukr: 'NGS', genba: '大阪',
      shimei: '川端（達）', kosu: 1, kaisha: 'グローライズ', kouban: 'INF-26-041'
    });
  });

  it('職人マスタから単価(rate)を落とす', () => {
    const json = { status:'ok', compact:1, headers:HEADERS, rows:[],
      members:[{name:'森',company:'GRHD',division:'ICT',rate:18000}],
      genbaMaster:[], jobsites:[] };
    const out = parseGasPayload(json);
    expect(out.members[0]).toEqual({name:'森',company:'GRHD',division:'ICT'});
    expect(JSON.stringify(out.members)).not.toContain('18000');
  });

  it('compactでない応答は受け付けない（形が変わると壊れるため明示的に落とす）', () => {
    expect(() => parseGasPayload({status:'ok', rows:[{'ID':'x'}]}))
      .toThrow(/compact/);
  });

  it('IDが空の行は捨てる（主キーにできないため）', () => {
    const row = new Array(19).fill('');
    row[1] = '2026-05-02'; row[4] = '森';
    const json = { status:'ok', compact:1, headers:HEADERS, rows:[row],
                   members:[], genbaMaster:[], jobsites:[] };
    expect(parseGasPayload(json).nippo).toHaveLength(0);
  });
});

describe('fetchWithRetry', () => {
  it('1回目が404でも2回目で成功すれば結果を返す（GASは5回に1回404を返す）', async () => {
    const calls = [];
    global.fetch = vi.fn(async (u) => {
      calls.push(u);
      if (calls.length === 1) return { ok:false, status:404, text: async () => '<html>' };
      return { ok:true, status:200, json: async () => ({status:'ok'}) };
    });
    const out = await fetchWithRetry('https://example.test/', 3);
    expect(out).toEqual({status:'ok'});
    expect(calls).toHaveLength(2);
  });

  it('回数を使い切ったら投げる', async () => {
    global.fetch = vi.fn(async () => ({ ok:false, status:404, text: async () => '<html>' }));
    await expect(fetchWithRetry('https://example.test/', 2)).rejects.toThrow(/404/);
  });
});

// --- ここから追加（計画からの変更点：500文ずつの分割投入とその失敗時の扱い） ---

// D1のprepare/bind/run/batchを模した簡易モック。
// bind()は新しいbindされた文（sql+args）を返し、batch配列にそのまま渡せる形にする。
function makeMockDB({ failOnBatchCall, failSyncLogWrite } = {}) {
  const batchCalls = [];      // batch()に渡された文の配列を、呼び出しごとに記録
  const syncLogRuns = [];     // sync_logへのINSERTで渡された引数を記録

  function makeStmt(sql, args = []) {
    return {
      sql,
      args,
      bind(...newArgs) {
        return makeStmt(sql, newArgs);
      },
      async run() {
        if (/INSERT OR REPLACE INTO sync_log/.test(sql)) {
          syncLogRuns.push(args);
          if (failSyncLogWrite) throw new Error('mock sync_log write failure');
        }
        return { success: true };
      },
    };
  }

  const db = {
    prepare(sql) {
      return makeStmt(sql, []);
    },
    async batch(stmts) {
      batchCalls.push(stmts);
      if (failOnBatchCall && batchCalls.length === failOnBatchCall) {
        throw new Error('mock batch failure on call ' + batchCalls.length);
      }
      return stmts.map(() => ({ success: true }));
    },
  };

  return { db, batchCalls, syncLogRuns };
}

// 有効な日報行を1件作る（id/sagyoubi/shimeiを一意にして主キー重複を避ける）
function makeRow(i) {
  const row = new Array(19).fill('');
  row[1] = '2026-05-02';       // 作業日
  row[4] = '作業員' + i;        // 氏名
  row[8] = 1;                  // 人工
  row[11] = 'グローライズ';      // 会社
  row[12] = 'id-' + i;         // ID
  return row;
}

describe('syncAll（500文ずつの分割投入）', () => {
  it('500件を超える行数を渡したとき、batchが複数回に分かれて呼ばれること', async () => {
    const rows = Array.from({ length: 600 }, (_, i) => makeRow(i));
    global.fetch = vi.fn(async () => ({
      ok: true, status: 200,
      json: async () => ({
        status: 'ok', compact: 1, headers: HEADERS, rows,
        members: [], genbaMaster: [], jobsites: [],
      }),
    }));

    const { db, batchCalls } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    const out = await syncAll(env);

    expect(out.ok).toBe(true);
    expect(out.rows).toBe(600);
    // 総文数 = DELETE nippo(1) + insert(600) + DELETE members/genba/jobsites(3) = 604 → 500文ずつで2回
    expect(batchCalls.length).toBeGreaterThan(1);
    for (const chunk of batchCalls) {
      expect(chunk.length).toBeLessThanOrEqual(500);
    }
  });

  it('途中のbatchが失敗したとき、例外を投げずにok:falseを返し、sync_logにok=0が記録されること', async () => {
    const rows = Array.from({ length: 600 }, (_, i) => makeRow(i));
    global.fetch = vi.fn(async () => ({
      ok: true, status: 200,
      json: async () => ({
        status: 'ok', compact: 1, headers: HEADERS, rows,
        members: [], genbaMaster: [], jobsites: [],
      }),
    }));

    // 2回目のbatch呼び出しで失敗させる（1回目のDELETE+一部INSERTは通った想定）
    const { db, syncLogRuns } = makeMockDB({ failOnBatchCall: 2 });
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    // syncAllは契約上、例外を投げない。戻り値のokで失敗を表す。
    const out = await syncAll(env);
    expect(out.ok).toBe(false);
    expect(out.rows).toBe(0);
    expect(out.message).toMatch(/mock batch failure/);

    expect(syncLogRuns).toHaveLength(1);
    const [at, loggedRows, ok, message] = syncLogRuns[0];
    expect(ok).toBe(0);
    expect(message).toMatch(/mock batch failure/);
    expect(typeof at).toBe('string');
  });

  it('sync_logへの書き込み自体が失敗しても、返すmessageには元の失敗原因が残ること', async () => {
    const rows = Array.from({ length: 600 }, (_, i) => makeRow(i));
    global.fetch = vi.fn(async () => ({
      ok: true, status: 200,
      json: async () => ({
        status: 'ok', compact: 1, headers: HEADERS, rows,
        members: [], genbaMaster: [], jobsites: [],
      }),
    }));

    // batchが2回目で失敗し、かつsync_logへの書き込み自体も失敗する状況。
    // それでもsyncAllが返すmessageは「元のbatch失敗」のままであること
    // （ログ書き込み失敗の理由にすり替わらないこと）を確認する。
    const { db, syncLogRuns } = makeMockDB({ failOnBatchCall: 2, failSyncLogWrite: true });
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    const out = await syncAll(env);

    expect(out.ok).toBe(false);
    expect(out.message).toMatch(/mock batch failure/);
    expect(out.message).not.toMatch(/mock sync_log write failure/);
    // 書き込みは試みられた（そして失敗した）ことは記録から分かる
    expect(syncLogRuns).toHaveLength(1);
  });
});
