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

  it('会社名セルの前後の空白は格納前に落とす（D1のWHERE kaisha=?は完全一致のため）', () => {
    const row = [...new Array(19).fill('')];
    row[1] = '2026-05-02';        // 作業日
    row[4] = '川端（達）';          // 氏名
    row[8] = 1;                   // 人工
    row[11] = '  グローライズ　'; // 会社（前後に半角・全角空白が紛れ込んだ想定）
    row[12] = 'abc-1';            // ID
    const json = { status:'ok', compact:1, headers:HEADERS, rows:[row],
                   members:[], genbaMaster:[], jobsites:[] };
    const out = parseGasPayload(json);
    expect(out.nippo).toHaveLength(1);
    expect(out.nippo[0].kaisha).toBe('グローライズ');
  });

  it('会社名セルが空・未定義でも落ちない（trim対象がnull/undefinedでも例外にならない）', () => {
    const row = new Array(19).fill('');
    row[1] = '2026-05-02'; row[4] = '森'; row[12] = 'abc-2';
    row[11] = undefined; // 会社が未定義
    const json = { status:'ok', compact:1, headers:HEADERS, rows:[row],
                   members:[], genbaMaster:[], jobsites:[] };
    const out = parseGasPayload(json);
    expect(out.nippo[0].kaisha).toBe('');
  });

  it('compactでない応答は受け付けない（形が変わると壊れるため明示的に落とす）', () => {
    expect(() => parseGasPayload({status:'ok', rows:[{'ID':'x'}]}))
      .toThrow(/compact/);
  });

  // ★2026-08-24 設計変更：以前はここで「IDが空の行は捨てる」としていたが、
  // これが本番データ突き合わせで14行欠落（車検期限リマインダー行の消失）を
  // 引き起こした。D1はGAS応答の忠実な写しにする方針へ変更したため、
  // IDが空でも行は捨てない（下の「行を捨てるフィルタの回帰防止」を参照）。
  it('IDが空の行も捨てない（連番seqが主キーなので、IDが無くても行を持てる）', () => {
    const row = new Array(19).fill('');
    row[1] = '2026-05-02'; row[4] = '森';
    const json = { status:'ok', compact:1, headers:HEADERS, rows:[row],
                   members:[], genbaMaster:[], jobsites:[] };
    const out = parseGasPayload(json);
    expect(out.nippo).toHaveLength(1);
    expect(out.nippo[0].id).toBe('');
  });
});

// --- ここから追加（2026-08-24 設計変更：本番データ突き合わせで見つかった
// 14行欠落＝車検期限リマインダー行の消失の再発防止）。
//
// 原因は parseGasPayload の `if (!rec.sagyoubi || !rec.shimei) continue;` 等、
// 行を捨てるフィルタだった。GASのdoGetはスプレッドシートの行をそのまま
// 返すだけで、「氏名が空の行（車両の車検期限だけを置いてある行）」も
// 「同名・同会社の職人が複数登録されている（元データの重複）」も正当に
// 存在する。D1はGAS応答の忠実な写しでなければならず、整形・掃除は
// 移行とは別の作業（利用者の裁定）。
describe('parseGasPayload（行を捨てるフィルタの回帰防止 = 2026-08-24 本番突き合わせで発覚した14行欠落）', () => {
  it('氏名が空の行が捨てられずD1へ入ること（車検期限リマインダー行の再現）', () => {
    // 実例: ID VKEN_なにわ432そ8800 / 作業日 2026-10-16 /
    // 現場名「ハイエース白 車検期限」/ メモ「車検期限日 なにわ432そ8800」/
    // 夜勤「予定」/ 更新者「車検期限登録」/ 氏名は空（人が出る予定ではないため）。
    const row = new Array(19).fill('');
    row[1] = '2026-10-16';                       // 作業日
    row[3] = 'ハイエース白 車検期限';               // 現場名
    row[4] = '';                                  // 氏名（正当に空）
    row[9] = '車検期限日 なにわ432そ8800';          // メモ
    row[10] = '予定';                              // 夜勤
    row[11] = 'グローライズ';                       // 会社
    row[12] = 'VKEN_なにわ432そ8800';               // ID
    row[13] = '車検期限登録';                       // 更新者
    const json = { status:'ok', compact:1, headers:HEADERS, rows:[row],
                   members:[], genbaMaster:[], jobsites:[] };

    const out = parseGasPayload(json);
    expect(out.nippo).toHaveLength(1);
    expect(out.nippo[0]).toMatchObject({
      id: 'VKEN_なにわ432そ8800', sagyoubi: '2026-10-16', genba: 'ハイエース白 車検期限',
      shimei: '', memo: '車検期限日 なにわ432そ8800', yakin: '予定', koushinsha: '車検期限登録'
    });
  });

  it('作業日が空の行も捨てられないこと（IDと氏名だけの行）', () => {
    const row = new Array(19).fill('');
    row[4] = '森'; row[12] = 'id-no-date';
    const json = { status:'ok', compact:1, headers:HEADERS, rows:[row],
                   members:[], genbaMaster:[], jobsites:[] };
    expect(parseGasPayload(json).nippo).toHaveLength(1);
  });

  it('同名・同会社の職人が3件あったら3件とも保持されること（元データの重複を畳まない）', () => {
    const json = {
      status: 'ok', compact: 1, headers: HEADERS, rows: [],
      members: [
        { name: '林電工(りんたろう)', company: '和信カインド', division: '電気', rate: 15000 },
        { name: '林電工(りんたろう)', company: '和信カインド', division: '電気', rate: 15000 },
        { name: '林電工(りんたろう)', company: '和信カインド', division: '電気', rate: 15000 }
      ],
      genbaMaster: [], jobsites: []
    };
    const out = parseGasPayload(json);
    expect(out.members).toHaveLength(3);
    for (const m of out.members) {
      expect(m).toEqual({ name: '林電工(りんたろう)', company: '和信カインド', division: '電気' });
    }
  });

  it('並び順が入力順のまま保たれること（重複排除・並び替えをしない）', () => {
    const rows = [];
    for (const id of ['c-3', 'a-1', 'b-2', 'a-1']) {  // 意図的に非ソート順＋重複ID
      const row = new Array(19).fill('');
      row[1] = '2026-05-02'; row[4] = '森'; row[12] = id;
      rows.push(row);
    }
    const json = { status:'ok', compact:1, headers:HEADERS, rows,
                   members:[], genbaMaster:[], jobsites:[] };
    const out = parseGasPayload(json);
    expect(out.nippo.map(r => r.id)).toEqual(['c-3', 'a-1', 'b-2', 'a-1']);
  });

  it('会社名(元請)・現場名が空の genba/jobsites 行も捨てられないこと', () => {
    const json = {
      status: 'ok', compact: 1, headers: HEADERS, rows: [],
      members: [], genbaMaster: [{ name: '', company: '' }],
      jobsites: [{ genba: '', loc: '', jobNo: '', completed: 0, billingMethod: '' }]
    };
    const out = parseGasPayload(json);
    expect(out.genba).toHaveLength(1);
    expect(out.jobsites).toHaveLength(1);
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
function makeMockDB({ failOnBatchCall, failSyncLogWrite, failOnPrepareCall } = {}) {
  const batchCalls = [];      // batch()に渡された文の配列を、呼び出しごとに記録
  const syncLogRuns = [];     // sync_logへのINSERTで渡された引数を記録
  let prepareCallCount = 0;

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
      prepareCallCount++;
      // GASが想定外の形を返し、文の組み立て中（bindより前のprepare自体）
      // で例外が出るケースを模す。
      if (failOnPrepareCall && prepareCallCount === failOnPrepareCall) {
        throw new Error('mock prepare failure on call ' + prepareCallCount);
      }
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

  it('env.DB.prepareが例外を投げても、syncAllは例外を投げずにok:falseを返し、sync_logにok=0が記録されること', async () => {
    // 文の組み立て中（batch投入より前）に例外が出るケース。
    // parseGasPayloadはnippoの値を素通しするため、GASが想定外の形
    // （配列やオブジェクト等）を返すとprepare/bindが例外を投げうる。
    // この区間が保護されていないと、syncAllの外へ例外がそのまま漏れる。
    const rows = [makeRow(1)];
    global.fetch = vi.fn(async () => ({
      ok: true, status: 200,
      json: async () => ({
        status: 'ok', compact: 1, headers: HEADERS, rows,
        members: [], genbaMaster: [], jobsites: [],
      }),
    }));

    // 文の組み立てで最初に呼ばれるprepare（DELETE FROM nippo）で失敗させる。
    const { db, syncLogRuns } = makeMockDB({ failOnPrepareCall: 1 });
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    const out = await syncAll(env);

    expect(out.ok).toBe(false);
    expect(out.rows).toBe(0);
    expect(out.message).toMatch(/mock prepare failure/);

    expect(syncLogRuns).toHaveLength(1);
    const [, , ok, message] = syncLogRuns[0];
    expect(ok).toBe(0);
    expect(message).toMatch(/mock prepare failure/);
  });
});
