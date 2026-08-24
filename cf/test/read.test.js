import { describe, it, expect } from 'vitest';
import { HEADERS, filterSnapshot, readSchedule } from '../src/read.js';

// ============================================================
// filterSnapshot（gas.js の doGet と完全に同じ絞り込み条件であること）
// ============================================================
function makePayload(overrides = {}) {
  return {
    headers: HEADERS,
    rows: [],
    members: [],
    genbaMaster: [],
    jobsites: [],
    ...overrides
  };
}

function makeRow(fields) {
  const row = new Array(19).fill('');
  for (const [h, v] of Object.entries(fields)) row[HEADERS.indexOf(h)] = v;
  return row;
}

describe('filterSnapshot（HEADERS）', () => {
  it('GASと同じ19個のヘッダを同じ順で持つ', () => {
    expect(HEADERS).toEqual(['登録日時','作業日','元請名','現場名','氏名','役割','出勤','退勤',
      '人工','メモ','夜勤','会社','ID','更新者','色','事業部','工番','作業区分','車両']);
  });
});

describe('filterSnapshot（company未指定・全社は絞り込みなし）', () => {
  it('companyが空文字なら全件そのまま返る', () => {
    const payload = makePayload({
      rows: [makeRow({ 会社: 'グローライズ' }), makeRow({ 会社: '和信カインド' })],
      members: [{ name: '森', company: 'グローライズ', division: '' }, { name: '田中', company: '和信カインド', division: '' }],
      genbaMaster: [{ name: 'A現場', company: 'グローライズ' }, { name: 'B現場', company: '' }],
      jobsites: [{ genba: 'A現場', loc: 'x', jobNo: '', completed: true, billingMethod: '応援' }]
    });
    const out = filterSnapshot(payload, '');
    expect(out.status).toBe('ok');
    expect(out.compact).toBe(1);
    expect(out.rows).toHaveLength(2);
    expect(out.members).toHaveLength(2);
    expect(out.genbaMaster).toHaveLength(2);
    expect(out.jobsites).toHaveLength(1);
  });

  it('companyが「全社」でも絞り込みなし扱い（gas.jsと同じ）', () => {
    const payload = makePayload({
      rows: [makeRow({ 会社: 'グローライズ' })],
      members: [{ name: '森', company: 'グローライズ', division: '' }]
    });
    const out = filterSnapshot(payload, '全社');
    expect(out.rows).toHaveLength(1);
    expect(out.members).toHaveLength(1);
  });
});

describe('filterSnapshot（日報rowsの会社絞り込み）', () => {
  it('会社セルをtrimしてから比較する（会社名に前後空白が紛れても一致させる）', () => {
    const payload = makePayload({
      rows: [makeRow({ 会社: '  グローライズ　', ID: 'a-1' }), makeRow({ 会社: '和信カインド', ID: 'b-1' })]
    });
    const out = filterSnapshot(payload, 'グローライズ');
    expect(out.rows).toHaveLength(1);
    expect(out.rows[0][HEADERS.indexOf('ID')]).toBe('a-1');
  });

  it('一致しない会社は除外される', () => {
    const payload = makePayload({ rows: [makeRow({ 会社: '和信カインド' })] });
    const out = filterSnapshot(payload, 'グローライズ');
    expect(out.rows).toHaveLength(0);
  });
});

describe('filterSnapshot（membersの会社絞り込み = gas.js:1240 と同条件）', () => {
  it('会社の完全一致のみで絞り込む（genbaMasterと違い「会社が空なら通す」例外は無い）', () => {
    const payload = makePayload({
      members: [
        { name: '森', company: 'グローライズ', division: '' },
        { name: '空会社太郎', company: '', division: '' }
      ]
    });
    const out = filterSnapshot(payload, 'グローライズ');
    expect(out.members).toHaveLength(1);
    expect(out.members[0].name).toBe('森');
  });

  it('職人マスタに単価は含まれない（元々sanitizeForStorageで除去済みの前提。ここでは型を変えないことを確認）', () => {
    const payload = makePayload({ members: [{ name: '森', company: 'GRHD', division: 'ICT' }] });
    const out = filterSnapshot(payload, '');
    expect(out.members[0]).toEqual({ name: '森', company: 'GRHD', division: 'ICT' });
    expect('rate' in out.members[0]).toBe(false);
  });
});

describe('filterSnapshot（genbaMasterの絞り込み = gas.js:1244 と同条件）', () => {
  it('name が空の行は絞り込みの有無に関わらず常に除外する', () => {
    const payload = makePayload({ genbaMaster: [{ name: '', company: '' }, { name: 'A現場', company: '' }] });
    expect(filterSnapshot(payload, '').genbaMaster).toHaveLength(1);
    expect(filterSnapshot(payload, 'グローライズ').genbaMaster).toHaveLength(1);
  });

  it('絞り込み時、companyが空（共通元請）なら通す', () => {
    const payload = makePayload({ genbaMaster: [{ name: '共通現場', company: '' }] });
    const out = filterSnapshot(payload, 'グローライズ');
    expect(out.genbaMaster).toHaveLength(1);
  });

  it('絞り込み時、companyが一致すれば通す・不一致なら除外する', () => {
    const payload = makePayload({
      genbaMaster: [{ name: 'G現場', company: 'グローライズ' }, { name: 'W現場', company: '和信カインド' }]
    });
    const out = filterSnapshot(payload, 'グローライズ');
    expect(out.genbaMaster.map(g => g.name)).toEqual(['G現場']);
  });
});

describe('filterSnapshot（jobsitesの絞り込み = gas.js:1256 と同条件）', () => {
  it('genba が空の行は常に除外する', () => {
    const payload = makePayload({ jobsites: [{ genba: '', loc: '', jobNo: '', completed: false, billingMethod: '' }] });
    expect(filterSnapshot(payload, '').jobsites).toHaveLength(0);
  });

  it('絞り込み時は、絞り込み後のgenbaMasterに含まれるgenbaのjobsitesだけ通す', () => {
    const payload = makePayload({
      genbaMaster: [{ name: 'G現場', company: 'グローライズ' }, { name: 'W現場', company: '和信カインド' }],
      jobsites: [
        { genba: 'G現場', loc: 'a', jobNo: '', completed: false, billingMethod: '応援' },
        { genba: 'W現場', loc: 'b', jobNo: '', completed: false, billingMethod: '応援' }
      ]
    });
    const out = filterSnapshot(payload, 'グローライズ');
    expect(out.jobsites.map(j => j.genba)).toEqual(['G現場']);
  });

  it('completed は真偽値のまま返る（画面が真偽値で判定するため）', () => {
    const payload = makePayload({
      genbaMaster: [{ name: 'A', company: '' }],
      jobsites: [{ genba: 'A', loc: 'b', jobNo: '', completed: true, billingMethod: '応援' }]
    });
    const out = filterSnapshot(payload, '');
    expect(out.jobsites[0].completed).toBe(true);
  });
});

// ============================================================
// readSchedule（D1アクセスを含む結線）
// ============================================================
// ★修正1（再レビュー対応・鮮度ガード）: readScheduleはsnapshotの存在だけでなく
// sync_logの直近の成功(ok=1)時刻も確認するようになった。そのためモックは
// snapshotPayloadに加えてsyncLog（[{at, ok}, ...]）も受け取れるようにする。
// デフォルトの `freshSuccess: true` は「たった今成功した」ログを1件自動的に
// 用意する（＝素朴に「snapshotがあればok」だった頃と同じ結果になる）ので、
// 既存の正常系テストは鮮度ガードを意識せず書けるままにしてある。
function makeMockDB({ snapshotPayload = null, syncLog = null, freshSuccess = true } = {}) {
  const log = syncLog != null
    ? syncLog
    : (freshSuccess ? [{ at: new Date().toISOString(), ok: 1, message: '' }] : []);
  const db = {
    prepare(sql) {
      return {
        all: async () => {
          if (/SELECT payload FROM snapshot/.test(sql)) {
            return { results: snapshotPayload != null ? [{ payload: snapshotPayload }] : [] };
          }
          if (/SELECT at FROM sync_log WHERE ok = 1/.test(sql)) {
            const oks = log.filter(l => Number(l.ok) === 1).sort((a, b) => b.at.localeCompare(a.at));
            return { results: oks.length ? [{ at: oks[0].at }] : [] };
          }
          return { results: [] };
        }
      };
    }
  };
  return db;
}

describe('readSchedule（snapshotが無い/壊れている場合の安全装置）', () => {
  it('snapshotが1行も無い（まだ一度も取り込みが成功していない）ときはエラーを返す', async () => {
    const env = { DB: makeMockDB({ snapshotPayload: null }) };
    const out = await readSchedule(env, '');
    expect(out.status).toBe('error');
    expect(out.rows).toBeUndefined();
  });

  it('保存済みpayloadのJSONが壊れていてもクラッシュせずエラーを返す', async () => {
    const env = { DB: makeMockDB({ snapshotPayload: '{not valid json' }) };
    const out = await readSchedule(env, '');
    expect(out.status).toBe('error');
  });
});

describe('readSchedule（修正1・再レビュー: 鮮度ガード）', () => {
  it('同期が失敗し続けていて（sync_logに直近の成功が無い）snapshotだけが残っている場合はstatus:errorを返す（Codexの再現ケース）', async () => {
    // 1,500,414バイトで同期失敗した直後：sync_logはok=0だけ、snapshotは前回成功時点のまま残っている、
    // という状況を模す。以前はここでstatus:'ok'を返してしまっていた（レビュー指摘）。
    const payload = JSON.stringify(makePayload({ rows: [makeRow({ ID: 'STORED' })] }));
    const env = { DB: makeMockDB({
      snapshotPayload: payload,
      syncLog: [
        { at: '2026-08-24T00:00:00.000Z', ok: 1, message: '' }, // 大昔の成功（15分以上前）
        { at: new Date().toISOString(), ok: 0, message: '件数が急減しました：...' } // たった今の失敗
      ]
    }) };
    const out = await readSchedule(env, '');
    expect(out.status).toBe('error');
    expect(out.rows).toBeUndefined();
  });

  it('直近の成功が15分以内なら新しいデータとしてstatus:okを返す', async () => {
    const payload = JSON.stringify(makePayload({ rows: [makeRow({ ID: 'a-1' })] }));
    const fiveMinAgo = new Date(Date.now() - 5 * 60 * 1000).toISOString();
    const env = { DB: makeMockDB({
      snapshotPayload: payload,
      syncLog: [{ at: fiveMinAgo, ok: 1, message: '' }]
    }) };
    const out = await readSchedule(env, '');
    expect(out.status).toBe('ok');
  });

  it('直近の成功が15分より古ければstatus:errorを返す（同期が長時間失敗し続けている想定）', async () => {
    const payload = JSON.stringify(makePayload({ rows: [makeRow({ ID: 'old-data' })] }));
    const twentyMinAgo = new Date(Date.now() - 20 * 60 * 1000).toISOString();
    const env = { DB: makeMockDB({
      snapshotPayload: payload,
      syncLog: [{ at: twentyMinAgo, ok: 1, message: '' }]
    }) };
    const out = await readSchedule(env, '');
    expect(out.status).toBe('error');
    expect(out.message).toMatch(/同期|成功/);
  });

  it('sync_logが1件も無ければ（snapshotだけ存在する想定外の状態）status:errorを返す', async () => {
    const payload = JSON.stringify(makePayload({ rows: [] }));
    const env = { DB: makeMockDB({ snapshotPayload: payload, syncLog: [] }) };
    const out = await readSchedule(env, '');
    expect(out.status).toBe('error');
  });

  it('ハッシュ一致でスキップされた同期でも「成功」としてsync_logの時刻が更新されるため、鮮度は新しいと判定される（変更が無いだけを古いと誤判定しない）', async () => {
    const payload = JSON.stringify(makePayload({ rows: [makeRow({ ID: 'unchanged' })] }));
    const env = { DB: makeMockDB({
      snapshotPayload: payload,
      syncLog: [
        { at: '2026-08-24T00:00:00.000Z', ok: 1, message: '' }, // 最初に書き込んだ時刻（古い）
        { at: new Date(Date.now() - 60 * 1000).toISOString(), ok: 1, message: '変更なし（書き込みをスキップしました）' } // 1分前に「変更なし」を確認
      ]
    }) };
    const out = await readSchedule(env, '');
    expect(out.status).toBe('ok');
  });
});

describe('readSchedule（正常系）', () => {
  it('保存済みsnapshotをJSON.parseし、gas.jsと同じ形で返す', async () => {
    const payload = JSON.stringify(makePayload({
      rows: [makeRow({ ID: 'abc-1', 会社: 'グローライズ' })],
      members: [{ name: '森', company: 'グローライズ', division: 'ICT' }],
      genbaMaster: [{ name: '大阪', company: '' }],
      jobsites: [{ genba: '大阪', loc: '本社', jobNo: '', completed: false, billingMethod: '応援' }]
    }));
    const env = { DB: makeMockDB({ snapshotPayload: payload }) };
    const out = await readSchedule(env, '');

    expect(out.status).toBe('ok');
    expect(out.compact).toBe(1);
    expect(out.rows).toHaveLength(1);
    expect(out.rows[0][HEADERS.indexOf('ID')]).toBe('abc-1');
    expect(out.members).toEqual([{ name: '森', company: 'グローライズ', division: 'ICT' }]);
    expect(out.jobsites[0].completed).toBe(false);
  });

  it('company指定で絞り込まれる', async () => {
    const payload = JSON.stringify(makePayload({
      rows: [makeRow({ ID: 'a-1', 会社: 'グローライズ' }), makeRow({ ID: 'b-1', 会社: '和信カインド' })],
      members: [
        { name: '森', company: 'グローライズ', division: 'ICT' },
        { name: '田中', company: '和信カインド', division: '' }
      ]
    }));
    const env = { DB: makeMockDB({ snapshotPayload: payload }) };
    const out = await readSchedule(env, 'グローライズ');

    expect(out.status).toBe('ok');
    expect(out.rows).toHaveLength(1);
    expect(out.members).toHaveLength(1);
    expect(out.members[0].name).toBe('森');
  });
});
