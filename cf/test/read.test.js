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
function makeMockDB({ snapshotPayload = null } = {}) {
  const db = {
    prepare(sql) {
      return {
        all: async () => {
          if (/SELECT payload FROM snapshot/.test(sql)) {
            return { results: snapshotPayload != null ? [{ payload: snapshotPayload }] : [] };
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
