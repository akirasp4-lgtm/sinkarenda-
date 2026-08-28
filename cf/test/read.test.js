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


// ============================================================
// 拠点（本社／関東支店）での絞り込み — 2026-08-26
//   依頼書 calendar_request_20260826.md ②③
// ============================================================
const H20 = [...HEADERS, '拠点'];
function makeRow20(fields) {
  const row = new Array(20).fill('');
  for (const [h, v] of Object.entries(fields)) row[H20.indexOf(h)] = v;
  return row;
}
function payload20(rows) {
  return { headers: H20, rows, members: [], genbaMaster: [], jobsites: [] };
}

describe('filterSnapshot（拠点での絞り込み）', () => {
  const rows = [
    makeRow20({ 会社: 'グローライズ', 現場名: '本社の現場', 拠点: '本社' }),
    makeRow20({ 会社: 'GRミツマ',     現場名: '関東の現場', 拠点: '関東支店' }),
    makeRow20({ 会社: 'グローライズ', 現場名: '合同の現場', 拠点: '両方' })
  ];

  it('拠点を指定しなければ全部返る（全社ビュー）', () => {
    const out = filterSnapshot(payload20(rows), '', '');
    expect(out.rows).toHaveLength(3);
  });

  it('★「両方」は本社ビューにも関東ビューにも出る（1件登録で両方に出す＝二重登録の廃止）', () => {
    const honsha = filterSnapshot(payload20(rows), '', '本社');
    const kanto  = filterSnapshot(payload20(rows), '', '関東支店');
    const loc = o => o.rows.map(r => r[H20.indexOf('現場名')]);
    expect(loc(honsha)).toContain('合同の現場');
    expect(loc(kanto)).toContain('合同の現場');
  });

  it('★他拠点の予定が1件も混ざらない', () => {
    const honsha = filterSnapshot(payload20(rows), '', '本社');
    expect(honsha.rows.map(r => r[H20.indexOf('現場名')])).not.toContain('関東の現場');
    const kanto = filterSnapshot(payload20(rows), '', '関東支店');
    expect(kanto.rows.map(r => r[H20.indexOf('現場名')])).not.toContain('本社の現場');
  });

  it('拠点が空欄の行は「本社」として扱う（過去データ埋めの既定値と揃える）', () => {
    const legacy = [makeRow20({ 会社: 'グローライズ', 現場名: '昔の現場', 拠点: '' })];
    expect(filterSnapshot(payload20(legacy), '', '本社').rows).toHaveLength(1);
    expect(filterSnapshot(payload20(legacy), '', '関東支店').rows).toHaveLength(0);
  });

  // ★★利用者指定（2026-08-26）: 関東はミツマとグローライズだけの話。
  //   ラーテル・和信カインドを混ぜてはいけない。
  it('★和信カインド・ラーテル・GRHDの予定は、本社ビューにも関東ビューにも出ない', () => {
    const others = [
      makeRow20({ 会社: '和信カインド', 現場名: '和信の現場', 拠点: '' }),
      makeRow20({ 会社: 'ラーテル',     現場名: 'ラーテルの現場', 拠点: '' }),
      makeRow20({ 会社: 'GRHD',        現場名: 'GRHDの現場', 拠点: '' })
    ];
    expect(filterSnapshot(payload20(others), '', '本社').rows).toHaveLength(0);
    expect(filterSnapshot(payload20(others), '', '関東支店').rows).toHaveLength(0);
    // 拠点で絞らなければ（＝これまでどおりの会社別の見方）ちゃんと出る
    expect(filterSnapshot(payload20(others), '', '').rows).toHaveLength(3);
    expect(filterSnapshot(payload20(others), '和信カインド', '').rows).toHaveLength(1);
  });

  it('★他社の行に拠点が誤って入っていても、拠点ビューには出さない（混ざり防止）', () => {
    const stray = [makeRow20({ 会社: '和信カインド', 現場名: '誤って本社と入った行', 拠点: '本社' })];
    expect(filterSnapshot(payload20(stray), '', '本社').rows).toHaveLength(0);
  });

  it('★会社の絞り込みと拠点の絞り込みは両立する（法人と拠点は別の軸）', () => {
    // GRミツマ法人だが本社案件、という行が正しく扱えること（依頼書の要件）
    const mixed = [
      makeRow20({ 会社: 'GRミツマ', 現場名: 'ミツマだが本社案件', 拠点: '本社' }),
      makeRow20({ 会社: 'GRミツマ', 現場名: 'ミツマの関東案件', 拠点: '関東支店' })
    ];
    const out = filterSnapshot(payload20(mixed), 'GRミツマ', '本社');
    expect(out.rows).toHaveLength(1);
    expect(out.rows[0][H20.indexOf('現場名')]).toBe('ミツマだが本社案件');
  });

  it('19列のまま（拠点列が無い古い取り込み）でも落ちず、全部返る', () => {
    const old = { headers: HEADERS, rows: [makeRow({ 会社: 'グローライズ' })],
                  members: [], genbaMaster: [], jobsites: [] };
    const out = filterSnapshot(old, '', '本社');
    expect(out.status).toBe('ok');
    expect(out.rows).toHaveLength(1);
  });
});

// ★2026-08-26 Codexレビュー[P1]#6#7 / [P2]#12 の再発防止
describe('filterSnapshot（Codexレビュー指摘）', () => {
  it('★「全拠点」（画面が使う語）でも絞り込まない — 画面とWorkerの語を揃える', () => {
    const rows = [
      makeRow20({ 会社: 'グローライズ', 現場名: 'A', 拠点: '本社' }),
      makeRow20({ 会社: 'GRミツマ',     現場名: 'B', 拠点: '関東支店' })
    ];
    expect(filterSnapshot(payload20(rows), '', '全拠点').rows).toHaveLength(2);
    expect(filterSnapshot(payload20(rows), '', '全社').rows).toHaveLength(2);   // 旧語も許容
  });

  it('★知らない拠点の値が来たら、混ぜずに0件にする（誤って全件返さない）', () => {
    const rows = [makeRow20({ 会社: 'グローライズ', 現場名: 'A', 拠点: '本社' })];
    expect(filterSnapshot(payload20(rows), '', '関西支店').rows).toHaveLength(0);
  });

  it('★拠点列がまだ無い取り込み（19列）でも、拠点で絞れば他事業の会社は返さない', () => {
    // 移行途中: D1が19列のまま。それでも「本社」で絞ったら和信カインドは出さない
    const old19 = {
      headers: HEADERS,
      rows: [ makeRow({ 会社: 'グローライズ' }), makeRow({ 会社: '和信カインド' }), makeRow({ 会社: 'ラーテル' }) ],
      members: [], genbaMaster: [], jobsites: []
    };
    const out = filterSnapshot(old19, '', '本社');
    expect(out.rows).toHaveLength(1);
    expect(out.rows[HEADERS.indexOf('会社')]).toBeUndefined();
    expect(out.rows[0][HEADERS.indexOf('会社')]).toBe('グローライズ');
    // 関東で絞れば、拠点列が無い＝全部「本社」扱いなので0件
    expect(filterSnapshot(old19, '', '関東支店').rows).toHaveLength(0);
  });
});

// ============================================================
// 資格（2026-08-28 追加）
// ============================================================
describe('資格の会社ごとの絞り込み', () => {
  const quals = [
    { name: '真柄', company: 'グローライズ', qual: '玉掛け', kind: '技能講習', expires: '' },
    { name: '誰か', company: '和信カインド', qual: 'フォークリフト', kind: '技能講習', expires: '' }
  ];
  it('会社を指定すると他社の資格は1件も混ざらない', () => {
    const out = filterSnapshot({ headers: ['日付', '会社'], rows: [], members: [], genbaMaster: [], jobsites: [], qualifications: quals }, '和信カインド', '');
    expect(out.qualifications).toEqual([quals[1]]);
  });
  it('全社なら全部返す', () => {
    const out = filterSnapshot({ headers: ['日付', '会社'], rows: [], members: [], genbaMaster: [], jobsites: [], qualifications: quals }, '全社', '');
    expect(out.qualifications).toEqual(quals);
  });
  it('★古いスナップショット（qualifications が無い）でも落ちない', () => {
    const out = filterSnapshot({ headers: ['日付', '会社'], rows: [], members: [], genbaMaster: [], jobsites: [] }, 'グローライズ', '');
    expect(out.qualifications).toEqual([]);
  });
});

describe('資格：グローライズとGRミツマは1つの名簿', () => {
  // ★資格マスタには統合前に取り込んだ 会社=GRミツマ の行が26件残っている。
  //   単純一致で絞ると、江頭さん・繁田さんの資格がグローライズの画面から消える。
  const quals = [
    { name: '江頭', company: 'GRミツマ', qual: '玉掛け', kind: '技能講習', expires: '' },
    { name: '真柄', company: 'グローライズ', qual: '高所作業車', kind: '技能講習', expires: '' },
    { name: '誰か', company: '和信カインド', qual: 'フォークリフト', kind: '技能講習', expires: '' }
  ];
  const pay = { headers: ['日付', '会社'], rows: [], members: [], genbaMaster: [], jobsites: [], qualifications: quals };

  it('★グローライズを指定してもGRミツマの人の資格が消えない', () => {
    const out = filterSnapshot(pay, 'グローライズ', '');
    expect(out.qualifications.map(q => q.name).sort()).toEqual(['江頭', '真柄']);
  });
  it('★GRミツマを指定してもグローライズの人の資格が出る（同じ名簿なので）', () => {
    const out = filterSnapshot(pay, 'GRミツマ', '');
    expect(out.qualifications.map(q => q.name).sort()).toEqual(['江頭', '真柄']);
  });
  it('★和信カインドにはグローライズ系が1件も混ざらない', () => {
    const out = filterSnapshot(pay, '和信カインド', '');
    expect(out.qualifications.map(q => q.name)).toEqual(['誰か']);
  });
});
