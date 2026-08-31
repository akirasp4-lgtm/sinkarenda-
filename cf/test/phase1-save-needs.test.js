// Phase 1: 必要人員条件の保存を、偽のスプレッドシートの上で実際に動かす。
//
// ★なぜ偽シートまで作るか:
//   現場マスタは列番号べた書きで読み書きしている。「13列目に書いたつもりが12列目だった」
//   という取り違えは、机上の検査（ソースの文字列を見るだけ）では絶対に捕まらない。
//   実際に読み書きさせて、**既存12列が1セルも変わっていないこと**を確かめる。
//
// ★社長指示で守るべきこと:
//   ・空欄は「条件未登録」。勝手な推測値で埋めない
//   ・変更履歴で追跡可能にする
//   ・既存データを勝手に変更しない
import { describe, it, expect, beforeEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const CODE = readFileSync(join(here, '..', '..', 'gas.js'), 'utf8');

// ---------------------------------------------------------------- 偽のシート
function makeSheet(rows) {
  const data = rows.map((r) => r.slice());
  return {
    _data: data,
    _writes: [],                                   // どのセルに書いたか（取り違えの検出用）
    getDataRange: () => ({ getValues: () => data.map((r) => r.slice()) }),
    getRange: (row, col, nr, nc) => ({
      setValue: (v) => { data[row - 1][col - 1] = v; },
      setValues: (vv) => { vv.forEach((r, k) => { data[row - 1 + k] = r.slice(); }); },
      getValues: () => [data[row - 1].slice(col - 1, col - 1 + (nc || 1))]
    }),
    appendRow: (r) => { data.push(r.slice()); },
    getMaxColumns: () => Math.max(...data.map((r) => r.length), 1),
    getLastRow: () => data.length,
    getLastColumn: () => Math.max(...data.map((r) => r.length), 1),
    insertColumnsAfter: () => {},
    getName: () => 'sheet'
  };
}

const JOB_HEADERS = ['元請名', '現場名', '工番', '事業部', '年度', '連番', '売上', '読み',
  '完了', '請求方式', '拠点', 'ステータス', '必要人数', '必要資格', '必要経験',
  '現場住所', '開始時間', '終了時間'];

// 既存の現場1件（12列ぶん埋まっていて、13〜18は空）
const SITE = ['きんでん東', 'A現場', 'ICT-26-001', 'ICT', 2026, 1, 500000, 'えーげんば',
  '', '応援', '本社', '施工中', '', '', '', '', '', ''];

const fakeLock = () => ({ tryLock: () => true, waitLock: () => {}, releaseLock: () => {}, hasLock: () => true });

let G, sheets;

beforeEach(() => {
  const jobsite = makeSheet([JOB_HEADERS.slice(), SITE.slice()]);
  const history = makeSheet([['日時', '操作', '旧ID', '新ID', '項目', '変更前', '変更後', '実行者']]);
  const oplog = makeSheet([['日時', '操作', '対象', '内容', '実行者']]);
  sheets = { jobsite, history, oplog };

  const box = vm.createContext({
    console, String, Number, Object, Array, Math, isFinite, JSON, Date, RegExp,
    SpreadsheetApp: { flush: () => {}, getActiveSpreadsheet: () => ss },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} }) },
    // 本物は用途ごとに3種類のロックを使い分けている。どれか1つでも欠けると
    // 「is not a function」で落ちるので、全部そろえておく。
    LockService: {
      getScriptLock: () => fakeLock(),
      getUserLock: () => fakeLock(),
      getDocumentLock: () => fakeLock()
    },
    ContentService: {
      MimeType: { JSON: 'json' },
      createTextOutput: (t) => ({ setMimeType: () => ({ _t: t }), _t: t })
    },
    Utilities: { formatDate: () => '' }
  });
  box.globalThis = box;

  // ★doPost は入口で日報シートのロックを取りに行くので、無いと落ちる。
  //   「現場マスタしか使わないから他は要らない」は通らない。
  const others = {};
  const ss = {
    getSheetByName: (n) => {
      if (n === '現場マスタ') return jobsite;
      if (n === '変更履歴') return history;
      if (n === '操作ログ') return oplog;
      if (!others[n]) others[n] = makeSheet([[]]);   // それ以外は空のシートを返す
      return others[n];
    },
    insertSheet: (n) => { others[n] = makeSheet([[]]); return others[n]; }
  };

  vm.runInContext(CODE + ';globalThis.__g = { doPost, getOrCreateJobSiteSheet_ };', box, { filename: 'gas.js' });
  G = box.__g;
});

// doPost をJSON本文で呼ぶ近道
function post(body) {
  const res = G.doPost({ postData: { contents: JSON.stringify(body) } });
  return JSON.parse(res._t);
}

const row = () => sheets.jobsite._data[1];

describe('必要人員条件の保存（set_site_needs）', () => {
  it('★13〜18列目に書き、既存12列は1セルも変えない', () => {
    const before = row().slice(0, 12);
    const r = post({
      action: 'set_site_needs', genba: 'きんでん東', loc: 'A現場',
      needCount: 6, needQuals: ['高所作業認定', 'フルハーネス型墜落制止用器具'],
      needExp: '楽天案件経験', address: '大阪市北区1-1', startAt: '08:00', endAt: '17:00',
      updatedBy: 'テスト'
    });
    expect(r.status, '保存が失敗: ' + (r.message || '')).toBe('ok');
    expect(row().slice(0, 12), '★既存12列が書き換わった').toEqual(before);
    expect(row()[12]).toBe(6);                                       // 必要人数
    expect(row()[13]).toBe('高所作業認定、フルハーネス型墜落制止用器具'); // 必要資格
    expect(row()[14]).toBe('楽天案件経験');
    expect(row()[15]).toBe('大阪市北区1-1');
    expect(row()[16]).toBe('08:00');
    expect(row()[17]).toBe('17:00');
  });

  it('★空欄を送ると「未登録」に戻せる（入力ミスを直せないと困る）', () => {
    post({ action: 'set_site_needs', genba: 'きんでん東', loc: 'A現場', needCount: 6, updatedBy: 'テスト' });
    expect(row()[12]).toBe(6);
    post({ action: 'set_site_needs', genba: 'きんでん東', loc: 'A現場', needCount: '', updatedBy: 'テスト' });
    expect(row()[12], '未登録に戻せない').toBe('');
  });

  it('★送っていない項目は触らない（勝手に消さない）', () => {
    post({
      action: 'set_site_needs', genba: 'きんでん東', loc: 'A現場',
      needCount: 6, address: '大阪市北区1-1', updatedBy: 'テスト'
    });
    // 住所だけ送り直す。必要人数はそのまま残るはず
    post({ action: 'set_site_needs', genba: 'きんでん東', loc: 'A現場', address: '京都市', updatedBy: 'テスト' });
    expect(row()[12], '送っていない必要人数が消えた').toBe(6);
    expect(row()[15]).toBe('京都市');
  });

  it('★0人は「未登録」として保存する（0人必要ではない）', () => {
    post({ action: 'set_site_needs', genba: 'きんでん東', loc: 'A現場', needCount: 0, updatedBy: 'テスト' });
    expect(row()[12]).toBe('');
  });

  it('おかしい時刻は空欄で保存する（推測で直さない）', () => {
    post({ action: 'set_site_needs', genba: 'きんでん東', loc: 'A現場', startAt: '25:00', updatedBy: 'テスト' });
    expect(row()[16]).toBe('');
  });

  it('必要資格は文字列でも配列でも受ける', () => {
    post({ action: 'set_site_needs', genba: 'きんでん東', loc: 'A現場', needQuals: '玉掛け、高所作業認定', updatedBy: 'テスト' });
    expect(row()[13]).toBe('玉掛け、高所作業認定');
  });

  it('★資格名の前後の空白は落とす（資格マスタと1文字も違ってはいけない）', () => {
    post({ action: 'set_site_needs', genba: 'きんでん東', loc: 'A現場', needQuals: [' 玉掛け ', '', ' 高所作業認定'], updatedBy: 'テスト' });
    expect(row()[13]).toBe('玉掛け、高所作業認定');
  });

  it('元請名が無ければ拒否する', () => {
    const r = post({ action: 'set_site_needs', loc: 'A現場', needCount: 3, updatedBy: 'テスト' });
    expect(r.status).toBe('error');
  });

  it('現場マスタに無い現場は拒否する（勝手に行を作らない）', () => {
    const n = sheets.jobsite._data.length;
    const r = post({ action: 'set_site_needs', genba: '存在しない元請', loc: 'X', needCount: 3, updatedBy: 'テスト' });
    expect(r.status).toBe('error');
    expect(sheets.jobsite._data.length, '行が増えた').toBe(n);
  });

  it('保存する項目が1つも無ければ拒否する', () => {
    const r = post({ action: 'set_site_needs', genba: 'きんでん東', loc: 'A現場', updatedBy: 'テスト' });
    expect(r.status).toBe('error');
  });
});

describe('変更履歴（社長指示14「追跡可能にする」）', () => {
  it('★変えた項目ごとに履歴が残る', () => {
    post({
      action: 'set_site_needs', genba: 'きんでん東', loc: 'A現場',
      needCount: 6, address: '大阪市北区1-1', updatedBy: '向井'
    });
    const rows = sheets.history._data.slice(1);
    expect(rows.length, '履歴が残っていない').toBe(2);
    const fields = rows.map((r) => String(r[4]));
    expect(fields.some((f) => f.includes('必要人数'))).toBe(true);
    expect(fields.some((f) => f.includes('現場住所'))).toBe(true);
    expect(rows[0][1]).toBe('site_needs');
    expect(rows[0][7]).toBe('向井');           // 実行者
  });

  it('★変更前が空欄なら「(未登録)」と残す（空欄と0の区別が後から分かるように）', () => {
    post({ action: 'set_site_needs', genba: 'きんでん東', loc: 'A現場', needCount: 6, updatedBy: 'テスト' });
    const r = sheets.history._data[1];
    expect(String(r[5])).toBe('(未登録)');
    expect(String(r[6])).toBe('6');
  });

  it('中身が変わっていなければ履歴を増やさない（同じ値の連打で汚れない）', () => {
    post({ action: 'set_site_needs', genba: 'きんでん東', loc: 'A現場', needCount: 6, updatedBy: 'テスト' });
    const n = sheets.history._data.length;
    const r = post({ action: 'set_site_needs', genba: 'きんでん東', loc: 'A現場', needCount: 6, updatedBy: 'テスト' });
    expect(r.changed).toBe(0);
    expect(sheets.history._data.length).toBe(n);
  });

  it('★履歴が書けなければ現場マスタを1セルも触らない', () => {
    // 履歴シートへの書き込みを壊す。
    // ★logHistory_ は appendRow ではなく getRange().setValues() で書く（gas.js:2301）。
    //   壊す場所を実装に合わせないと、テストが素通りして「守れているつもり」になる。
    sheets.history.getRange = () => ({
      setValues: () => { throw new Error('履歴シートが壊れている'); },
      setValue: () => { throw new Error('履歴シートが壊れている'); }
    });
    const before = row().slice();
    const r = post({ action: 'set_site_needs', genba: 'きんでん東', loc: 'A現場', needCount: 6, updatedBy: 'テスト' });
    expect(r.status).toBe('error');
    expect(row(), '履歴が書けないのに原本を書き換えた').toEqual(before);
  });
});
