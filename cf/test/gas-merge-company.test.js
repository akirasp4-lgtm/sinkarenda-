// 元請の統一を「会社ごと」に限定できているかを、偽のスプレッドシートの上で実際に動かす。
//
// ★なぜ必要か（2026-08-31 Codexレビュー P1）:
//   従来の merge_genba は会社の列を見ずに全行を書き換えていた。
//   実データで確認したところ、
//     「不動産」   … グローライズ 1件 / GRHD 21件
//     「オリエンス」… グローライズ 2件 / 和信カインド 114件
//     「ラーテル」  … グローライズ 24件 / 和信カインド 13件
//   グローライズの画面で「1件」と確認して実行すると、GRHDの21件まで
//   書き換わっていた。「和信カインド・ラーテル・GRHDは触らない」という
//   このVaultの決まりに正面から反する。
//
//   取り消せない操作なので、机上の確認では足りない。実際にシートを読み書きさせて
//   「他社の行が1行も変わっていないこと」を確かめる。
import { describe, it, expect, beforeEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const CODE = readFileSync(join(here, '..', '..', 'gas.js'), 'utf8');
const EXPORT = ';globalThis.__m = { mergeGenbaForCompany_, mergeGenba_ };';

function makeSheet(rows) {
  const data = rows.map((r) => r.slice());
  const sheet = {
    _data: data,
    getDataRange: () => ({ getValues: () => data.map((r) => r.slice()) }),
    getRange: (row, col) => ({
      setValue: (v) => { data[row - 1][col - 1] = v; }
    }),
    deleteRow: (row) => { data.splice(row - 1, 1); },
    // ★中身を入れ替える。data を差し替えるのではなく、その場で書き換えること。
    //   sheet._data = [...] と書いても、上の closure が見ているのは元の配列のまま
    //   ＝テストの用意が本体に届かない（実際にこれで空振りした）。
    setRows: (rows) => { data.length = 0; rows.forEach((r) => data.push(r.slice())); }
  };
  return sheet;
}

const NIPPO_HEADERS = ['登録日時', '作業日', '元請名', '現場名', '氏名', '会社'];
const nippo = (genba, company, name) => ['', '2026-08-01', genba, '現場', name || 'A', company];

function makeSS(sheets) {
  return { getSheetByName: (n) => sheets[n] || null };
}

let M;
let SHEETS;
beforeEach(() => {
  const sandbox = vm.createContext({ console, String, Number, Object, Array, JSON, Date });
  sandbox.globalThis = sandbox;
  vm.runInContext(CODE + EXPORT, sandbox, { filename: 'gas.js' });
  M = sandbox.__m;
});

function build() {
  const nippoSheet = makeSheet([
    NIPPO_HEADERS,
    nippo('不動産', 'グローライズ', 'G1'),
    nippo('不動産', 'GRHD', 'H1'),
    nippo('不動産', 'GRHD', 'H2'),
    nippo('オリエンス', '和信カインド', 'W1'),
    nippo('きんでん東', 'グローライズ', 'G2')
  ]);
  const archiveSheet = makeSheet([
    NIPPO_HEADERS,
    nippo('不動産', 'グローライズ', 'G3'),
    nippo('不動産', '和信カインド', 'W2')
  ]);
  const jobsiteSheet = makeSheet([
    ['元請名', '現場名'],
    ['不動産', 'ある現場']
  ]);
  const genbaSheet = makeSheet([
    ['名前', '会社', '読み'],
    ['不動産', 'グローライズ', ''],
    ['不動産', 'GRHD', ''],
    ['GR不動産', 'グローライズ', '']
  ]);
  SHEETS = { nippoSheet, archiveSheet, jobsiteSheet, genbaSheet };
  // gas.js の定数名に合わせて差し込む
  const names = {};
  names[M.__SHEET_NAME || '日報データ'] = nippoSheet;
  return { nippoSheet, archiveSheet, jobsiteSheet, genbaSheet };
}

// gas.js のシート名定数を読む（決め打ちしない）
function sheetNames(sandboxCode) {
  const pick = (k) => {
    const m = sandboxCode.match(new RegExp('const\\s+' + k + "\\s*=\\s*'([^']+)'"));
    return m ? m[1] : null;
  };
  return {
    SHEET_NAME: pick('SHEET_NAME'),
    ARCHIVE_SHEET: pick('ARCHIVE_SHEET'),
    JOBSITE_SHEET: pick('JOBSITE_SHEET'),
    GENBA_MASTER_SHEET: pick('GENBA_MASTER_SHEET')
  };
}

const N = sheetNames(CODE);

function ss() {
  const s = build();
  const map = {};
  map[N.SHEET_NAME] = s.nippoSheet;
  map[N.ARCHIVE_SHEET] = s.archiveSheet;
  map[N.JOBSITE_SHEET] = s.jobsiteSheet;
  map[N.GENBA_MASTER_SHEET] = s.genbaSheet;
  return { ss: makeSS(map), sheets: s };
}

const genbaOf = (sheet) => sheet._data.slice(1).map((r) => r[2]);
const companyOf = (sheet) => sheet._data.slice(1).map((r) => r[5]);

describe('gas.js のシート名定数が読めている（テストの前提）', () => {
  it('4つとも見つかる', () => {
    expect(N.SHEET_NAME).toBeTruthy();
    expect(N.ARCHIVE_SHEET).toBeTruthy();
    expect(N.JOBSITE_SHEET).toBeTruthy();
    expect(N.GENBA_MASTER_SHEET).toBeTruthy();
  });
});

describe('会社を限定した元請の統一（mergeGenbaForCompany_）', () => {
  it('★指定した会社の予定だけ書き換える', () => {
    const { ss: s, sheets } = ss();
    const r = M.mergeGenbaForCompany_(s, '不動産', 'GR不動産', 'グローライズ');
    expect(r.nippoUpdated).toBe(1);
    expect(genbaOf(sheets.nippoSheet)).toEqual(
      ['GR不動産', '不動産', '不動産', 'オリエンス', 'きんでん東']);
  });

  it('★他社（GRHD・和信カインド）の行は1行も変えない', () => {
    const { ss: s, sheets } = ss();
    M.mergeGenbaForCompany_(s, '不動産', 'GR不動産', 'グローライズ');
    const rows = sheets.nippoSheet._data.slice(1);
    rows.forEach((row) => {
      if (row[5] !== 'グローライズ') {
        expect(row[2], '他社の行が書き換わっている').not.toBe('GR不動産');
      }
    });
    expect(sheets.archiveSheet._data[2][2]).toBe('不動産');   // 和信カインドのアーカイブ
  });

  it('★他社が同じ名前を使っていた件数を返す（何を触らなかったか分かるように）', () => {
    const { ss: s } = ss();
    const r = M.mergeGenbaForCompany_(s, '不動産', 'GR不動産', 'グローライズ');
    // 日報のGRHD 2件 + アーカイブの和信カインド 1件
    expect(r.skippedOtherCompanies).toBe(3);
  });

  it('アーカイブも同じ会社のぶんだけ書き換える', () => {
    const { ss: s, sheets } = ss();
    const r = M.mergeGenbaForCompany_(s, '不動産', 'GR不動産', 'グローライズ');
    expect(r.archiveUpdated).toBe(1);
    expect(sheets.archiveSheet._data[1][2]).toBe('GR不動産');
  });

  it('★他社も使っている元請では現場マスタを触らない（会社の列が無く分けられないため）', () => {
    const { ss: s, sheets } = ss();
    const r = M.mergeGenbaForCompany_(s, '不動産', 'GR不動産', 'グローライズ');
    expect(sheets.jobsiteSheet._data[1][0]).toBe('不動産');
    expect(r.jobsiteUpdated).toBe(0);
    expect(r.masterAction).toBe('jobsite_skipped_shared');
  });

  it('その会社しか使っていない元請なら現場マスタも直す', () => {
    const { ss: s, sheets } = ss();
    sheets.jobsiteSheet.setRows([['元請名', '現場名'], ['きんでん東', 'ある現場']]);
    const r = M.mergeGenbaForCompany_(s, 'きんでん東', 'きんでん', 'グローライズ');
    expect(r.jobsiteUpdated).toBe(1);
    expect(sheets.jobsiteSheet._data[1][0]).toBe('きんでん');
  });

  it('★元請マスタは (名前, 会社) の組で見る（他社の登録を消さない）', () => {
    const { ss: s, sheets } = ss();
    M.mergeGenbaForCompany_(s, '不動産', 'GR不動産', 'グローライズ');
    const rows = sheets.genbaSheet._data.slice(1);
    const grhd = rows.filter((r) => r[1] === 'GRHD');
    expect(grhd.length, 'GRHDの登録が消えている').toBe(1);
    expect(grhd[0][0]).toBe('不動産');
  });

  it('★自社の登録が無いとき、他社の登録を巻き込んで消さない', () => {
    // ★守りを外して赤くなるか試したとき、これが無くて素通りした。
    //   自社の登録行が先に見つかる並びだと、他社除外を外しても結果が同じで気付けない。
    //   「自社には登録が無く、他社にだけある」形にすると差が出る。
    const { ss: s, sheets } = ss();
    sheets.genbaSheet.setRows([
      ['名前', '会社', '読み'],
      ['不動産', 'GRHD', ''],          // 他社の登録だけがある
      ['GR不動産', 'グローライズ', '']
    ]);
    M.mergeGenbaForCompany_(s, '不動産', 'GR不動産', 'グローライズ');
    const grhd = sheets.genbaSheet._data.slice(1).filter((r) => r[1] === 'GRHD');
    expect(grhd.length, 'GRHDの登録が消された').toBe(1);
    expect(grhd[0][0], 'GRHDの登録が書き換えられた').toBe('不動産');
  });

  it('★統一先が他社にしか無い場合、それを「もうある」と誤解しない', () => {
    const { ss: s, sheets } = ss();
    sheets.genbaSheet.setRows([
      ['名前', '会社', '読み'],
      ['不動産', 'グローライズ', ''],
      ['GR不動産', 'GRHD', '']          // 統一先は他社にだけ登録がある
    ]);
    M.mergeGenbaForCompany_(s, '不動産', 'GR不動産', 'グローライズ');
    const gl = sheets.genbaSheet._data.slice(1).filter((r) => r[1] === 'グローライズ');
    // 消すのではなく、自社の行を改名するのが正しい（消すと自社に登録が無くなる）
    expect(gl.length, 'グローライズの登録が消えた').toBe(1);
    expect(gl[0][0]).toBe('GR不動産');
  });

  it('統一先が同じ会社に既にあれば、やめる方の登録を消す', () => {
    const { ss: s, sheets } = ss();
    const before = sheets.genbaSheet._data.length;
    M.mergeGenbaForCompany_(s, '不動産', 'GR不動産', 'グローライズ');
    expect(sheets.genbaSheet._data.length).toBe(before - 1);
  });

  it('会社を指定しない従来の関数は今までどおり全社を書き換える（既存機能の回帰確認）', () => {
    const { ss: s } = ss();
    const r = M.mergeGenba_(s, '不動産', 'GR不動産');
    expect(r.nippoUpdated).toBe(3);   // グローライズ1 + GRHD2
  });
});
