// GRミツマを「グローライズ 関東支店」へ寄せる処理を、偽スプレッドシート上で実際に動かす。
//
// ★利用者指示（2026-08-28）:
//   「和信、ラーテル、GRHDは触らないで下さい。
//     また、GRミツマは関東支店としてグローライズに統合されるので消して下さい」
//
// ★会社欄を書き換えても情報が失われない根拠（実測）:
//   日報データの GRミツマ 85件は全件 拠点=関東支店。グローライズ 2,285件は全件 本社。
//   「どちらの拠点か」は既に拠点列が持っている。
import { describe, it, expect, beforeEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const CODE = readFileSync(join(here, '..', '..', 'gas.js'), 'utf8');

const EXPORT = `
;globalThis.__gas = { mergeMitsumaIntoGrowise, planMitsumaRows_, mergeMemberRows_, MITSUMA };
`;

const H21 = ['登録日時','作業日','元請名','現場名','氏名','役割','出勤','退勤','人工','メモ',
             '夜勤','会社','ID','更新者','色','事業部','工番','作業区分','車両','拠点','部隊'];

function makeSheet(data) {
  return {
    _data: data,
    getName() { return this._name; },
    getDataRange() { const d = this._data; return { getValues: () => d.map(r => r.slice()) }; },
    getLastRow() { return this._data.length; },
    getRange(row, col, numRows, numCols) {
      const self = this;
      return {
        setValues(vals) {
          for (let i = 0; i < vals.length; i++) {
            const r = row - 1 + i;
            while (self._data.length <= r) self._data.push([]);
            for (let j = 0; j < vals[i].length; j++) self._data[r][col - 1 + j] = vals[i][j];
          }
        },
        clearContent() {
          for (let i = 0; i < numRows; i++) {
            const r = row - 1 + i;
            if (!self._data[r]) continue;
            for (let j = 0; j < (numCols || 1); j++) self._data[r][col - 1 + j] = '';
          }
        }
      };
    }
  };
}

const row = (name, company, kyoten, loc) => H21.map(h =>
  h === '氏名' ? name : h === '会社' ? company : h === '拠点' ? kyoten
  : h === '現場名' ? (loc || 'A現場') : h === '作業日' ? '2026-09-01'
  : h === '元請名' ? 'きんでん東' : h === '作業区分' ? '現場作業' : '');

function build(overrides) {
  const sheets = Object.assign({
    '日報データ': makeSheet([H21,
      row('中島', 'グローライズ', '本社'),
      row('高田（関東）', 'GRミツマ', '関東支店'),
      row('柳澤（関東）', 'GRミツマ', '関東支店'),
      row('元', '和信カインド', ''),           // ★触らない
      row('いくや', 'ラーテル', ''),            // ★触らない
      row('奥田', 'GRHD', '')                   // ★触らない
    ]),
    'アーカイブ': makeSheet([H21,
      row('内村（関東）', 'GRミツマ', '関東支店'),
      row('東', 'グローライズ', '本社')
    ]),
    '職人マスタ': makeSheet([
      ['', '会社', '事業部', '単価', '既定部隊', '有効'],
      ['中島', 'グローライズ', 'ICT', 28000, '第一部隊', '○'],
      ['江頭', 'グローライズ', 'ICT', 0, '', '○'],
      ['江頭', 'GRミツマ', '', 0, '', '○'],          // ★二重登録
      ['繁田', 'グローライズ', 'GRB', 0, '', '○'],
      ['繁田', 'GRミツマ', '', 0, '', '○'],          // ★二重登録
      ['高田（関東）', 'GRミツマ', 'GRM', 20000, '', '○'],
      ['元', '和信カインド', '', 0, '', '○'],        // ★触らない
      ['奥田', 'GRHD', '', 0, '', '○']               // ★触らない
    ]),
    '元請マスタ': makeSheet([
      ['元請名', '会社', '読み'],
      ['きんでん西', 'グローライズ', 'きんでんにし'],
      ['きんでん東', 'GRミツマ', 'きんでんひがし'],
      ['ラーテル', 'ラーテル', 'らーてる']
    ]),
    '現場マスタ': makeSheet([['元請名','現場名','工番','事業部','年度','連番','売上','読み','完了','請求方式','拠点','ステータス']]),
    '操作ログ': makeSheet([['日時','操作','対象','詳細','実行者']]),
    '変更履歴': makeSheet([['日時','操作','旧ID','新ID','項目','変更前','変更後','実行者']])
  }, overrides || {});
  Object.keys(sheets).forEach(n => { sheets[n]._name = n; });
  const ss = { getSheetByName: (n) => sheets[n] || null, insertSheet: (n) => (sheets[n] = makeSheet([[]])) };
  const sandbox = vm.createContext({
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush() {} },
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    Utilities: { formatDate: (d, tz, f) => String(f) },
    ContentService: {}, UrlFetchApp: {},
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} }) },
    Logger: { log() {} }, console
  });
  vm.runInContext(CODE + EXPORT, sandbox, { filename: 'gas.js' });
  return { g: sandbox.__gas, sheets };
}

const col = (sheet, h) => {
  const d = sheet._data;
  let i = d[0].indexOf(h);
  if (i < 0 && h === '氏名') i = 0;
  if (i < 0) throw new Error('見出しが見つからない: ' + h);
  return d.slice(1).map(r => String(r[i] == null ? '' : r[i]));
};

let ctx;
beforeEach(() => { ctx = build(); });

describe('GRミツマ → グローライズ 関東支店（会社欄の統合）', () => {
  it('dry-run は1文字も書かない', () => {
    const before = JSON.stringify(ctx.sheets);
    const rep = ctx.g.mergeMitsumaIntoGrowise(false);
    expect(rep.dryRun).toBe(true);
    expect(JSON.stringify(ctx.sheets)).toBe(before);
  });

  it('dry-run が変更の件数を返す', () => {
    const rep = ctx.g.mergeMitsumaIntoGrowise(false);
    const nippo = rep.シート.find(s => s.sheet === '日報データ');
    expect(nippo.会社の変更).toBe(2);
    expect(rep.シート.find(s => s.sheet === 'アーカイブ').会社の変更).toBe(1);
    expect(rep.中止理由).toBe('');
  });

  it('★日報データの会社が グローライズ になる', () => {
    ctx.g.mergeMitsumaIntoGrowise(true);
    expect(col(ctx.sheets['日報データ'], '会社'))
      .toEqual(['グローライズ', 'グローライズ', 'グローライズ', '和信カインド', 'ラーテル', 'GRHD']);
  });

  it('★拠点は書き換えない（関東支店のまま）', () => {
    ctx.g.mergeMitsumaIntoGrowise(true);
    expect(col(ctx.sheets['日報データ'], '拠点'))
      .toEqual(['本社', '関東支店', '関東支店', '', '', '']);
  });

  it('★拠点が空のGRミツマ行だけ 関東支店 を補う（空だと画面が本社として扱うため）', () => {
    const c = build({ '日報データ': Object.assign(makeSheet([H21,
      row('高田（関東）', 'GRミツマ', ''),
      row('中島', 'グローライズ', '本社')
    ]), { _name: '日報データ' }) });
    const rep = c.g.mergeMitsumaIntoGrowise(true);
    expect(rep.シート.find(s => s.sheet === '日報データ').拠点を補った).toBe(1);
    expect(col(c.sheets['日報データ'], '拠点')).toEqual(['関東支店', '本社']);
  });

  it('★グローライズの拠点（本社）は1つも変えない', () => {
    const before = col(ctx.sheets['日報データ'], '拠点')[0];
    ctx.g.mergeMitsumaIntoGrowise(true);
    expect(col(ctx.sheets['日報データ'], '拠点')[0]).toBe(before);
  });

  it('★和信カインド・ラーテル・GRHD は1文字も触らない（利用者指示）', () => {
    ctx.g.mergeMitsumaIntoGrowise(true);
    const co = col(ctx.sheets['日報データ'], '会社');
    expect(co.slice(3)).toEqual(['和信カインド', 'ラーテル', 'GRHD']);
    const m = ctx.sheets['職人マスタ']._data.slice(1).filter(r => r[0] === '元' || r[0] === '奥田');
    expect(m.map(r => r[1]).sort()).toEqual(['GRHD', '和信カインド']);
  });

  it('行数は1行も増減しない（日報データ・アーカイブ）', () => {
    const a = ctx.sheets['日報データ']._data.length, b = ctx.sheets['アーカイブ']._data.length;
    ctx.g.mergeMitsumaIntoGrowise(true);
    expect(ctx.sheets['日報データ']._data.length).toBe(a);
    expect(ctx.sheets['アーカイブ']._data.length).toBe(b);
  });
});

describe('職人マスタ', () => {
  it('★GRミツマの人がグローライズへ移り、二重登録が1人にまとまる', () => {
    const rep = ctx.g.mergeMitsumaIntoGrowise(true);
    expect(rep.職人マスタ.前).toBe(8);
    expect(rep.職人マスタ.後).toBe(6);   // 江頭・繁田 が1人ずつになる
    const names = col(ctx.sheets['職人マスタ'], '氏名').filter(Boolean);
    expect(names.filter(n => n === '江頭').length).toBe(1);
    expect(names.filter(n => n === '繁田').length).toBe(1);
  });

  it('★まとめても事業部は消えない（空の側に引きずられない）', () => {
    ctx.g.mergeMitsumaIntoGrowise(true);
    const d = ctx.sheets['職人マスタ']._data.slice(1);
    expect((d.find(r => r[0] === '江頭') || [])[2]).toBe('ICT');
    expect((d.find(r => r[0] === '繁田') || [])[2]).toBe('GRB');
  });

  it('関東の人の事業部・単価はそのまま残る', () => {
    ctx.g.mergeMitsumaIntoGrowise(true);
    const t = ctx.sheets['職人マスタ']._data.slice(1).find(r => r[0] === '高田（関東）');
    expect(t[1]).toBe('グローライズ');
    expect(t[2]).toBe('GRM');
    expect(t[3]).toBe(20000);
  });

  it('★値が食い違うときは1文字も書かずに中止する', () => {
    const c = build({ '職人マスタ': Object.assign(makeSheet([
      ['', '会社', '事業部', '単価', '既定部隊', '有効'],
      ['江頭', 'グローライズ', 'ICT', 30000, '', '○'],
      ['江頭', 'GRミツマ', 'INF', 25000, '', '○']   // 事業部も単価も食い違う
    ]), { _name: '職人マスタ' }) });
    const beforeN = JSON.stringify(c.sheets['日報データ']._data);
    const beforeM = JSON.stringify(c.sheets['職人マスタ']._data);
    const rep = c.g.mergeMitsumaIntoGrowise(true);
    expect(rep.食い違い.length).toBeGreaterThan(0);
    expect(rep.中止理由).toContain('食い違い');
    expect(JSON.stringify(c.sheets['日報データ']._data)).toBe(beforeN);
    expect(JSON.stringify(c.sheets['職人マスタ']._data)).toBe(beforeM);
  });
});

describe('元請マスタ', () => {
  it('会社欄だけ グローライズ になる', () => {
    ctx.g.mergeMitsumaIntoGrowise(true);
    expect(col(ctx.sheets['元請マスタ'], '会社'))
      .toEqual(['グローライズ', 'グローライズ', 'ラーテル']);
  });

  it('★同じ（元請名・会社）になる行は1行にまとめる（読みが同じなら中止しない）', () => {
    const c = build({ '元請マスタ': Object.assign(makeSheet([
      ['元請名', '会社', '読み'],
      ['きんでん東', 'グローライズ', 'ひがし'],
      ['きんでん東', 'GRミツマ', 'ひがし']   // 寄せると同じ組になる
    ]), { _name: '元請マスタ' }) });
    const rep = c.g.mergeMitsumaIntoGrowise(true);
    expect(rep.中止理由).toBe('');
    const rows = c.sheets['元請マスタ']._data.slice(1).filter(r => r[0]);
    expect(rows.length).toBe(1);
    expect(rows[0]).toEqual(['きんでん東', 'グローライズ', 'ひがし']);
  });
});

describe('もう一度動かしても何も起きない（べき等）', () => {
  it('2回目は変更0件', () => {
    ctx.g.mergeMitsumaIntoGrowise(true);
    const after = JSON.stringify(ctx.sheets);
    const rep = ctx.g.mergeMitsumaIntoGrowise(true);
    expect(rep.シート.reduce((a, s) => a + s.会社の変更, 0)).toBe(0);
    expect(JSON.stringify(ctx.sheets)).toBe(after);
  });
});

describe('元請名の寄せ（2026-08-28 利用者判断「ミツマは関東支店になるんだよ」）', () => {
  const withJisha = () => build({
    '日報データ': Object.assign(makeSheet([H21,
      (function () { const r = row('高田（関東）', 'GRミツマ', '関東支店'); r[H21.indexOf('元請名')] = 'GRミツマ自社'; return r; })(),
      (function () { const r = row('中島', 'グローライズ', '本社'); r[H21.indexOf('元請名')] = 'グローライズ自社'; return r; })(),
      (function () { const r = row('元', '和信カインド', ''); r[H21.indexOf('元請名')] = 'GRミツマ自社'; return r; })()  // ★他社の行は触らない
    ]), { _name: '日報データ' })
  });

  it('★元請名「GRミツマ自社」が「グローライズ自社」になる', () => {
    const c = withJisha();
    c.g.mergeMitsumaIntoGrowise(true);
    expect(col(c.sheets['日報データ'], '元請名')[0]).toBe('グローライズ自社');
  });

  it('★和信カインドの行の元請名は触らない（利用者指示）', () => {
    const c = withJisha();
    c.g.mergeMitsumaIntoGrowise(true);
    expect(col(c.sheets['日報データ'], '元請名')[2]).toBe('GRミツマ自社');
    expect(col(c.sheets['日報データ'], '会社')[2]).toBe('和信カインド');
  });

  it('元請マスタの「GRミツマ自社」も「グローライズ自社」になる', () => {
    const c = build({
      '元請マスタ': Object.assign(makeSheet([
        ['元請名', '会社', '読み'],
        ['GRミツマ自社', 'GRミツマ', 'みつまじしゃ']
      ]), { _name: '元請マスタ' })
    });
    c.g.mergeMitsumaIntoGrowise(true);
    expect(c.sheets['元請マスタ']._data[1][0]).toBe('グローライズ自社');
    expect(c.sheets['元請マスタ']._data[1][1]).toBe('グローライズ');
  });
});

describe('元請マスタの重複を1行にまとめる', () => {
  it('★同じ元請名が両社にあれば1行になる', () => {
    const c = build({
      '元請マスタ': Object.assign(makeSheet([
        ['元請名', '会社', '読み'],
        ['児玉通信', 'グローライズ', ''],
        ['児玉通信', 'GRミツマ', 'こだまつうしん']
      ]), { _name: '元請マスタ' })
    });
    const rep = c.g.mergeMitsumaIntoGrowise(true);
    expect(rep.中止理由).toBe('');
    const rows = c.sheets['元請マスタ']._data.slice(1).filter(r => r[0]);
    expect(rows.length).toBe(1);
    expect(rows[0]).toEqual(['児玉通信', 'グローライズ', 'こだまつうしん']);
  });

  it('★読みが食い違うときは、利用者が決めた読みを使う（きんでん東＝きんでんひがし）', () => {
    const c = build({
      '元請マスタ': Object.assign(makeSheet([
        ['元請名', '会社', '読み'],
        ['きんでん東', 'グローライズ', 'きんでんとう'],
        ['きんでん東', 'GRミツマ', 'きんでんひがし']
      ]), { _name: '元請マスタ' })
    });
    const rep = c.g.mergeMitsumaIntoGrowise(true);
    expect(rep.中止理由).toBe('');
    const rows = c.sheets['元請マスタ']._data.slice(1).filter(r => r[0]);
    expect(rows.length).toBe(1);
    expect(rows[0][2]).toBe('きんでんひがし');
  });

  it('★決めていない読みの食い違いは1文字も書かずに中止する', () => {
    const c = build({
      '元請マスタ': Object.assign(makeSheet([
        ['元請名', '会社', '読み'],
        ['ハイテックス', 'グローライズ', 'はいてっくす'],
        ['ハイテックス', 'GRミツマ', 'ハイテクス']
      ]), { _name: '元請マスタ' })
    });
    const before = JSON.stringify(c.sheets['日報データ']._data);
    const rep = c.g.mergeMitsumaIntoGrowise(true);
    expect(rep.中止理由).toContain('読み');
    expect(JSON.stringify(c.sheets['日報データ']._data)).toBe(before);
  });

  it('他社（和信カインド）の同名の元請は別のまま', () => {
    const c = build({
      '元請マスタ': Object.assign(makeSheet([
        ['元請名', '会社', '読み'],
        ['オリエンス', 'グローライズ', ''],
        ['オリエンス', '和信カインド', '']
      ]), { _name: '元請マスタ' })
    });
    const rep = c.g.mergeMitsumaIntoGrowise(true);
    expect(rep.中止理由).toBe('');
    expect(c.sheets['元請マスタ']._data.slice(1).filter(r => r[0]).length).toBe(2);
  });
});
