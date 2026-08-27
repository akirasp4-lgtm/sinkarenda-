// gas.js を「偽のスプレッドシート」の上で実際に動かす結合テスト。
//
// 単体テスト（gas-phase1.test.js）は純粋な関数だけを見ている。
// こちらは buildDailyValues_ のように「シートを読んで行を組み立てる」処理を、
// 本番と同じ形のデータで通す。列の位置ずれ・書き忘れはここでしか出ない。
import { describe, it, expect, beforeEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const GAS_PATH = join(here, '..', '..', 'gas.js');
const CODE = readFileSync(GAS_PATH, 'utf8');

const EXPORT = `
;globalThis.__gas = { HEADERS, buildDailyValues_, getMemberButaiMap_, lookupMemberButai_, memberKey_, diffDailyRows_, rowFullJson_ };`;

// ── 最低限のスプレッドシートの真似
function makeSheet(rows) {
  const data = rows.map(r => r.slice());
  const sheet = {
    _data: data,
    getDataRange: () => ({ getValues: () => data.map(r => r.slice()) }),
    getMaxColumns: () => Math.max(...data.map(r => r.length), 1),
    getMaxRows: () => data.length,
    getLastRow: () => data.length,
    getLastColumn: () => Math.max(...data.map(r => r.length), 1),
    insertColumnsAfter: () => {},
    insertRowsAfter: () => {},
    appendRow: (r) => data.push(r.slice()),
    getRange: (row, col, nRows, nCols) => ({
      getValues: () => {
        const out = [];
        for (let i = 0; i < (nRows || 1); i++) {
          const src = data[row - 1 + i] || [];
          out.push(src.slice(col - 1, col - 1 + (nCols || 1)));
        }
        return out;
      },
      setValue: (v) => {
        while (data.length < row) data.push([]);
        data[row - 1][col - 1] = v;
      },
      setValues: (vals) => {
        vals.forEach((rv, i) => {
          while (data.length < row + i) data.push([]);
          rv.forEach((v, j) => { data[row - 1 + i][col - 1 + j] = v; });
        });
      },
      clearContent: () => {}
    })
  };
  return sheet;
}

function makeContext(sheets) {
  const ss = {
    getSheetByName: (n) => sheets[n] || null,
    insertSheet: (n) => (sheets[n] = makeSheet([[]]))
  };
  const sandbox = vm.createContext({
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush() {} },
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    Utilities: {
      // ★fmtDate_ / fmtTime_ が使う。gas.js は起動時に tzFastOk_() でこれを試し、
      //   期待どおりの文字列が返れば以降は素のDateメソッドで組み立てる（速い経路）。
      //   ここを空にしておくと日付・時刻の変換が例外になり、履歴の突き合わせが壊れる。
      formatDate: (d, tz, fmt) => {
        const p = (n) => String(n).padStart(2, '0');
        return String(fmt)
          .replace('yyyy', d.getFullYear())
          .replace('MM', p(d.getMonth() + 1))
          .replace('dd', p(d.getDate()))
          .replace('HH', p(d.getHours()))
          .replace('mm', p(d.getMinutes()))
          .replace('ss', p(d.getSeconds()));
      }
    },
    ContentService: {}, PropertiesService: {},
    UrlFetchApp: {}, Logger: { log() {} }, console
  });
  vm.runInContext(CODE + EXPORT, sandbox, { filename: 'gas.js' });
  return { g: sandbox.__gas, ss, sheets };
}

// 本番と同じ列構成の職人マスタ・現場マスタ
const MEMBER_ROWS = [
  ['氏名', '会社', '事業部', '単価', '既定部隊', '有効'],
  ['元', 'グローライズ', 'INF', 25000, '第二部隊', '○'],
  ['中島', 'グローライズ', 'ICT', 24000, '', '○'],
  ['デモ', 'グローライズ', '', 0, '', '×']
];
const JOBSITE_ROWS = [
  ['元請名', '現場名', '工番', '事業部', '年度', '連番', '売上', '読み', '完了', '請求方式', '拠点', 'ステータス'],
  ['きんでん西', 'A現場', 'INF-26-001', 'INF', 2026, 1, 0, '', '', '現場ごと', '本社', '施工中'],
  ['きんでん西', '関東B', 'ICT-26-002', 'ICT', 2026, 2, 0, '', '✓', '現場ごと', '関東支店', '']
];

let ctx;
beforeEach(() => {
  ctx = makeContext({
    '日報データ': makeSheet([[]]),
    '職人マスタ': makeSheet(MEMBER_ROWS.map(r => r.slice())),
    '現場マスタ': makeSheet(JOBSITE_ROWS.map(r => r.slice())),
    '元請マスタ': makeSheet([['元請名', '会社', '読み']])
  });
});

describe('buildDailyValues_（実際にシートを読んで行を組み立てる）', () => {
  const baseRow = {
    date: '2026-08-28', genba: 'きんでん西', loc: 'A現場', name: '元', role: '代表',
    start: '08:00', end: '17:00', kosu: 1, memo: '', company: 'グローライズ',
    id: 'ID1', updatedBy: '向', color: '', workType: '現場作業', vehicle: ''
  };

  it('★21列ちょうど出る（列がずれていない）', () => {
    const out = ctx.g.buildDailyValues_(ctx.ss, [baseRow], '向');
    expect(out.length).toBe(1);
    expect(out[0].length).toBe(21);
    expect(ctx.g.HEADERS.length).toBe(21);
  });

  it('★既存20列の中身が正しい位置に入っている', () => {
    const out = ctx.g.buildDailyValues_(ctx.ss, [baseRow], '向')[0];
    const H = ctx.g.HEADERS;
    expect(out[H.indexOf('作業日')]).toBe('2026-08-28');
    expect(out[H.indexOf('元請名')]).toBe('きんでん西');
    expect(out[H.indexOf('現場名')]).toBe('A現場');
    expect(out[H.indexOf('氏名')]).toBe('元');
    expect(out[H.indexOf('役割')]).toBe('代表');     // ★保存値は「代表」のまま
    expect(out[H.indexOf('会社')]).toBe('グローライズ');
    expect(out[H.indexOf('ID')]).toBe('ID1');
    expect(out[H.indexOf('作業区分')]).toBe('現場作業');
  });

  it('画面が部隊を送らなければ、職人マスタの既定部隊が入る', () => {
    const out = ctx.g.buildDailyValues_(ctx.ss, [baseRow], '向')[0];
    expect(out[ctx.g.HEADERS.indexOf('部隊')]).toBe('第二部隊');
  });

  it('画面が部隊を送ってきたらそれを使う', () => {
    const out = ctx.g.buildDailyValues_(ctx.ss, [{ ...baseRow, butai: '第四部隊' }], '向')[0];
    expect(out[ctx.g.HEADERS.indexOf('部隊')]).toBe('第四部隊');
  });

  it('★画面が空欄を送ってきたら空欄のまま（既定部隊で戻さない）', () => {
    const out = ctx.g.buildDailyValues_(ctx.ss, [{ ...baseRow, butai: '' }], '向')[0];
    expect(out[ctx.g.HEADERS.indexOf('部隊')]).toBe('');
  });

  it('既定部隊が無い人は空欄になる', () => {
    const out = ctx.g.buildDailyValues_(ctx.ss, [{ ...baseRow, name: '中島' }], '向')[0];
    expect(out[ctx.g.HEADERS.indexOf('部隊')]).toBe('');
  });

  it('拠点は現場マスタから入る（部隊を足しても壊れていない）', () => {
    const out = ctx.g.buildDailyValues_(ctx.ss, [{ ...baseRow, loc: '関東B' }], '向')[0];
    expect(out[ctx.g.HEADERS.indexOf('拠点')]).toBe('関東支店');
  });

  it('拠点の軸を持たない会社は拠点が空（部隊とは独立）', () => {
    const out = ctx.g.buildDailyValues_(ctx.ss, [{ ...baseRow, company: '和信カインド' }], '向')[0];
    expect(out[ctx.g.HEADERS.indexOf('拠点')]).toBe('');
  });

  it('複数人まとめて登録しても全員21列', () => {
    const out = ctx.g.buildDailyValues_(ctx.ss,
      [baseRow, { ...baseRow, name: '中島', role: '同行', id: 'ID1' }], '向');
    expect(out.length).toBe(2);
    out.forEach(r => expect(r.length).toBe(21));
    expect(out[0][ctx.g.HEADERS.indexOf('部隊')]).toBe('第二部隊');
    expect(out[1][ctx.g.HEADERS.indexOf('部隊')]).toBe('');
  });

  it('★既定部隊の鍵は (会社, 氏名)（同名の別会社を取り違えない）', () => {
    const map = ctx.g.getMemberButaiMap_(ctx.ss);
    expect(map[ctx.g.memberKey_('グローライズ', '元')]).toBe('第二部隊');
    expect(ctx.g.lookupMemberButai_(map, 'グローライズ', '元')).toBe('第二部隊');
    expect(ctx.g.lookupMemberButai_(map, 'グローライズ', '中島')).toBe('');
  });

  it('会社が分からなくても氏名で引ける（1人しかいない名前なら）', () => {
    const map = ctx.g.getMemberButaiMap_(ctx.ss);
    expect(ctx.g.lookupMemberButai_(map, '', '元')).toBe('第二部隊');
  });

  it('★同名が別会社にいて部隊が違うときは氏名だけでは引かない（取り違え防止）', () => {
    const ctx2 = makeContext({
      '日報データ': makeSheet([[]]),
      '職人マスタ': makeSheet([
        ['氏名', '会社', '事業部', '単価', '既定部隊', '有効'],
        ['元', 'グローライズ', '', 0, '第二部隊', '○'],
        ['元', 'GRミツマ', '', 0, '第三部隊', '○']
      ]),
      '現場マスタ': makeSheet(JOBSITE_ROWS.map(r => r.slice())),
      '元請マスタ': makeSheet([['元請名', '会社', '読み']])
    });
    const map = ctx2.g.getMemberButaiMap_(ctx2.ss);
    expect(ctx2.g.lookupMemberButai_(map, 'グローライズ', '元')).toBe('第二部隊');
    expect(ctx2.g.lookupMemberButai_(map, 'GRミツマ', '元')).toBe('第三部隊');
    expect(ctx2.g.lookupMemberButai_(map, '', '元')).toBe('');   // どちらか分からないので入れない
  });

  it('会社が違えば既定部隊も会社ごとに正しく入る', () => {
    const ctx2 = makeContext({
      '日報データ': makeSheet([[]]),
      '職人マスタ': makeSheet([
        ['氏名', '会社', '事業部', '単価', '既定部隊', '有効'],
        ['元', 'グローライズ', '', 0, '第二部隊', '○'],
        ['元', 'GRミツマ', '', 0, '第三部隊', '○']
      ]),
      '現場マスタ': makeSheet(JOBSITE_ROWS.map(r => r.slice())),
      '元請マスタ': makeSheet([['元請名', '会社', '読み']])
    });
    const base = {
      date: '2026-08-28', genba: 'きんでん西', loc: 'A現場', name: '元', role: '代表',
      start: '08:00', end: '17:00', kosu: 1, memo: '', id: 'ID1',
      updatedBy: '向', color: '', workType: '現場作業', vehicle: ''
    };
    const bi = ctx2.g.HEADERS.indexOf('部隊');
    expect(ctx2.g.buildDailyValues_(ctx2.ss, [{ ...base, company: 'グローライズ' }], '向')[0][bi]).toBe('第二部隊');
    expect(ctx2.g.buildDailyValues_(ctx2.ss, [{ ...base, company: 'GRミツマ' }], '向')[0][bi]).toBe('第三部隊');
  });
});

describe('変更履歴が実際の行の形で動く', () => {
  const mk = (over) => {
    const base = {
      date: '2026-08-28', genba: 'きんでん西', loc: 'A現場', name: '元', role: '代表',
      start: '08:00', end: '17:00', kosu: 1, memo: '', company: 'グローライズ',
      id: 'OLD', updatedBy: '向', color: '', workType: '現場作業', vehicle: ''
    };
    return { ...base, ...over };
  };

  it('★組み立てた行同士を突き合わせられる（列の形が揃っている）', () => {
    const oldRows = ctx.g.buildDailyValues_(ctx.ss, [mk({ id: 'OLD' })], '向');
    const newRows = ctx.g.buildDailyValues_(ctx.ss, [mk({ id: 'NEW', memo: '変更した' })], '向');
    const d = ctx.g.diffDailyRows_(ctx.g.HEADERS, oldRows, newRows);
    const memo = d.find(x => x.field === 'メモ');
    expect(memo).toBeTruthy();
    expect(memo.before).toBe('');
    expect(memo.after).toBe('変更した');
    expect(memo.oldId).toBe('OLD');
    expect(memo.newId).toBe('NEW');
  });

  it('部隊だけを変えた編集も記録される', () => {
    const oldRows = ctx.g.buildDailyValues_(ctx.ss, [mk({ id: 'OLD', butai: '第一部隊' })], '向');
    const newRows = ctx.g.buildDailyValues_(ctx.ss, [mk({ id: 'NEW', butai: '第三部隊' })], '向');
    const d = ctx.g.diffDailyRows_(ctx.g.HEADERS, oldRows, newRows);
    const b = d.find(x => x.field === '部隊');
    expect(b.before).toBe('第一部隊');
    expect(b.after).toBe('第三部隊');
  });

  it('★削除の記録から元の予定を復元できる（21項目そろっている）', () => {
    const rows = ctx.g.buildDailyValues_(ctx.ss, [mk({ id: 'DEL', memo: '大事なメモ' })], '向');
    const o = JSON.parse(ctx.g.rowFullJson_(ctx.g.HEADERS, rows[0]));
    expect(Object.keys(o).length).toBe(21);
    expect(o['メモ']).toBe('大事なメモ');
    expect(o['氏名']).toBe('元');
    expect(o['部隊']).toBe('第二部隊');
    expect(o['拠点']).toBe('本社');
    expect(o['作業区分']).toBe('現場作業');
  });
});
