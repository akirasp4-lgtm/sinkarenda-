// 二重登録の統合を「偽のスプレッドシート」の上で実際に動かす結合テスト。
//
// 本番データを触る移行処理なので、単体テストだけでは足りない。
// 実際にシートを読み書きさせて、
//   ・行数が1行も増減しないこと
//   ・会社を1つも書き換えないこと
//   ・二度実行しても結果が変わらないこと
//   ・アーカイブも直ること
// を確かめる。
import { describe, it, expect, beforeEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const CODE = readFileSync(join(here, '..', '..', 'gas.js'), 'utf8');
const EXPORT = `
;globalThis.__gas = { HEADERS, mergedMemberName_, mergedUpdaterName_, planNameMerge_,
  mergeDuplicateMembers, MEMBER_MERGE_BY_COMPANY, UPDATER_MERGE, MEMBER_MERGE_COMPANY, buildDailyValues_ };`;

function makeSheet(rows) {
  const data = rows.map(r => r.slice());
  return {
    _data: data,
    getName: () => data.__name || 'sheet',
    getDataRange: () => ({ getValues: () => data.map(r => r.slice()) }),
    getMaxColumns: () => Math.max(...data.map(r => r.length), 1),
    getMaxRows: () => data.length,
    getLastRow: () => data.length,
    getLastColumn: () => Math.max(...data.map(r => r.length), 1),
    insertColumnsAfter: () => {}, insertRowsAfter: () => {},
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
      setValue: (v) => { while (data.length < row) data.push([]); data[row - 1][col - 1] = v; },
      setValues: (vals) => {
        vals.forEach((rv, i) => {
          while (data.length < row + i) data.push([]);
          rv.forEach((v, j) => { data[row - 1 + i][col - 1 + j] = v; });
        });
      },
      clearContent: () => {
        for (let i = 0; i < (nRows || 1); i++) {
          const r = data[row - 1 + i];
          if (!r) continue;
          for (let j = 0; j < (nCols || 1); j++) r[col - 1 + j] = '';
        }
      }
    })
  };
}

const H21 = ['登録日時','作業日','元請名','現場名','氏名','役割','出勤','退勤','人工','メモ',
             '夜勤','会社','ID','更新者','色','事業部','工番','作業区分','車両','拠点','部隊'];

// 実データと同じ形の予定行を作る
function row(name, company, updatedBy, id) {
  const o = {'登録日時':'2026/08/01 10:00','作業日':'2026-08-01','元請名':'きんでん西',
    '現場名':'A現場','氏名':name,'役割':'代表','出勤':'08:00','退勤':'17:00','人工':1,
    'メモ':'','夜勤':'','会社':company,'ID':id,'更新者':updatedBy,'色':'','事業部':'',
    '工番':'','作業区分':'現場作業','車両':'','拠点':'本社','部隊':''};
  return H21.map(h => o[h]);
}

function build() {
  const sheets = {
    '日報データ': makeSheet([H21,
      row('高田', 'GRミツマ', '高田', 'A1'),
      row('GRME髙田', 'グローライズ', '向', 'A2'),
      row('柳澤', 'GRミツマ', '向', 'A3'),
      row('栁澤', 'GRミツマ', '向', 'A4'),
      row('GRME栁澤', 'グローライズ', '向', 'A5'),
      row('内村', 'GRミツマ', '向', 'A6'),
      row('中島', 'グローライズ', '中島', 'A7')     // 対象外
    ]),
    'アーカイブ': makeSheet([H21,
      row('高田', 'GRミツマ', '高田', 'B1'),
      row('GRME髙田', 'グローライズ', '向', 'B2'),
      row('東', 'グローライズ', '東', 'B3')          // 対象外
    ]),
    '職人マスタ': makeSheet([
      ['氏名','会社','事業部','単価','既定部隊','有効'],
      ['高田', 'GRミツマ', '', 30000, '', '○'],
      ['GRME髙田', 'グローライズ', 'ICT', 0, '', '○'],
      ['柳澤', 'GRミツマ', '', 0, '', '○'],
      ['栁澤', 'GRミツマ', '', 25000, '', '○'],
      ['GRME栁澤', 'グローライズ', '', 0, '', '○'],
      ['内村', 'GRミツマ', '', 0, '', '○'],
      ['GRME内村', 'グローライズ', '', 0, '', '○'],
      ['中島', 'グローライズ', 'ICT', 28000, '第一部隊', '○']
    ]),
    '現場マスタ': makeSheet([['元請名','現場名','工番','事業部','年度','連番','売上','読み','完了','請求方式','拠点','ステータス']]),
    '元請マスタ': makeSheet([['元請名','会社','読み']]),
    '操作ログ': makeSheet([['日時','操作','対象','詳細','実行者']])
  };
  const ss = { getSheetByName: (n) => sheets[n] || null, insertSheet: (n) => (sheets[n] = makeSheet([[]])) };
  const sandbox = vm.createContext({
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush() {} },
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    Utilities: { formatDate: (d, tz, f) => String(f) },
    ContentService: {},
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} }) },
    UrlFetchApp: {}, Logger: { log() {} }, console
  });
  vm.runInContext(CODE + EXPORT, sandbox, { filename: 'gas.js' });
  return { g: sandbox.__gas, ss, sheets };
}

const col = (sheet, h) => {
  const d = sheet._data, i = d[0].indexOf(h);
  return d.slice(1).map(r => String(r[i] == null ? '' : r[i]));
};

let ctx;
beforeEach(() => { ctx = build(); });

describe('二重登録の統合（偽スプレッドシート上で実際に動かす）', () => {
  it('dry-run では1文字も書き換えない', () => {
    const before = JSON.stringify(ctx.sheets['日報データ']._data);
    const rep = ctx.g.mergeDuplicateMembers(false);
    expect(rep.dryRun).toBe(true);
    expect(JSON.stringify(ctx.sheets['日報データ']._data)).toBe(before);
  });

  it('dry-run が直す件数を正しく数える', () => {
    const rep = ctx.g.mergeDuplicateMembers(false);
    expect(rep.シート.length).toBe(2);           // 日報データ と アーカイブ
    expect(rep.食い違い.length).toBe(0);
    expect(rep.中止理由).toBe('');
  });

  it('★実行すると氏名がまとまる', () => {
    ctx.g.mergeDuplicateMembers(true);
    const names = col(ctx.sheets['日報データ'], '氏名');
    expect(names).toEqual(['高田（関東）','高田（関東）','柳澤（関東）','柳澤（関東）',
                           '柳澤（関東）','内村（関東）','中島']);
  });

  it('★会社は1つも書き換えない（本社案件に入った記録を壊さない）', () => {
    const before = col(ctx.sheets['日報データ'], '会社');
    ctx.g.mergeDuplicateMembers(true);
    expect(col(ctx.sheets['日報データ'], '会社')).toEqual(before);
    expect(col(ctx.sheets['日報データ'], '会社')).toContain('グローライズ');
    expect(col(ctx.sheets['日報データ'], '会社')).toContain('GRミツマ');
  });

  it('★アーカイブも直る（忘れると3ヶ月より前だけ名前が割れる）', () => {
    ctx.g.mergeDuplicateMembers(true);
    expect(col(ctx.sheets['アーカイブ'], '氏名')).toEqual(['高田（関東）','高田（関東）','東']);
  });

  it('更新者も直る（同じ人の操作履歴が2つに割れない）', () => {
    ctx.g.mergeDuplicateMembers(true);
    expect(col(ctx.sheets['日報データ'], '更新者')).toContain('高田（関東）');
    expect(col(ctx.sheets['日報データ'], '更新者')).not.toContain('高田');
  });

  it('★行数は1行も増減しない', () => {
    const n = ctx.sheets['日報データ']._data.length;
    const a = ctx.sheets['アーカイブ']._data.length;
    ctx.g.mergeDuplicateMembers(true);
    expect(ctx.sheets['日報データ']._data.length).toBe(n);
    expect(ctx.sheets['アーカイブ']._data.length).toBe(a);
  });

  it('★対象外の人は1文字も変わらない', () => {
    ctx.g.mergeDuplicateMembers(true);
    expect(col(ctx.sheets['日報データ'], '氏名')).toContain('中島');
    expect(col(ctx.sheets['アーカイブ'], '氏名')).toContain('東');
  });

  it('★職人マスタが1人1行にまとまる', () => {
    ctx.g.mergeDuplicateMembers(true);
    const names = col(ctx.sheets['職人マスタ'], '氏名').filter(Boolean);
    expect(names.sort()).toEqual(['中島', '内村（関東）', '柳澤（関東）', '高田（関東）'].sort());
  });

  it('★単価（給料の元数字）を失わない', () => {
    ctx.g.mergeDuplicateMembers(true);
    const d = ctx.sheets['職人マスタ']._data;
    const find = (n) => d.slice(1).find(r => String(r[0]).trim() === n);
    expect(Number(find('高田（関東）')[3])).toBe(30000);   // 高田(30000) と GRME髙田(0) を寄せる
    expect(Number(find('柳澤（関東）')[3])).toBe(25000);   // 栁澤(25000) を拾う
    expect(String(find('高田（関東）')[2])).toBe('ICT');   // GRME髙田の事業部を拾う
  });

  it('まとめた人の会社はGRミツマ（関東支店の実態）', () => {
    ctx.g.mergeDuplicateMembers(true);
    const d = ctx.sheets['職人マスタ']._data;
    ['高田（関東）','柳澤（関東）','内村（関東）'].forEach(n => {
      const r = d.slice(1).find(x => String(x[0]).trim() === n);
      expect(r[1]).toBe('GRミツマ');
    });
  });

  it('★★二度実行しても結果が変わらない（冪等）', () => {
    ctx.g.mergeDuplicateMembers(true);
    const after1 = JSON.stringify({
      n: ctx.sheets['日報データ']._data, a: ctx.sheets['アーカイブ']._data,
      m: ctx.sheets['職人マスタ']._data
    });
    ctx.g.mergeDuplicateMembers(true);
    const after2 = JSON.stringify({
      n: ctx.sheets['日報データ']._data, a: ctx.sheets['アーカイブ']._data,
      m: ctx.sheets['職人マスタ']._data
    });
    expect(after2).toBe(after1);
  });

  it('単価が食い違うときは中止して何も書かない', () => {
    // 高田(30000) と GRME髙田(28000) の両方に値があると決められない
    ctx.sheets['職人マスタ']._data[2][3] = 28000;
    const before = JSON.stringify(ctx.sheets['職人マスタ']._data);
    const rep = ctx.g.mergeDuplicateMembers(true);
    expect(rep.食い違い.length).toBeGreaterThan(0);
    expect(JSON.stringify(ctx.sheets['職人マスタ']._data)).toBe(before);
  });
});

describe('★Codexレビューで直した点の追試', () => {
  it('[P1]#1 他社の同姓同名を巻き込まない', () => {
    // 和信カインドに別人の「高田」がいる状態を作る
    ctx.sheets['日報データ']._data.push(row('高田', '和信カインド', '元', 'X1'));
    ctx.g.mergeDuplicateMembers(true);
    const d = ctx.sheets['日報データ']._data;
    const last = d[d.length - 1];
    expect(String(last[H21.indexOf('氏名')])).toBe('高田');       // 変えない
    expect(String(last[H21.indexOf('会社')])).toBe('和信カインド');
  });

  it('★[P1]#2 食い違いがあれば予定データも1文字も書かない', () => {
    ctx.sheets['職人マスタ']._data[2][3] = 28000;   // 単価が食い違う
    const beforeN = JSON.stringify(ctx.sheets['日報データ']._data);
    const beforeA = JSON.stringify(ctx.sheets['アーカイブ']._data);
    const beforeM = JSON.stringify(ctx.sheets['職人マスタ']._data);
    const rep = ctx.g.mergeDuplicateMembers(true);
    expect(rep.食い違い.length).toBeGreaterThan(0);
    expect(rep.中止理由).toContain('食い違い');
    expect(JSON.stringify(ctx.sheets['日報データ']._data)).toBe(beforeN);
    expect(JSON.stringify(ctx.sheets['アーカイブ']._data)).toBe(beforeA);
    expect(JSON.stringify(ctx.sheets['職人マスタ']._data)).toBe(beforeM);
  });

  it('[P2]#5 職人マスタが無ければ何も書かずに中止する', () => {
    delete ctx.sheets['職人マスタ'];
    const before = JSON.stringify(ctx.sheets['日報データ']._data);
    const rep = ctx.g.mergeDuplicateMembers(true);
    expect(rep.中止理由).toContain('職人マスタ');
    expect(JSON.stringify(ctx.sheets['日報データ']._data)).toBe(before);
  });

  it('★[P2]#6 保存の入口でも旧名を読み替える（開いたままの端末が復活させない）', () => {
    const out = ctx.g.buildDailyValues_(ctx.ss, [{
      date:'2026-09-01', genba:'きんでん西', loc:'A現場',
      name:'GRME髙田', role:'代表', start:'08:00', end:'17:00', kosu:1, memo:'',
      company:'グローライズ', id:'Z1', updatedBy:'高田', color:'', workType:'現場作業', vehicle:''
    }], '高田')[0];
    expect(out[H21.indexOf('氏名')]).toBe('高田（関東）');
    expect(out[H21.indexOf('更新者')]).toBe('高田（関東）');
    expect(out[H21.indexOf('会社')]).toBe('グローライズ');   // 会社は触らない
  });

  it('保存の入口: 対象外の人は1文字も変えない', () => {
    const out = ctx.g.buildDailyValues_(ctx.ss, [{
      date:'2026-09-01', genba:'きんでん西', loc:'A現場',
      name:'中島', role:'代表', start:'08:00', end:'17:00', kosu:1, memo:'',
      company:'グローライズ', id:'Z2', updatedBy:'中島', color:'', workType:'現場作業', vehicle:''
    }], '中島')[0];
    expect(out[H21.indexOf('氏名')]).toBe('中島');
    expect(out[H21.indexOf('更新者')]).toBe('中島');
  });
});
