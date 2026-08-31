// 職人を消したとき、変更履歴が残ることを実際に動かして確かめる（2026-08-31）。
//
// ★なぜ必要か:
//   事業部・既定部隊・有効・単価の変更は全部 logOperation_ を書いているのに、
//   一番取り返しがつかない「消す」だけが**1行も記録していなかった**。
//   誰がいつ誰を消したのか、後から誰にも分からない状態だった。
//   社長指示 §0「変更履歴を維持する」／依頼書の必須仕様⑦に反する。
//
// ★守ること:
//   ・消す前に履歴を書く。書けなければ1行も消さずエラー（履歴が命綱）
//   ・単価（日当）は履歴に書かない。変更履歴は画面から読めるので、
//     せっかく窓口から外した日当が履歴経由で出てしまう
import { describe, it, expect, beforeEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const CODE = readFileSync(join(here, '..', '..', 'gas.js'), 'utf8');

function makeSheet(rows) {
  const data = rows.map((r) => r.slice());
  return {
    _data: data,
    _failWrite: false,
    getDataRange: () => ({ getValues: () => data.map((r) => r.slice()) }),
    getRange(row, col, nr, nc) {
      const self = this;
      return {
        setValue: (v) => { data[row - 1][col - 1] = v; },
        setValues: (vv) => {
          if (self._failWrite) throw new Error('書き込めません（テスト）');
          vv.forEach((r, k) => { data[row - 1 + k] = r.slice(); });
        },
        getValues: () => [data[row - 1].slice(col - 1, col - 1 + (nc || 1))]
      };
    },
    appendRow: (r) => { data.push(r.slice()); },
    deleteRow: (row) => { data.splice(row - 1, 1); },
    getMaxColumns: () => Math.max(...data.map((r) => r.length), 1),
    getLastRow: () => data.length,
    getLastColumn: () => Math.max(...data.map((r) => r.length), 1),
    insertColumnsAfter: () => {},
    getName: () => 'sheet'
  };
}

const MEMBER_HEADERS = ['氏名', '会社', '事業部', '単価', '既定部隊', '有効'];
// ★単価 23000 が履歴に出てはいけない
const TARO = ['山田太郎', 'グローライズ', 'ICT', 23000, '第一部隊', '有効'];
const HANAKO = ['田中花子', '和信カインド', '設備', 19000, '', ''];

const fakeLock = () => ({ tryLock: () => true, waitLock: () => {}, releaseLock: () => {}, hasLock: () => true });

let G, sheets;

beforeEach(() => {
  const member = makeSheet([MEMBER_HEADERS.slice(), TARO.slice(), HANAKO.slice()]);
  const history = makeSheet([['日時', '操作', '旧ID', '新ID', '項目', '変更前', '変更後', '実行者']]);
  const oplog = makeSheet([['日時', '操作', '対象', '内容', '実行者']]);
  sheets = { member, history, oplog };

  const box = vm.createContext({
    console, String, Number, Object, Array, Math, isFinite, JSON, Date, RegExp,
    SpreadsheetApp: { flush: () => {}, getActiveSpreadsheet: () => ss },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} }) },
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

  const others = {};
  const ss = {
    getSheetByName: (n) => {
      if (n === '職人マスタ') return member;
      if (n === '変更履歴') return history;
      if (n === '操作ログ') return oplog;
      if (!others[n]) others[n] = makeSheet([[]]);
      return others[n];
    },
    insertSheet: (n) => { others[n] = makeSheet([[]]); return others[n]; }
  };

  vm.runInContext(CODE + ';globalThis.__g = { doPost };', box, { filename: 'gas.js' });
  G = box.__g;
});

const post = (body) => JSON.parse(G.doPost({ postData: { contents: JSON.stringify(body) } })._t);
const names = () => sheets.member._data.slice(1).map((r) => r[0]);
const hist = () => sheets.history._data.slice(1);
const del = (extra) => post(Object.assign(
  { action: 'remove_member', name: '山田太郎', company: 'グローライズ', updatedBy: '事務員A' }, extra || {}));

describe('職人を消したら変更履歴が残る', () => {
  it('消せる（今までどおり動く）', () => {
    const r = del();
    expect(r.status).toBe('ok');
    expect(r.removed).toBe('山田太郎');
    expect(names()).toEqual(['田中花子']);
  });

  it('★変更履歴に1行以上残る（今まで0行だった）', () => {
    del();
    expect(hist().length).toBeGreaterThan(0);
  });

  it('★誰を消したか・誰が消したかが残る', () => {
    del();
    const flat = hist().map((r) => r.join('|')).join('\n');
    expect(flat).toContain('remove_member');
    expect(flat).toContain('山田太郎');
    expect(flat).toContain('グローライズ');
    expect(flat).toContain('事務員A');
  });

  it('★消える前の中身（事業部・部隊・有効）が残る', () => {
    del();
    const flat = hist().map((r) => r.join('|')).join('\n');
    expect(flat).toContain('ICT');
    expect(flat).toContain('第一部隊');
    expect(flat).toContain('有効');
  });

  it('★★単価（日当）は履歴に書かない', () => {
    // 変更履歴は画面から読める。ここに書くと、窓口から外した日当が履歴経由で出る。
    del();
    const flat = hist().map((r) => r.join('|')).join('\n');
    expect(flat).not.toContain('23000');
    expect(flat).not.toContain('単価');
  });

  it('★履歴が書けなければ、1人も消さずにエラーを返す', () => {
    sheets.history._failWrite = true;
    const r = del();
    expect(r.status).toBe('error');
    expect(String(r.message)).toContain('変更履歴');
    expect(names(), '★履歴が書けていないのに消えた').toEqual(['山田太郎', '田中花子']);
  });

  it('★他社の同姓同名を巻き込まない（会社が違えば消さない）', () => {
    const r = post({
      action: 'remove_member', name: '山田太郎', company: '和信カインド', updatedBy: '事務員A'
    });
    expect(r.removed).toBe(null);
    expect(names()).toEqual(['山田太郎', '田中花子']);
    expect(hist().length, '消していないのに履歴を書いた').toBe(0);
  });

  it('居ない人を消そうとしても履歴を汚さない', () => {
    post({ action: 'remove_member', name: '居ない人', company: 'グローライズ', updatedBy: '事務員A' });
    expect(hist().length).toBe(0);
  });

  it('空欄の項目は履歴に出さない（田中花子は部隊も有効も空）', () => {
    post({ action: 'remove_member', name: '田中花子', company: '和信カインド', updatedBy: '事務員A' });
    const fields = hist().map((r) => r[4]);
    expect(fields).toContain('氏名');
    expect(fields).toContain('事業部');
    expect(fields).not.toContain('既定部隊');
    expect(fields).not.toContain('有効');
  });

  it('操作ログにも残る', () => {
    del();
    const flat = sheets.oplog._data.slice(1).map((r) => r.join('|')).join('\n');
    expect(flat).toContain('remove_member');
    expect(flat).toContain('山田太郎');
  });
});

// ================================================================
// ログイン調査（2026-08-31）で見つかった記録の抜け 2件。
//
// 予定の削除・職人の削除・現場の削除は全部
// 「履歴が書けなければ処理を中止する」まで徹底しているのに、
// 元請マスタの削除だけが1行も記録していなかった。
// 請求単価は金額に直結するのに操作ログが無かった。
// ================================================================

describe('元請を消したら記録が残る（remove_genba）', () => {
  let G2, sheets2;

  beforeEach(() => {
    const genba = makeSheet([['元請名', '会社', '読み'],
      ['きんでん東', 'グローライズ', 'きんでんひがし'],
      ['エクシオ', '和信カインド', 'えくしお']]);
    const history = makeSheet([['日時', '操作', '旧ID', '新ID', '項目', '変更前', '変更後', '実行者']]);
    const oplog = makeSheet([['日時', '操作', '対象', '内容', '実行者']]);
    sheets2 = { genba, history, oplog };

    const box = vm.createContext({
      console, String, Number, Object, Array, Math, isFinite, JSON, Date, RegExp,
      SpreadsheetApp: { flush: () => {}, getActiveSpreadsheet: () => ss },
      PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} }) },
      LockService: {
        getScriptLock: () => fakeLock(), getUserLock: () => fakeLock(), getDocumentLock: () => fakeLock()
      },
      ContentService: {
        MimeType: { JSON: 'json' },
        createTextOutput: (t) => ({ setMimeType: () => ({ _t: t }), _t: t })
      },
      Utilities: { formatDate: () => '2026-08-31 14:00' }
    });
    box.globalThis = box;
    const others = {};
    const ss = {
      getSheetByName: (n) => {
        if (n === '元請マスタ') return genba;
        if (n === '変更履歴') return history;
        if (n === '操作ログ') return oplog;
        if (!others[n]) others[n] = makeSheet([[]]);
        return others[n];
      },
      insertSheet: (n) => { others[n] = makeSheet([[]]); return others[n]; }
    };
    vm.runInContext(CODE + ';globalThis.__g = { doPost };', box, { filename: 'gas.js' });
    G2 = box.__g;
  });

  const post2 = (b) => JSON.parse(G2.doPost({ postData: { contents: JSON.stringify(b) } })._t);
  const genbaNames = () => sheets2.genba._data.slice(1).map((r) => r[0]);
  const hist2 = () => sheets2.history._data.slice(1);
  const del2 = () => post2({
    action: 'remove_genba', name: 'きんでん東', company: 'グローライズ', updatedBy: '事務員A'
  });

  it('消せる（今までどおり動く）', () => {
    expect(del2().removed).toBe('きんでん東');
    expect(genbaNames()).toEqual(['エクシオ']);
  });

  it('★変更履歴に残る（今まで0行だった）', () => {
    del2();
    const flat = hist2().map((r) => r.join('|')).join('\n');
    expect(hist2().length).toBeGreaterThan(0);
    expect(flat).toContain('remove_genba');
    expect(flat).toContain('きんでん東');
    expect(flat).toContain('グローライズ');
    expect(flat).toContain('事務員A');
  });

  it('★履歴が書けなければ、1件も消さずにエラーを返す', () => {
    sheets2.history._failWrite = true;
    const r = del2();
    expect(r.status).toBe('error');
    expect(String(r.message)).toContain('変更履歴');
    expect(genbaNames(), '★履歴が書けていないのに消えた').toEqual(['きんでん東', 'エクシオ']);
  });

  it('★他社の同名を巻き込まない', () => {
    const r = post2({
      action: 'remove_genba', name: 'エクシオ', company: 'グローライズ', updatedBy: '事務員A'
    });
    expect(r.removed).toBe(null);
    expect(genbaNames()).toEqual(['きんでん東', 'エクシオ']);
    expect(hist2().length, '消していないのに履歴を書いた').toBe(0);
  });

  it('操作ログにも残る', () => {
    del2();
    const flat = sheets2.oplog._data.slice(1).map((r) => r.join('|')).join('\n');
    expect(flat).toContain('remove_genba');
    expect(flat).toContain('きんでん東');
  });
});
