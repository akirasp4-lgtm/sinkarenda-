// gas.js（Google Apps Script用）の「純粋な関数」だけを取り出して試験する。
//
// ★2026-08-27 実測で判明した注意点（この方式でないと動かない）:
//   vm に読み込んでも `const HEADERS = ...` は**コンテキストの属性にならない**
//   （const/let は字句束縛でグローバルオブジェクトに載らない。var と function だけが載る）。
//   そのため gas.js の末尾に「同じ字句スコープのまま外へ出す」1行を足してから実行する。
//   ここに列挙し忘れた名前はテストから見えないので、関数を足したら必ずここにも足すこと。
import { describe, it, expect, beforeAll } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const GAS_PATH = join(here, '..', '..', 'gas.js');

const EXPORT_SNIPPET = `
;globalThis.__gas = {
  HEADERS, BUTAI_VALUES,
  normalizeButai_, resolveButai_, normalizeMemberActive_
};`;

let ctx;   // sandbox.__gas
beforeAll(() => {
  const code = readFileSync(GAS_PATH, 'utf8');
  // Apps Script のグローバルを最低限だけ用意する（純粋関数の試験が目的）
  const sandbox = vm.createContext({
    SpreadsheetApp: { getActiveSpreadsheet: () => null, flush() {} },
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    Utilities: {}, ContentService: {}, PropertiesService: {},
    UrlFetchApp: {}, Logger: { log() {} }, console
  });
  vm.runInContext(code + EXPORT_SNIPPET, sandbox, { filename: 'gas.js' });
  ctx = sandbox.__gas;
});

describe('HEADERS', () => {
  it('21列で、21列目が部隊', () => {
    expect(ctx.HEADERS.length).toBe(21);
    expect(ctx.HEADERS[20]).toBe('部隊');
  });

  it('先頭19列は1つも動いていない', () => {
    expect(ctx.HEADERS.slice(0, 19)).toEqual([
      '登録日時', '作業日', '元請名', '現場名', '氏名', '役割', '出勤', '退勤',
      '人工', 'メモ', '夜勤', '会社', 'ID', '更新者', '色', '事業部', '工番', '作業区分', '車両'
    ]);
  });

  it('20列目は拠点のまま', () => {
    expect(ctx.HEADERS[19]).toBe('拠点');
  });
});

describe('normalizeButai_', () => {
  it('1〜4部隊はそのまま通す', () => {
    ['1部隊', '2部隊', '3部隊', '4部隊'].forEach(v =>
      expect(ctx.normalizeButai_(v)).toBe(v));
  });

  it('前後の空白を落とす', () => {
    expect(ctx.normalizeButai_('  2部隊 ')).toBe('2部隊');
  });

  it('知らない値は空にする', () => {
    ['5部隊', '部隊', 'A班', '1', 1, null, undefined, ''].forEach(v =>
      expect(ctx.normalizeButai_(v)).toBe(''));
  });

  it('部隊の値は1〜4部隊の4つだけ', () => {
    expect(ctx.BUTAI_VALUES).toEqual(['1部隊', '2部隊', '3部隊', '4部隊']);
  });
});

describe('resolveButai_', () => {
  it('画面が値を送ってきたらそれを使う', () => {
    expect(ctx.resolveButai_({ butai: '3部隊' }, '1部隊')).toBe('3部隊');
  });

  it('★画面が「空欄」を送ってきたら空欄のまま（既定値で上書きしない）', () => {
    // 事務所・休みなど「部隊に属さない」を明示できるようにするため。
    // 拠点で起きたバグ（手で消した値が既定値に戻る）を繰り返さない。
    expect(ctx.resolveButai_({ butai: '' }, '1部隊')).toBe('');
  });

  it('画面が項目そのものを送ってこなければ職人マスタの既定部隊を使う', () => {
    expect(ctx.resolveButai_({}, '1部隊')).toBe('1部隊');
  });

  it('既定部隊も無ければ空', () => {
    expect(ctx.resolveButai_({}, '')).toBe('');
    expect(ctx.resolveButai_({}, undefined)).toBe('');
  });

  it('既定部隊が壊れた値でも空にする', () => {
    expect(ctx.resolveButai_({}, '9部隊')).toBe('');
  });

  it('画面が送ってきた値が壊れていれば空（既定値へは戻さない）', () => {
    expect(ctx.resolveButai_({ butai: 'A班' }, '1部隊')).toBe('');
  });

  it('rowがnull/undefinedでも落ちない', () => {
    expect(ctx.resolveButai_(null, '2部隊')).toBe('2部隊');
    expect(ctx.resolveButai_(undefined, '')).toBe('');
  });
});

describe('職人の有効/無効', () => {
  it('×だけが無効。それ以外は全部有効', () => {
    ['×', 'x', 'X', '✕'].forEach(v =>
      expect(ctx.normalizeMemberActive_(v)).toBe(false));
    ['○', 'o', '', '　', undefined, null, true].forEach(v =>
      expect(ctx.normalizeMemberActive_(v)).toBe(true));
  });

  it('★空欄は有効（既存71件を巻き込まないための既定）', () => {
    expect(ctx.normalizeMemberActive_('')).toBe(true);
  });
});
