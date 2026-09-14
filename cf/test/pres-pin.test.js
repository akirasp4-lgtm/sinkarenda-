// 社長用カレンダーの鍵を、コードから設定欄（スクリプトプロパティ）へ移した件の検証。
//
// ★なぜ必要か（2026-09-14・実測した事実）:
//   鍵の4桁が gas.js と president.html の両方に直接書かれており、
//   president.html は公開されているため、ログインも何もせずに
//   HTTP 200 でファイルごと取得でき、中に「APIのアドレス」と「鍵」が
//   そろって入っていた。つまり社長の予定を読む・書く・消すのに必要な物が
//   全部公開されていた。
//
// ★ここで見張るのは2点:
//   1. コードに鍵の値が戻ってこないこと
//   2. 設定欄が空のとき、素通しではなく全拒否（fail-closed）になること
//      — 空文字どうしの比較で誰でも通る、という壊れ方をしやすい所。
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..', '..');
const CODE = readFileSync(join(ROOT, 'gas.js'), 'utf8');

// gas.js をそのまま読み込み、GAS側のAPIだけ偽物に差し替える。
// 実装とテストが同じコードを見るので、テストだけが実装とずれる事故が起きない。
function loadGas({ props = {}, presRows = [] } = {}) {
  const sandbox = { console };
  sandbox.globalThis = sandbox;
  sandbox.PropertiesService = {
    getScriptProperties: () => ({ getProperty: k => (k in props ? props[k] : null) })
  };
  sandbox.ContentService = {
    MimeType: { JSON: 'json' },
    createTextOutput: t => ({ _t: t, setMimeType(){ return this; } })
  };
  const sheet = {
    getDataRange: () => ({ getValues: () => presRows }),
    appendRow: () => {}, getMaxColumns: () => 20, getRange: () => ({ setValue(){} })
  };
  sandbox.SpreadsheetApp = { getActiveSpreadsheet: () => ({ getSheetByName: () => sheet, insertSheet: () => sheet }) };
  sandbox.LockService = {
    getUserLock: () => ({ tryLock: () => true, releaseLock(){} }),
    getScriptLock: () => ({ tryLock: () => true, releaseLock(){} })
  };
  sandbox.Session = { getScriptTimeZone: () => 'Asia/Tokyo' };
  sandbox.Utilities = { getUuid: () => 'x', formatDate: d => String(d) };
  vm.createContext(sandbox);
  vm.runInContext(CODE, sandbox, { filename: 'gas.js' });
  return sandbox;
}

// ContentService の戻りからJSONを取り出す
const parse = r => JSON.parse(r._t);

describe('社長用カレンダーの鍵は設定欄から読む', () => {
  it('★コードに鍵の値を書かない（このファイルは公開リポジトリへ上がる）', () => {
    expect(CODE, 'PRES_PIN に値を直書きしている')
      .not.toMatch(/PRES_PIN\s*=\s*['"][^'"]+['"]/);
    expect(CODE, '設定欄から読む形になっていない').toMatch(/getProperty\(['"]PRES_PIN['"]\)/);
  });

  it('設定欄に値があれば、それを鍵として読む', () => {
    const g = loadGas({ props: { PRES_PIN: 'せってい済みの鍵' } });
    expect(g.presPin_()).toBe('せってい済みの鍵');
  });

  it('設定欄が空なら鍵は空を返す（値を作らない）', () => {
    expect(loadGas({ props: {} }).presPin_()).toBe('');
    expect(loadGas({ props: { PRES_PIN: '   ' } }).presPin_()).toBe('');
  });

  it('★設定欄が未設定なら、何を送っても通さない（fail-closed）', () => {
    const g = loadGas({ props: {} });
    // 一番危ない形: 送る側も空。'' === '' で素通りする実装だとここで落ちる。
    expect(parse(g.handlePresidentAction_({ pin: '' }, 'pres_list', 'x')).status).toBe('error');
    expect(parse(g.handlePresidentAction_({}, 'pres_list', 'x')).status).toBe('error');
    // ★ここに実在の値を書かない（このファイルは公開リポジトリへ上がる）。
    expect(parse(g.handlePresidentAction_({ pin: 'なにか適当な値' }, 'pres_list', 'x')).status).toBe('error');
  });

  it('設定欄と一致すれば通り、違えば通さない', () => {
    const g = loadGas({ props: { PRES_PIN: '正しい鍵' } });
    expect(parse(g.handlePresidentAction_({ pin: '正しい鍵' }, 'pres_list', 'x')).status).toBe('ok');
    expect(parse(g.handlePresidentAction_({ pin: 'ちがう鍵' }, 'pres_list', 'x')).status).toBe('error');
    expect(parse(g.handlePresidentAction_({ pin: '' }, 'pres_list', 'x')).status).toBe('error');
  });

  it('★書き込み側（追加・変更・削除）も同じ判定で守られている', () => {
    const g = loadGas({ props: {} });
    for (const act of ['pres_add', 'pres_update', 'pres_delete']) {
      expect(parse(g.handlePresidentAction_({ pin: '' }, act, 'x')).status, act).toBe('error');
    }
  });
});
