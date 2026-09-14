// GR JARVIS 連携用の「読むだけの鍵」（2026-09-14）
//
// ★社長の依頼:
//   「まずは安全のため、JARVISから予定を書き換えるんじゃなくて読取専用で接続して、
//     正常に同期できることを確認してから、予定追加・変更まで広げようと思ってる」
//
// ★このテストが守るもの（壊れたら赤くなる）:
//   1. 読むだけの鍵で pres_add / pres_update / pres_delete が**絶対に通らない**
//   2. 鍵が未設定なら誰も通さない（fail-closed）
//   3. 空の鍵を送っても通らない（'' === '' の素通りを防ぐ）
//   4. 読むだけの鍵では社長の4桁を要求しない（PINを渡さずに済む＝依頼の主旨）
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const CODE = readFileSync(join(here, '..', '..', 'gas.js'), 'utf8');

function load(props) {
  const store = Object.assign({}, props);
  const sandbox = vm.createContext({
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({
        getSheetByName: () => ({
          getDataRange: () => ({
            getValues: () => ([
              ['登録日時', 'タイトル', '開始日', '開始時刻', '終了日', '終了時刻', '場所', 'メモ', 'カテゴリ', '色', 'ID', '更新者'],
              ['2026-09-10', '打合せ', '2026-09-16', '13:30', '2026-09-16', '15:00', '本社', '見積', '商談', '#1D9E75', 'Pabc', 'president']
            ])
          })
        })
      })
    },
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
    LockService: { getUserLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    Utilities: { formatDate: (d) => String(d) },
    ContentService: {
      createTextOutput: (s) => ({ setMimeType: () => ({ _body: s }) }),
      MimeType: { JSON: 'json' }
    },
    UrlFetchApp: {},
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => (k in store ? store[k] : null) }) },
    Logger: { log() {} }, console
  });
  vm.runInContext(CODE + ';globalThis.__p = { handlePresidentAction_, presReadOnlyOk_, presReadToken_ };',
    sandbox, { filename: 'gas.js' });
  return sandbox.__p;
}

const parse = (res) => JSON.parse(res._body);

describe('読むだけの鍵は一覧取得にしか効かない', () => {
  const P = load({ PRES_PIN: '9999', PRES_RO_TOKEN: 'ro-secret' });

  it('正しい鍵で一覧は取れる（社長の4桁は渡していない）', () => {
    const r = parse(P.handlePresidentAction_({ ro_token: 'ro-secret' }, 'pres_list', 'jarvis'));
    expect(r.status).toBe('ok');
    expect(r.readOnly).toBe(true);
    expect(Array.isArray(r.rows)).toBe(true);
  });

  // ★ここが本丸。鍵が漏れても予定を壊せないことを保証する。
  it('★読むだけの鍵では 追加 が通らない', () => {
    const r = parse(P.handlePresidentAction_({ ro_token: 'ro-secret' }, 'pres_add', 'jarvis'));
    expect(r.status).toBe('error');
    expect(r.message).toContain('認証');
  });

  it('★読むだけの鍵では 変更 が通らない', () => {
    const r = parse(P.handlePresidentAction_({ ro_token: 'ro-secret' }, 'pres_update', 'jarvis'));
    expect(r.status).toBe('error');
  });

  it('★読むだけの鍵では 削除 が通らない', () => {
    const r = parse(P.handlePresidentAction_({ ro_token: 'ro-secret' }, 'pres_delete', 'jarvis'));
    expect(r.status).toBe('error');
  });

  it('間違った鍵では一覧も取れない', () => {
    const r = parse(P.handlePresidentAction_({ ro_token: 'chigau' }, 'pres_list', 'x'));
    expect(r.status).toBe('error');
  });

  it('鍵を送らなければ従来どおり4桁が要る', () => {
    expect(parse(P.handlePresidentAction_({}, 'pres_list', 'x')).status).toBe('error');
    expect(parse(P.handlePresidentAction_({ pin: '9999' }, 'pres_list', 'x')).status).toBe('ok');
  });
});

describe('鍵が未設定なら誰も通さない（fail-closed）', () => {
  const P = load({ PRES_PIN: '9999' });   // PRES_RO_TOKEN を入れていない

  it('未設定のとき、鍵を送っても通らない', () => {
    expect(parse(P.handlePresidentAction_({ ro_token: 'nandemo' }, 'pres_list', 'x')).status).toBe('error');
  });

  // ★空文字どうしの比較で素通りする事故を防ぐ。PRES_PIN で実際に踏んだ形。
  it('★未設定のとき、空の鍵でも通らない', () => {
    expect(P.presReadOnlyOk_({ ro_token: '' }, 'pres_list')).toBe(false);
    expect(parse(P.handlePresidentAction_({ ro_token: '' }, 'pres_list', 'x')).status).toBe('error');
  });

  it('設定してあっても、空の鍵は通らない', () => {
    const P2 = load({ PRES_PIN: '9999', PRES_RO_TOKEN: 'ro-secret' });
    expect(P2.presReadOnlyOk_({ ro_token: '' }, 'pres_list')).toBe(false);
    expect(P2.presReadOnlyOk_({}, 'pres_list')).toBe(false);
  });
});

describe('Worker側の読取専用の窓口', () => {
  const SRC = readFileSync(join(here, '..', 'src', 'index.js'), 'utf8');

  it('窓口がある', () => {
    expect(SRC).toContain("url.pathname === '/api/president-readonly'");
  });

  it('★書き込みの窓口では読取専用の鍵を使っていない', () => {
    const i = SRC.indexOf("url.pathname === '/api/pres-sync'");
    const body = SRC.slice(i, i + 600);
    expect(body).toContain('checkPresPin');
    expect(body).not.toContain('checkPresReadToken');
  });

  it('鍵が未設定なら503で止める（fail-closed）', () => {
    const i = SRC.indexOf('async function checkPresReadToken');
    const body = SRC.slice(i, SRC.indexOf('\n}', i));
    expect(body).toContain('PRES_RO_TOKENが未設定');
    expect(body).toContain("given === ''");
  });

  it('項目名を英語にし Asia/Tokyo で組み立てて返す', () => {
    const i = SRC.indexOf("url.pathname === '/api/president-readonly'");
    const body = SRC.slice(i, i + 3200);
    expect(body).toContain("timeZone: 'Asia/Tokyo'");
    expect(body).toContain("+09:00");
    expect(body).toContain('allDay');
    expect(body).toContain("id:");
    expect(body).toContain("title:");
    expect(body).toContain("location:");
    expect(body).toContain("notes:");
  });

  it('★終日予定（開始時刻が空）を allDay として返す', () => {
    const i = SRC.indexOf("url.pathname === '/api/president-readonly'");
    const body = SRC.slice(i, i + 3200);
    expect(body).toContain('const allDay = !st;');
  });
});
