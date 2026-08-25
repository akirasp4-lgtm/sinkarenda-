import { describe, it, expect } from 'vitest';
import { createRequire } from 'node:module';

// send-queue.js は sync-guard.js と同じUMD風の素のスクリプト。
// 画面（<script src>）とテスト（require）が同一ファイルを見るため、
// 「実装を変えたのにテストは古いまま」が原理的に起きない。
const require = createRequire(import.meta.url);
const SQ = require('../../send-queue.js');

// localStorage の代わり。setItem を失敗させられる。
function makeStorage(opts) {
  const o = opts || {};
  const map = new Map();
  return {
    failSet: !!o.failSet,
    failRemove: !!o.failRemove,
    getItem(k) { return map.has(k) ? map.get(k) : null; },
    setItem(k, v) { if (this.failSet) throw new Error('quota'); map.set(k, String(v)); },
    removeItem(k) { if (this.failRemove) throw new Error('no'); map.delete(k); },
    _map: map
  };
}

const ROW = { id: 'ID-1', date: '2026-09-01', name: '山田', company: 'グローライズ' };
const ITEM = { id: 'ID-1', rows: [ROW], company: 'グローライズ' };

describe('send-queue.js Task1: 箱（永続化・投入・一覧）', () => {
  it('storageが使えるときは isStorageUsable() が true', () => {
    const q = SQ.createSendQueue({ storage: makeStorage(), tabId: 'tab-a' });
    expect(q.isStorageUsable()).toBe(true);
  });

  it('setItem が失敗する storage では isStorageUsable() が false（呼び出し側は従来どおり同期送信に倒す）', () => {
    const q = SQ.createSendQueue({ storage: makeStorage({ failSet: true }), tabId: 'tab-a' });
    expect(q.isStorageUsable()).toBe(false);
  });

  it('storage が null（localStorage自体が無い環境）でも false を返して落ちない', () => {
    const q = SQ.createSendQueue({ storage: null, tabId: 'tab-a' });
    expect(q.isStorageUsable()).toBe(false);
  });

  it('enqueue すると count が増え、list に既定値が入って返る', () => {
    const q = SQ.createSendQueue({ storage: makeStorage(), tabId: 'tab-a' });
    expect(q.enqueue(ITEM, 1000)).toBe(true);
    expect(q.count()).toBe(1);
    const items = q.list();
    expect(items[0].id).toBe('ID-1');
    expect(items[0].createdAt).toBe(1000);
    expect(items[0].attempts).toBe(0);
    expect(items[0].nextAt).toBe(0);
    expect(items[0].gaveUp).toBe(false);
  });

  it('list() は複製を返す（返り値のトップレベルの値を書き換えても箱の中身は変わらない。rowsの深い複製は別テストで検査）', () => {
    const q = SQ.createSendQueue({ storage: makeStorage(), tabId: 'tab-a' });
    q.enqueue(ITEM, 1000);
    q.list()[0].attempts = 999;
    expect(q.list()[0].attempts).toBe(0);
  });

  it('別のタブが作った queue でも、同じ storage を見れば同じ内容が読める（タブ間共有）', () => {
    const st = makeStorage();
    const a = SQ.createSendQueue({ storage: st, tabId: 'tab-a' });
    a.enqueue(ITEM, 1000);
    const b = SQ.createSendQueue({ storage: st, tabId: 'tab-b' });
    expect(b.count()).toBe(1);
  });

  it('★既に読み込み済みのqueueでも、別タブが後から入れた項目が見える（メモリに抱え込まない）', () => {
    const st = makeStorage();
    const a = SQ.createSendQueue({ storage: st, tabId: 'tab-a' });
    expect(a.count()).toBe(0);        // ここで一度読ませる
    const b = SQ.createSendQueue({ storage: st, tabId: 'tab-b' });
    b.enqueue(ITEM, 1000);            // 別タブが後から入れる
    expect(a.count()).toBe(1);        // ← 0 のままだと下のテストで消える
  });

  it('★別タブが後から入れた項目を、先に開いていたタブの書き込みが消さない（未送信の消失防止）', () => {
    const st = makeStorage();
    const a = SQ.createSendQueue({ storage: st, tabId: 'tab-a' });
    expect(a.count()).toBe(0);        // 先に開いていたタブが一度読む
    const b = SQ.createSendQueue({ storage: st, tabId: 'tab-b' });
    b.enqueue({ id: 'B-1', rows: [ROW], company: 'X' }, 1000);
    a.enqueue({ id: 'A-1', rows: [ROW], company: 'X' }, 2000);
    const ids = SQ.createSendQueue({ storage: st, tabId: 'tab-c' }).list().map(x => x.id);
    expect(ids.sort()).toEqual(['A-1', 'B-1']);   // どちらも残っていること
  });

  it('storageが使えないときはメモリで動く（このタブの中では送信を続けられる）', () => {
    const q = SQ.createSendQueue({ storage: makeStorage({ failSet: true }), tabId: 'tab-a' });
    expect(q.isStorageUsable()).toBe(false);
    expect(q.enqueue(ITEM, 1000)).toBe(true);
    expect(q.count()).toBe(1);
  });

  it('maxItems を超えたら enqueue が false を返す（呼び出し側は従来どおり同期送信に倒す）', () => {
    const q = SQ.createSendQueue({ storage: makeStorage(), tabId: 'tab-a', maxItems: 2 });
    expect(q.enqueue({ id: 'A', rows: [ROW], company: 'X' }, 1)).toBe(true);
    expect(q.enqueue({ id: 'B', rows: [ROW], company: 'X' }, 2)).toBe(true);
    expect(q.enqueue({ id: 'C', rows: [ROW], company: 'X' }, 3)).toBe(false);
    expect(q.count()).toBe(2);
  });

  it('rows が空／id が空の項目は enqueue を拒否する（壊れた項目を貯めない）', () => {
    const q = SQ.createSendQueue({ storage: makeStorage(), tabId: 'tab-a' });
    expect(q.enqueue({ id: '', rows: [ROW], company: 'X' }, 1)).toBe(false);
    expect(q.enqueue({ id: 'A', rows: [], company: 'X' }, 1)).toBe(false);
    expect(q.enqueue(null, 1)).toBe(false);
    expect(q.count()).toBe(0);
  });

  it('storageの値がJSON文字列として壊れていて丸ごとparseに失敗しても、空として扱い落ちない（下の「壊れた項目」テストとは別ケース＝JSON.parse自体の失敗）', () => {
    const st = makeStorage();
    st._map.set('yotei-pending-add-v1', '{壊れ');
    const q = SQ.createSendQueue({ storage: st, tabId: 'tab-a' });
    expect(q.count()).toBe(0);
    expect(q.enqueue(ITEM, 1)).toBe(true);
  });

  it('★Critical修正: 壊れた項目（null・空文字id・文字列・rowsが配列でない）が混ざっていても list()/pendingRows()/count() が落ちず、正常な項目だけが残る', () => {
    const st = makeStorage();
    const good = {
      id: 'GOOD-1', rows: [ROW], company: 'グローライズ',
      createdAt: 1, attempts: 0, nextAt: 0, lastError: '', gaveUp: false,
      owner: 'tab-a', claimedAt: 0
    };
    st._map.set('yotei-pending-add-v1', JSON.stringify({
      v: 1,
      items: [null, { id: '', rows: [] }, '文字列', { id: 'BAD-2', rows: 'notarray' }, good]
    }));
    const q = SQ.createSendQueue({ storage: st, tabId: 'tab-a' });
    expect(() => q.count()).not.toThrow();
    expect(() => q.list()).not.toThrow();
    expect(() => q.pendingRows()).not.toThrow();
    expect(q.count()).toBe(1);
    expect(q.list().map(x => x.id)).toEqual(['GOOD-1']);
    expect(q.pendingRows().length).toBe(1);
  });

  it('★Important1修正: mutate()は_internalsとして公開され、自分が読んだ後に別インスタンスがenqueueした項目を、自分のenqueueが上書きしない', () => {
    const st = makeStorage();
    const a = SQ.createSendQueue({ storage: st, tabId: 'tab-a' });
    expect(typeof a._internals.mutate).toBe('function'); // Task2以降が使う入口が公開されていること
    a.count(); // 先に一度読ませておく
    const b = SQ.createSendQueue({ storage: st, tabId: 'tab-b' });
    b.enqueue({ id: 'B-1', rows: [ROW], company: 'X' }, 1000);
    expect(a.enqueue({ id: 'A-1', rows: [ROW], company: 'X' }, 2000)).toBe(true);
    const ids = SQ.createSendQueue({ storage: st, tabId: 'tab-c' }).list().map(x => x.id);
    expect(ids.sort()).toEqual(['A-1', 'B-1']);
  });

  it('★Important2修正(a): storage不使用（メモリ）モードでも、enqueueに渡したrowsを後から書き換えても箱の中身は変わらない', () => {
    const q = SQ.createSendQueue({ storage: null, tabId: 'tab-a' });
    const row = { id: 'ID-1', date: '2026-09-01', name: '山田', company: 'グローライズ' };
    expect(q.enqueue({ id: 'X1', rows: [row], company: 'Y' }, 1000)).toBe(true);
    row.name = 'CORRUPTED';
    expect(q.list()[0].rows[0].name).toBe('山田');
  });

  it('★Important2修正(b): storage不使用（メモリ）モードでも、list()の返り値のrows[0]を書き換えても次のlist()に影響しない', () => {
    const q = SQ.createSendQueue({ storage: null, tabId: 'tab-a' });
    q.enqueue(ITEM, 1000);
    q.list()[0].rows[0].name = 'CORRUPTED';
    expect(q.list()[0].rows[0].name).toBe('山田');
  });

  it('pendingRows(company) は会社が一致する項目の rows だけを平坦化して返す', () => {
    const q = SQ.createSendQueue({ storage: makeStorage(), tabId: 'tab-a' });
    q.enqueue({ id: 'A', rows: [ROW, ROW], company: 'グローライズ' }, 1);
    q.enqueue({ id: 'B', rows: [ROW], company: '和信カインド' }, 2);
    expect(q.pendingRows('グローライズ').length).toBe(2);
    expect(q.pendingRows('和信カインド').length).toBe(1);
    expect(q.pendingRows().length).toBe(3);
  });
});
