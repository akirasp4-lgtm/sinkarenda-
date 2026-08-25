import { describe, it, expect } from 'vitest';
import { createRequire } from 'node:module';

// send-queue.js は sync-guard.js と同じUMD風の素のスクリプト。
// 画面（<script src>）とテスト（require）が同一ファイルを見るため、
// 「実装を変えたのにテストは古いまま」が原理的に起きない。
const require = createRequire(import.meta.url);
const SQ = require('../../send-queue.js');

// localStorage の代わり。setItem を失敗させられる。
// ★修正ラウンド2: 本物の localStorage と同じ形で length / key(i) を持たせる
// （send-queue.js が「1件＝1キー」の走査に storage.length / storage.key(i) を
// 使うようになったため）。Map の挿入順で返せば十分。
function makeStorage(opts) {
  const o = opts || {};
  const map = new Map();
  return {
    failSet: !!o.failSet,
    failRemove: !!o.failRemove,
    getItem(k) { return map.has(k) ? map.get(k) : null; },
    setItem(k, v) { if (this.failSet) throw new Error('quota'); map.set(k, String(v)); },
    removeItem(k) { if (this.failRemove) throw new Error('no'); map.delete(k); },
    get length() { return map.size; },
    key(i) {
      const keys = Array.from(map.keys());
      return i >= 0 && i < keys.length ? keys[i] : null;
    },
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

  // ★修正ラウンド2・変更1で「1つのキーに全項目をまとめる」形をやめたため、
  // このテストは「1件のキーの値が壊れているケース」に作り替えた。
  it('1件のキーの値がJSON文字列として壊れていても、そのキーだけ読み飛ばして落ちない（他は正常に動く。壊れたキー自体は消さない）', () => {
    const st = makeStorage();
    st._map.set('yotei-pending-add-v1:BAD-1', '{壊れ');
    const q = SQ.createSendQueue({ storage: st, tabId: 'tab-a' });
    expect(q.count()).toBe(0);
    expect(q.enqueue(ITEM, 1)).toBe(true);
    expect(q.count()).toBe(1);
    expect(st._map.has('yotei-pending-add-v1:BAD-1')).toBe(true); // 消えていない
  });

  it('★Critical修正: 壊れた項目（JSON破損・空文字id・オブジェクトでない・rowsが配列でない）が別々のキーに混ざっていても list()/pendingRows()/count() が落ちず、正常な項目だけが残る（壊れたキー自体は消さない）', () => {
    const st = makeStorage();
    const good = {
      id: 'GOOD-1', rows: [ROW], company: 'グローライズ',
      createdAt: 1, attempts: 0, nextAt: 0, lastError: '', gaveUp: false,
      owner: 'tab-a', claimedAt: 0
    };
    st._map.set('yotei-pending-add-v1:BAD-1', '{壊れ');                                   // JSON破損
    st._map.set('yotei-pending-add-v1:BAD-2', JSON.stringify({ id: '', rows: [] }));      // id空
    st._map.set('yotei-pending-add-v1:BAD-3', JSON.stringify('文字列'));                   // オブジェクトでない
    st._map.set('yotei-pending-add-v1:BAD-4', JSON.stringify({ id: 'BAD-4', rows: 'x' })); // rowsが配列でない
    st._map.set('yotei-pending-add-v1:GOOD-1', JSON.stringify(good));
    const q = SQ.createSendQueue({ storage: st, tabId: 'tab-a' });
    expect(() => q.count()).not.toThrow();
    expect(() => q.list()).not.toThrow();
    expect(() => q.pendingRows()).not.toThrow();
    expect(q.count()).toBe(1);
    expect(q.list().map(x => x.id)).toEqual(['GOOD-1']);
    expect(q.pendingRows().length).toBe(1);
    // ★修正ラウンド2・変更1: 壊れたキーは読み飛ばすだけで、消してはいけない
    expect(st._map.has('yotei-pending-add-v1:BAD-1')).toBe(true);
    expect(st._map.has('yotei-pending-add-v1:BAD-2')).toBe(true);
    expect(st._map.has('yotei-pending-add-v1:BAD-3')).toBe(true);
    expect(st._map.has('yotei-pending-add-v1:BAD-4')).toBe(true);
  });

  // ★修正ラウンド2・変更1で mutate()（丸ごと読み直して丸ごと書き戻す仕組み）自体を
  // 廃止したため、_internals の検査対象を1件単位の入出力に置き換えた。
  // 「自分が読んだ後に別インスタンスが入れた項目を、自分のenqueueが上書きしない」
  // という検査内容自体は、1件＝1キーになったことで構造的に保証されるようになったが、
  // 回帰を防ぐため検査は残す。
  it('★Important1修正の引き継ぎ: _internalsに1件単位の入出力（readItem/writeItem/deleteItem/listAllItems）が公開され、自分が読んだ後に別インスタンスが入れた項目を、自分のenqueueが上書きしない', () => {
    const st = makeStorage();
    const a = SQ.createSendQueue({ storage: st, tabId: 'tab-a' });
    expect(typeof a._internals.readItem).toBe('function');
    expect(typeof a._internals.writeItem).toBe('function');
    expect(typeof a._internals.deleteItem).toBe('function');
    expect(typeof a._internals.listAllItems).toBe('function');
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

  // ★修正ラウンド2・変更3: このオリジンには別の大きなキャッシュ（実測700KB）も
  // 同居している。走査は storageKey + ':' の前方一致に厳密に限定し、他のキーを
  // 拾わないこと・触らないことを検査する。
  it('★変更3: 無関係なキー（別アプリのキャッシュ・旧形式の1キーまとめ）を list() が拾わない', () => {
    const st = makeStorage();
    const cacheValue = JSON.stringify({ huge: 'unrelated cache data' });
    st._map.set('yotei-cache-v1', cacheValue);
    st._map.set('yotei-pending-add-v1', JSON.stringify({ v: 1, items: [ITEM] })); // 旧形式そのもの
    const q = SQ.createSendQueue({ storage: st, tabId: 'tab-a' });
    expect(q.count()).toBe(0);
    expect(q.list()).toEqual([]);
    expect(q.enqueue(ITEM, 1000)).toBe(true);
    expect(q.count()).toBe(1);
    // 無関係なキー・旧形式のキーは書き換えられても消されてもいない
    expect(st._map.get('yotei-cache-v1')).toBe(cacheValue);
    expect(st._map.has('yotei-pending-add-v1')).toBe(true);
  });

  it('★変更3: 走査は prefix 前方一致のキーだけを読み、無関係なキーを読んだり消したりしない', () => {
    const st = makeStorage();
    st._map.set('yotei-cache-v1', JSON.stringify({ huge: '...' }));
    st._map.set('yotei-pending-add-v1', JSON.stringify({ v: 1, items: [ITEM] }));
    st._map.set('yotei-pending-add-v1:ID-1', JSON.stringify({
      id: 'ID-1', rows: [ROW], company: 'グローライズ',
      createdAt: 1000, attempts: 0, nextAt: 0, lastError: '', gaveUp: false,
      owner: 'tab-a', claimedAt: 0
    }));

    const touchedGet = [];
    const realGetItem = st.getItem.bind(st);
    st.getItem = function (k) { touchedGet.push(k); return realGetItem(k); };
    const removed = [];
    const realRemoveItem = st.removeItem.bind(st);
    st.removeItem = function (k) { removed.push(k); return realRemoveItem(k); };

    const q = SQ.createSendQueue({ storage: st, tabId: 'tab-a' });
    expect(q.list().map(x => x.id)).toEqual(['ID-1']);
    expect(touchedGet).not.toContain('yotei-cache-v1');
    expect(touchedGet).not.toContain('yotei-pending-add-v1');
    expect(removed).not.toContain('yotei-cache-v1');
    expect(removed).not.toContain('yotei-pending-add-v1');
  });
});

describe('send-queue.js Task2: 送信権・バックオフ・諦め', () => {
  it('enqueue した直後は nextDue で取れる', () => {
    const q = SQ.createSendQueue({ storage: makeStorage(), tabId: 'tab-a' });
    q.enqueue(ITEM, 1000);
    expect(q.nextDue(1000).id).toBe('ID-1');
  });

  it('beginSend は fetch の前に attempts を +1 して永続化する（D9の本体）', () => {
    const st = makeStorage();
    const q = SQ.createSendQueue({ storage: st, tabId: 'tab-a' });
    q.enqueue(ITEM, 1000);
    const r = q.beginSend('ID-1', 1000);
    expect(r).not.toBeNull();
    expect(r.wasRetry).toBe(false);
    // 別インスタンス＝storageから読み直しても attempts が1になっている
    const other = SQ.createSendQueue({ storage: st, tabId: 'tab-b' });
    expect(other.list()[0].attempts).toBe(1);
  });

  it('★送信中に落ちて失敗の記録が残らなくても、次は再送扱い（wasRetry:true）になる', () => {
    const st = makeStorage();
    const a = SQ.createSendQueue({ storage: st, tabId: 'tab-a' });
    a.enqueue(ITEM, 1000);
    a.beginSend('ID-1', 1000);   // ここでタブが落ちた想定（markSent も markFailed も呼ばれない）
    const b = SQ.createSendQueue({ storage: st, tabId: 'tab-b' });
    const r = b.beginSend('ID-1', 1000 + 31000); // リース切れ後に別タブが拾う
    expect(r).not.toBeNull();
    expect(r.wasRetry).toBe(true); // ← ここが false だと二重登録になる
  });

  it('attempts の永続化に失敗したら beginSend は null を返す（記録できないなら送らない）', () => {
    const st = makeStorage();
    const q = SQ.createSendQueue({ storage: st, tabId: 'tab-a' });
    q.enqueue(ITEM, 1000);
    st.failSet = true;
    expect(q.beginSend('ID-1', 1000)).toBeNull();
  });

  it('2つのタブが同時に beginSend しても、送信権を取れるのは片方だけ', () => {
    const st = makeStorage();
    const a = SQ.createSendQueue({ storage: st, tabId: 'tab-a' });
    const b = SQ.createSendQueue({ storage: st, tabId: 'tab-b' });
    a.enqueue(ITEM, 1000);
    const ra = a.beginSend('ID-1', 1000);
    const rb = b.beginSend('ID-1', 1000);
    expect(ra).not.toBeNull();
    expect(rb).toBeNull();
  });

  it('初回送信は enqueue したタブだけが行う（他タブはリース経過まで拾わない）', () => {
    const st = makeStorage();
    const a = SQ.createSendQueue({ storage: st, tabId: 'tab-a' });
    a.enqueue(ITEM, 1000);
    const b = SQ.createSendQueue({ storage: st, tabId: 'tab-b' });
    expect(b.nextDue(1000)).toBeNull();             // まだ拾わない
    expect(b.nextDue(1000 + 31000)).not.toBeNull(); // リース経過後は拾う
  });

  it('markSent は箱から消す。token が違えば消さない（古い試行が新しい試行を消せない）', () => {
    const q = SQ.createSendQueue({ storage: makeStorage(), tabId: 'tab-a' });
    q.enqueue(ITEM, 1000);
    const r = q.beginSend('ID-1', 1000);
    expect(q.markSent('ID-1', 'ちがうtoken')).toBe(false);
    expect(q.count()).toBe(1);
    expect(q.markSent('ID-1', r.token)).toBe(true);
    expect(q.count()).toBe(0);
  });

  it('markFailed でバックオフが 5秒→15秒→45秒→2分→5分 と伸び、5分で頭打ちになる', () => {
    // ★時刻を進めながら回すこと。markFailed が nextAt を未来に置くため、
    // now を進めずに beginSend を呼ぶと isDue が false になり null が返る。
    const q = SQ.createSendQueue({ storage: makeStorage(), tabId: 'tab-a', giveUpAfter: 99 });
    q.enqueue(ITEM, 0);
    const waits = [];
    let now = 0;
    for (let i = 0; i < 7; i++) {
      const r = q.beginSend('ID-1', now);
      expect(r).not.toBeNull();
      q.markFailed('ID-1', r.token, 'えらー', now);
      const nextAt = q.list()[0].nextAt;
      waits.push(nextAt - now);
      now = nextAt;   // バックオフ明けまで進める
    }
    expect(waits).toEqual([5000, 15000, 45000, 120000, 300000, 300000, 300000]);
  });

  it('バックオフ中は nextDue に出てこない', () => {
    const q = SQ.createSendQueue({ storage: makeStorage(), tabId: 'tab-a' });
    q.enqueue(ITEM, 0);
    const r = q.beginSend('ID-1', 0);
    q.markFailed('ID-1', r.token, 'えらー', 0);
    expect(q.nextDue(4999)).toBeNull();
    expect(q.nextDue(5000)).not.toBeNull();
  });

  it('giveUpAfter 回失敗したら自動再送を止める（勝手に捨てない・件数は残る）', () => {
    const q = SQ.createSendQueue({ storage: makeStorage(), tabId: 'tab-a', giveUpAfter: 3 });
    q.enqueue(ITEM, 0);
    let now = 0;
    for (let i = 0; i < 3; i++) {
      const r = q.beginSend('ID-1', now);   // ★バックオフ明けまで now を進める
      expect(r).not.toBeNull();
      q.markFailed('ID-1', r.token, 'えらー', now);
      now = q.list()[0].nextAt;
    }
    expect(q.list()[0].gaveUp).toBe(true);
    expect(q.gaveUpCount()).toBe(1);
    expect(q.count()).toBe(1);                    // 捨てていない
    expect(q.nextDue(999999999)).toBeNull();      // 自動では拾わない
  });

  it('retryNow で諦めた項目を手動で送り直せる', () => {
    const q = SQ.createSendQueue({ storage: makeStorage(), tabId: 'tab-a', giveUpAfter: 1 });
    q.enqueue(ITEM, 0);
    const r = q.beginSend('ID-1', 0);
    q.markFailed('ID-1', r.token, 'えらー', 0);
    expect(q.nextDue(999999999)).toBeNull();
    expect(q.retryNow('ID-1', 0)).toBe(true);
    expect(q.nextDue(0)).not.toBeNull();
  });

  it('存在しないIDへの操作は false / null を返して落ちない', () => {
    const q = SQ.createSendQueue({ storage: makeStorage(), tabId: 'tab-a' });
    expect(q.beginSend('ない', 0)).toBeNull();
    expect(q.markSent('ない', 't')).toBe(false);
    expect(q.markFailed('ない', 't', 'e', 0)).toBe(false);
    expect(q.retryNow('ない', 0)).toBe(false);
  });

  // ★修正ラウンド2・変更1: 保存の形を「1つのキーに全項目」から「1件＝1キー」に
  // 変えたことで、修正ラウンド1のCritical C-1（別タブの古いスナップショットの
  // 丸ごと書き戻しによる巻き戻り）はそもそも成立しなくなった（タブBが別項目を
  // 書いても、タブAの項目のキーには構造的に触れないため）。この構造的な効果を
  // 直接確認する。
  it('★変更1の構造確認: 別タブが別項目を同時に書いても、キーが分かれているため互いに一切影響しない', () => {
    const st = makeStorage();
    const a = SQ.createSendQueue({ storage: st, tabId: 'tab-a' });
    a.enqueue(ITEM, 1000); // ID-1
    const b = SQ.createSendQueue({ storage: st, tabId: 'tab-b' });
    const ITEM2 = { id: 'ID-2', rows: [ROW], company: 'グローライズ' };

    const realSetItem = st.setItem.bind(st);
    st.setItem = function (k, v) {
      if (k === 'yotei-pending-add-v1:ID-1') {
        // タブAがID-1のキーへ書いている「最中」に、タブBがID-2を新規登録する
        b.enqueue(ITEM2, 2000);
      }
      realSetItem(k, v);
    };

    const r = a.beginSend('ID-1', 1000);
    expect(r).not.toBeNull(); // ID-2側の割り込みはID-1に一切影響しない
    expect(r.wasRetry).toBe(false);
    const c = SQ.createSendQueue({ storage: st, tabId: 'tab-c' });
    expect(c.list().map(x => x.id).sort()).toEqual(['ID-1', 'ID-2']); // 両方残っている
  });

  // ★修正ラウンド1・Critical C-1（設計書D2(b)）の読み直し確認は、1件キーの
  // 世界でも「同じキー」への競合が起きた場合に備えて残してある（localStorage
  // にはCASが無いため、既存のリース・初回所有権をすり抜けた場合の保険）。
  // この保険がまだ働くことを、同じ項目キーへの割り込みで確認する。
  it('★C-1（1件キーの世界に引き継ぎ）: 同じ項目キーへの書き込みが割り込んでも、beginSendは読み直しで検知してnullを返す', () => {
    const st = makeStorage();
    const a = SQ.createSendQueue({ storage: st, tabId: 'tab-a' });
    a.enqueue(ITEM, 1000); // ID-1: attempts=0, owner=tab-a
    const b = SQ.createSendQueue({ storage: st, tabId: 'tab-b' });

    const realSetItem = st.setItem.bind(st);
    let writeCount = 0;
    st.setItem = function (k, v) {
      realSetItem(k, v);
      writeCount++;
      if (writeCount === 1 && k === 'yotei-pending-add-v1:ID-1') {
        // タブAがID-1のキーへ書いた直後、タブBが同じキーへ古い内容（attempts=0）を
        // 割り込ませて書く（localStorageにCASが無いため、同一キーへの競合は
        // 依然として残る）
        b._internals.writeItem({
          id: 'ID-1', rows: [ROW], company: 'グローライズ', createdAt: 1000,
          attempts: 0, nextAt: 0, lastError: '', gaveUp: false,
          owner: 'tab-b', claimedAt: 0
        });
      }
    };

    const r = a.beginSend('ID-1', 1000);
    expect(r).toBeNull(); // 巻き戻りを読み直しで検知し、送らせない
  });

  // ★修正ラウンド2・変更2: 書き込み失敗時に removeItem していたのをやめた。
  // 未送信そのものを全消去してしまう致命的な副作用だったため。
  it('★変更2: 書き込みが失敗しても既存キーを消さない。storageが回復すれば古い内容がそのまま読まれて送信される（救済）', () => {
    const st = makeStorage();
    const a = SQ.createSendQueue({ storage: st, tabId: 'tab-a' });
    a.enqueue(ITEM, 1000);
    expect(st._map.has('yotei-pending-add-v1:ID-1')).toBe(true);

    st.failSet = true;
    a.retryNow('ID-1', 1000); // 書き込み失敗 → usable=false になるが、キーは消えない
    expect(a.isStorageUsable()).toBe(false);
    expect(st._map.has('yotei-pending-add-v1:ID-1')).toBe(true); // ★消えていない
    const raw = JSON.parse(st._map.get('yotei-pending-add-v1:ID-1'));
    expect(raw.id).toBe('ID-1'); // 失敗前の内容がそのまま残っている

    st.failSet = false; // quotaが回復
    const b = SQ.createSendQueue({ storage: st, tabId: 'tab-b' }); // 次にページを開いた想定
    expect(b.list().map(x => x.id)).toEqual(['ID-1']); // 救済＝正常に読まれる
  });

  // ★修正ラウンド1・Important I-2（メモリ運転に落ちたタブが storage を上書きして
  // 別タブの未送信を消す）を、変更1・変更2を踏まえて引き継ぐ。1件＝1キーになった
  // ことで、タブAが usable=false のまま書いても構造的に自分のキーにしか触れない
  // （タブBの ID-2 のキーへ触りようがない）。さらに変更2により、タブA自身の
  // ID-1 のキーも消えずに残る。旧テストは「タブBの項目だけが残る」ことを検査
  // していたが、変更2によって「両方残る」に改善されたため、その通りに検査する。
  it('★I-2の引き継ぎ（変更1・変更2により強化）: usable=falseに落ちたタブが後から書いても、別タブの未送信どころか自分の未送信も消えない', () => {
    const st = makeStorage();
    const a = SQ.createSendQueue({ storage: st, tabId: 'tab-a' });
    a.enqueue(ITEM, 1000); // ID-1
    st.failSet = true;
    a.retryNow('ID-1', 1000); // 書き込み失敗 → usable=false・メモリ運転へ
    expect(a.isStorageUsable()).toBe(false);
    expect(st._map.has('yotei-pending-add-v1:ID-1')).toBe(true); // 消えていない（変更2）

    st.failSet = false; // quotaが回復
    const b = SQ.createSendQueue({ storage: st, tabId: 'tab-b' });
    const ITEM2 = { id: 'ID-2', rows: [ROW], company: 'グローライズ' };
    expect(b.enqueue(ITEM2, 2000)).toBe(true);

    a.beginSend('ID-1', 3000); // usableが戻っていないタブAは、ID-1のキーにしか触れようがない

    const c = SQ.createSendQueue({ storage: st, tabId: 'tab-c' });
    expect(c.list().map(x => x.id).sort()).toEqual(['ID-1', 'ID-2']); // ★両方残っている
  });

  // ★修正ラウンド1・Minor M-2: token不一致のテストはmarkSentにしかなかった。
  // markFailedでも「古い試行がnextAt/claimedAt/lastErrorを触れない」ことを固定する。
  it('markFailed は token が違えば nextAt/claimedAt/lastError を書き換えない（古い試行が新しい試行の状態を上書きしない）', () => {
    const q = SQ.createSendQueue({ storage: makeStorage(), tabId: 'tab-a' });
    q.enqueue(ITEM, 1000);
    q.beginSend('ID-1', 1000); // attempts=1, token=tab-a:1
    const before = q.list()[0];
    expect(q.markFailed('ID-1', 'ちがうtoken', 'えらー', 2000)).toBe(false);
    const after = q.list()[0];
    expect(after.nextAt).toBe(before.nextAt);
    expect(after.claimedAt).toBe(before.claimedAt);
    expect(after.lastError).toBe(before.lastError);
    expect(after.gaveUp).toBe(before.gaveUp);
  });
});
