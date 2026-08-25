# 案A 保存の楽観的表示＋裏送信 実装計画

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 予定の新規登録を「押した瞬間に画面へ出す」形にし、送信は裏で行う。体感4.4秒→0秒。

**Architecture:** 未送信の登録を localStorage の箱（`send-queue.js`＝純ロジック）に貯め、画面側の送信係が1件ずつGASへ送る。2回目以降の送信では、送る直前にGASから最新を読んで「もう入っていないか」を確認してから送る（二重登録の防止）。`attempts` は送信の**前**に永続化する。

**Tech Stack:** 素のJavaScript（ビルド無し）、localStorage、vitest（`cf/` 配下）

**Spec:** `docs/superpowers/specs/2026-08-25-optimistic-save-design.md`

## Global Constraints

- **計画書に書いてある行番号は「元のファイル」の位置であり、目安でしかない。**
  Task 3 が `index.html` に行を挿入するため、Task 4 以降では実際の行番号がずれる。
  **必ず引用してあるコードの中身で場所を特定すること**（行番号で探して当てない）

- **`gas.js` を1行も変更しない。** Cloudflare（`cf/src/`）・D1・`backend.json` も変更しない
- 触ってよいのは `send-queue.js`（新規）/ `index.html` / `admin.html` / `cf/test/send-queue.test.js`（新規）のみ
- **`index.html` と `admin.html` の変更は1文字も差があってはならない**（過去のレビューで両画面の乖離が繰り返し重大指摘になっている。行番号だけが違う）
- `send-queue.js` は `sync-guard.js` と同じUMD風の作り（`import`/`export` 構文を使わない素のスクリプト。ブラウザでは `<script src>`、Nodeでは `require`）
- **既存テスト153件を1件も壊さない**
- 既存の安全装置を壊さない（起動時キャッシュ / 失敗時に前回内容を保持 / `dataLoadOk` ガード / 会社切替時の競合防止 / 鮮度ガード / preferGasTracker / `timeoutSignal` のフォールバック / インラインフォールバック）
- `git push` はしない（利用者が判断する）
- 定数の既定値: `storageKey='yotei-pending-add-v1'` / `maxItems=50` / `leaseMs=30000` / `giveUpAfter=10` / `backoffMs=[5000,15000,45000,120000,300000]`

---

### Task 1: `send-queue.js` の箱（永続化・投入・一覧）

**Files:**
- Create: `send-queue.js`
- Test: `cf/test/send-queue.test.js`

**Interfaces:**
- Consumes: なし
- Produces:
  - `createSendQueue(opts)` → queue オブジェクト
    - `opts = { storage, storageKey, tabId, maxItems, leaseMs, giveUpAfter, backoffMs }`
    - `storage` は localStorage 互換（`getItem`/`setItem`/`removeItem`）。テストから差し替える
  - `queue.isStorageUsable()` → `boolean`
  - `queue.enqueue(item, now)` → `boolean`（`item = {id, rows, company}`）
  - `queue.list()` → `Array`（保存されている項目の複製）
  - `queue.count()` → `number`
  - `queue.pendingRows(company)` → `Array`（`company` が一致する項目の `rows` を平坦化。`company` 省略で全件）
  - 保存形式: `{ v:1, items:[ {id, rows, company, createdAt, attempts, nextAt, lastError, gaveUp, owner, claimedAt} ] }`

- [ ] **Step 1: 失敗するテストを書く**

`cf/test/send-queue.test.js` を新規作成:

```javascript
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

  it('list() は複製を返す（返り値を書き換えても箱の中身は変わらない）', () => {
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

  it('壊れたJSONが storage に入っていても空として扱い、落ちない', () => {
    const st = makeStorage();
    st._map.set('yotei-pending-add-v1', '{壊れ');
    const q = SQ.createSendQueue({ storage: st, tabId: 'tab-a' });
    expect(q.count()).toBe(0);
    expect(q.enqueue(ITEM, 1)).toBe(true);
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
```

- [ ] **Step 2: テストを走らせて失敗することを確認**

Run: `cd cf && npx vitest run test/send-queue.test.js`
Expected: FAIL（`Cannot find module '../../send-queue.js'`）

- [ ] **Step 3: `send-queue.js` を書く**

`sync-guard.js` と同じ外枠にする（ファイル末尾の `module.exports` / グローバル代入の形をそのまま真似る）。

```javascript
// send-queue.js — 未送信の登録を貯める箱（案A: 楽観的表示＋裏送信）
//
// 設計書: docs/superpowers/specs/2026-08-25-optimistic-save-design.md
//
// この箱は「次にどれを送るか」だけを判断する純ロジックで、DOMにもfetchにも
// 触らない。だからNode/vitestからそのままテストできる（sync-guard.jsと同じ方針）。
// 画面側（index.html/admin.html）が実際の送信を担当する。
(function (root) {
  'use strict';

  var DEFAULT_KEY = 'yotei-pending-add-v1';
  var DEFAULT_MAX = 50;
  var DEFAULT_LEASE_MS = 30000;
  var DEFAULT_GIVE_UP = 10;
  var DEFAULT_BACKOFF = [5000, 15000, 45000, 120000, 300000];

  function createSendQueue(opts) {
    var o = opts || {};
    var storage = o.storage || null;
    var storageKey = o.storageKey || DEFAULT_KEY;
    var tabId = String(o.tabId || 'tab');
    var maxItems = typeof o.maxItems === 'number' ? o.maxItems : DEFAULT_MAX;
    var leaseMs = typeof o.leaseMs === 'number' ? o.leaseMs : DEFAULT_LEASE_MS;
    var giveUpAfter = typeof o.giveUpAfter === 'number' ? o.giveUpAfter : DEFAULT_GIVE_UP;
    var backoffMs = Array.isArray(o.backoffMs) && o.backoffMs.length ? o.backoffMs.slice() : DEFAULT_BACKOFF.slice();

    // storageが使えるか一度だけ実際に書いて確かめる。
    // ★使えない場合、呼び出し側は楽観化を一切行わず従来どおり同期送信で待たせる
    //   （設計書D3）。未送信を無言で失うことを絶対にしないため。
    var usable = false;
    if (storage) {
      try {
        storage.setItem(storageKey + '-probe', '1');
        storage.removeItem(storageKey + '-probe');
        usable = true;
      } catch (e) { usable = false; }
    }

    // storageが使えなくなったときだけ使う控え。usableな間はこれを読まない。
    var memory = null;

    // ★着手前スキャンで発見した欠陥（2026-08-25）:
    // 「一度読んだらメモリに抱えて二度とstorageを読み直さない」実装にすると、
    // 別タブが後から入れた未送信を、先に開いていたタブが見失う。さらに
    // writeStateは状態を丸ごと書き戻すため、先に開いていたタブが次に何か
    // 書いた時点で **別タブの未送信が消える**（登録が黙って失われる＝
    // このアプリで一番防ぎたい事故）。
    // → storageが使える間は毎回storageから読み直す。localStorageの読み取りは
    //   小さなJSONのparseなので、この頻度では速度上の問題にならない。
    function readState() {
      if (storage && usable) {
        var st = { v: 1, items: [] };
        try {
          var raw = storage.getItem(storageKey);
          if (raw) {
            var parsed = JSON.parse(raw);
            if (parsed && Array.isArray(parsed.items)) st = { v: 1, items: parsed.items };
          }
        } catch (e) { /* 壊れていたら空として扱う */ }
        memory = st;
        return st;
      }
      if (!memory) memory = { v: 1, items: [] };
      return memory;
    }

    function writeState(st) {
      memory = st;
      if (!storage) return false;
      try {
        storage.setItem(storageKey, JSON.stringify(st));
        return true;
      } catch (e) {
        // ★sync-guard.jsの6回目レビュー修正2と同じ手当て:
        // 古い内容が残ったまま読み勝つことを防ぐため、キーを消してメモリへ倒す。
        try { storage.removeItem(storageKey); } catch (e2) { /* noop */ }
        usable = false;
        return false;
      }
    }

    function copyItems(items) {
      return items.map(function (it) {
        return {
          id: it.id, rows: (it.rows || []).slice(), company: it.company || '',
          createdAt: it.createdAt || 0, attempts: it.attempts || 0,
          nextAt: it.nextAt || 0, lastError: it.lastError || '',
          gaveUp: !!it.gaveUp, owner: it.owner || '', claimedAt: it.claimedAt || 0
        };
      });
    }

    return {
      isStorageUsable: function () { return usable; },

      enqueue: function (item, now) {
        if (!item) return false;
        var id = String(item.id || '');
        var rows = Array.isArray(item.rows) ? item.rows : [];
        if (!id || rows.length === 0) return false;
        var st = readState();
        if (st.items.length >= maxItems) return false;
        st.items.push({
          id: id, rows: rows.slice(), company: String(item.company || ''),
          createdAt: typeof now === 'number' ? now : 0,
          attempts: 0, nextAt: 0, lastError: '', gaveUp: false,
          owner: tabId, claimedAt: 0
        });
        writeState(st);
        return true;
      },

      list: function () { return copyItems(readState().items); },
      count: function () { return readState().items.length; },

      pendingRows: function (company) {
        var want = typeof company === 'string' && company !== '' ? company : null;
        var out = [];
        readState().items.forEach(function (it) {
          if (want !== null && it.company !== want) return;
          (it.rows || []).forEach(function (r) { out.push(r); });
        });
        return out;
      }
    };
  }

  var api = { createSendQueue: createSendQueue };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = api;
  } else {
    for (var k in api) {
      if (Object.prototype.hasOwnProperty.call(api, k)) root[k] = api[k];
    }
  }
})(typeof window !== 'undefined' ? window : (typeof globalThis !== 'undefined' ? globalThis : this));
```

- [ ] **Step 4: テストが通ることを確認**

Run: `cd cf && npx vitest run`
Expected: PASS（既存153件＋新規14件＝167件）

- [ ] **Step 5: コミット**

```bash
git add send-queue.js cf/test/send-queue.test.js
git commit -m "feat(send-queue): 未送信の登録を貯める箱（永続化・投入・一覧）"
```

---

### Task 2: 送信権・バックオフ・諦め

**Files:**
- Modify: `send-queue.js`
- Test: `cf/test/send-queue.test.js`

**Interfaces:**
- Consumes: Task 1 の `createSendQueue` / `readState` / `writeState` / `copyItems` / `tabId` / `leaseMs` / `giveUpAfter` / `backoffMs`
- Produces:
  - `queue.nextDue(now)` → `item`（複製）または `null`
  - `queue.beginSend(id, now)` → `{ token: string, wasRetry: boolean }` または `null`
  - `queue.markSent(id, token)` → `boolean`
  - `queue.markFailed(id, token, message, now)` → `boolean`
  - `queue.retryNow(id, now)` → `boolean`
  - `queue.gaveUpCount()` → `number`

**この Task の最重要事項（設計書 D9）:** `beginSend` は fetch を投げる**前**に `attempts` を +1 して永続化する。永続化に失敗したら `null` を返して**送らせない**。送信中にタブが落ちても、次に拾われたとき必ず再送扱い（`wasRetry:true`）になり、存在確認が働く。

- [ ] **Step 1: 失敗するテストを書く**

`cf/test/send-queue.test.js` の末尾に追記:

```javascript
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
});
```

- [ ] **Step 2: テストを走らせて失敗することを確認**

Run: `cd cf && npx vitest run test/send-queue.test.js`
Expected: FAIL（`q.nextDue is not a function`）

- [ ] **Step 3: `send-queue.js` にヘルパ3つを足す**

`copyItems` の直後（`return {` の直前）に追加:

```javascript
    function findItem(st, id) {
      for (var i = 0; i < st.items.length; i++) {
        if (st.items[i].id === String(id)) return st.items[i];
      }
      return null;
    }

    // tokenは「どのタブの、何回目の試行か」。これが一致しない限り
    // markSent/markFailed は反映しない（古い試行が新しい試行の状態を
    // 上書きするのを防ぐ。sync-guard.jsの5回目レビュー修正7と同じ考え方）。
    function tokenOf(item) { return item.owner + ':' + item.attempts; }

    // 今この項目を拾ってよいか。
    function isDue(it, now) {
      if (it.gaveUp) return false;
      if ((it.nextAt || 0) > now) return false;
      // 誰かが送信中（リース有効）なら、その持ち主以外は拾わない
      if (it.claimedAt && (it.claimedAt + leaseMs) > now && it.owner !== tabId) return false;
      // ★初回送信は enqueue したタブだけが行う（設計書D2(a)）。
      // 他タブはリース経過後にしか拾わない＝そのときは必ず wasRetry になり、
      // 存在確認（画面側）が働くため二重登録にならない。
      if ((it.attempts || 0) === 0 && it.owner && it.owner !== tabId
          && (now - (it.createdAt || 0)) < leaseMs) return false;
      return true;
    }
```

- [ ] **Step 4: `send-queue.js` の `return { ... }` に6つのメソッドを足す**

`pendingRows` の直後に追加（末尾のカンマに注意）:

```javascript
      nextDue: function (now) {
        var n = typeof now === 'number' ? now : 0;
        var st = readState();
        for (var i = 0; i < st.items.length; i++) {
          if (isDue(st.items[i], n)) return copyItems([st.items[i]])[0];
        }
        return null;
      },

      // ★fetchの前に attempts を +1 して永続化する（設計書D9）。
      // 永続化できなければ null を返して送らせない（記録が残らない以上、
      // 次回に再送扱いへ倒せず、二重登録の危険が残るため）。
      beginSend: function (id, now) {
        var n = typeof now === 'number' ? now : 0;
        var st = readState();
        var it = findItem(st, id);
        if (!it || !isDue(it, n)) return null;
        var wasRetry = (it.attempts || 0) >= 1;
        it.attempts = (it.attempts || 0) + 1;
        it.owner = tabId;
        it.claimedAt = n;
        if (!writeState(st)) return null;
        return { token: tokenOf(it), wasRetry: wasRetry };
      },

      markSent: function (id, token) {
        var st = readState();
        var it = findItem(st, id);
        if (!it || tokenOf(it) !== String(token)) return false;
        st.items = st.items.filter(function (x) { return x !== it; });
        writeState(st);
        return true;
      },

      markFailed: function (id, token, message, now) {
        var n = typeof now === 'number' ? now : 0;
        var st = readState();
        var it = findItem(st, id);
        if (!it || tokenOf(it) !== String(token)) return false;
        var idx = Math.min(Math.max((it.attempts || 1) - 1, 0), backoffMs.length - 1);
        it.nextAt = n + backoffMs[idx];
        it.lastError = String(message || '');
        it.claimedAt = 0;
        if ((it.attempts || 0) >= giveUpAfter) it.gaveUp = true;
        writeState(st);
        return true;
      },

      retryNow: function (id, now) {
        var n = typeof now === 'number' ? now : 0;
        var st = readState();
        var it = findItem(st, id);
        if (!it) return false;
        it.gaveUp = false;
        it.nextAt = 0;
        it.claimedAt = 0;
        it.owner = tabId;
        writeState(st);
        return true;
      },

      gaveUpCount: function () {
        return readState().items.filter(function (it) { return !!it.gaveUp; }).length;
      },
```

- [ ] **Step 5: テストが通ることを確認**

Run: `cd cf && npx vitest run`
Expected: PASS（既存153件＋Task1の14件＋Task2の12件＝179件）

- [ ] **Step 6: コミット**

```bash
git add send-queue.js cf/test/send-queue.test.js
git commit -m "feat(send-queue): 送信権・バックオフ・諦め。attemptsは送信の前に永続化（D9）"
```

---

### Task 3: `index.html` に送信係を配線（`submitNippo` はまだ従来どおり）

**Files:**
- Modify: `index.html`
- Test: 手動（`node --check` ＋ ブラウザのコンソール）

**Interfaces:**
- Consumes: Task 2 までの `createSendQueue` 全メソッド
- Produces:
  - `sendQueue`（グローバル。`createSendQueue` の戻り値）
  - `async function drainQueue()` → `Promise<void>`（例外を投げない。多重起動しない）
  - `function scheduleDrain(ms)` → `void`
  - `async function checkAlreadyLanded(item)` → `Promise<boolean>`（失敗時は例外を投げる）
  - `function renderPendingUi()` → `void`（このTaskでは空実装。中身はTask 4）

**このTaskの狙い:** 「箱に入っている物を安全に送る」機能だけを先に作り、レビューを通す。`submitNippo` はまだ触らないので、**この時点でも利用者の見た目は1ミリも変わらない。**

- [ ] **Step 1: `send-queue.js` を読み込む＋読めなかったときの最小フォールバック**

`index.html` の `<script src="sync-guard.js"></script>`（975行付近）の直後に追加。フォールバックは `sync-guard.js` の前例（977〜992行）と同じ考え方＝**読めなかったら楽観化を一切せず従来どおりに倒す。**

```html
<script src="send-queue.js"></script>
<script>
// ★send-queue.jsが読めなかったとき（配置漏れ・部分配布・キャッシュ不整合）に
// 画面が壊れないための最小フォールバック。sync-guard.jsと同じ方針で、
// 「使えない」と答えるだけの箱を置く＝submitNippoは従来どおり同期送信で待たせる。
if(typeof createSendQueue==='undefined'){
  try{console.error('[予定管理] send-queue.jsの読み込みに失敗しました。保存は従来どおり（送信を待つ）で動作します');}catch(e){}
  window.createSendQueue=function(){return{
    isStorageUsable:function(){return false;},
    enqueue:function(){return false;},
    list:function(){return [];},
    count:function(){return 0;},
    pendingRows:function(){return [];},
    nextDue:function(){return null;},
    beginSend:function(){return null;},
    markSent:function(){return false;},
    markFailed:function(){return false;},
    retryNow:function(){return false;},
    gaveUpCount:function(){return 0;}
  };};
}
</script>
```

- [ ] **Step 2: `sendQueue` を作る**

`const SYNC_TIMEOUT_MS=20000;`（1025行）の直後に追加:

```javascript
// ★案A: 未送信の登録を貯める箱。タブごとに違うIDを持たせ、
// 「初回送信は登録したタブだけが行う」判定に使う（設計書D2）。
const SEND_TAB_ID='t'+Math.random().toString(36).slice(2)+Date.now().toString(36);
const sendQueue=createSendQueue({
  storage:(function(){try{return window.localStorage;}catch(e){return null;}})(),
  tabId:SEND_TAB_ID
});
let drainRunning=false;
let drainTimer=null;
```

- [ ] **Step 3: `renderPendingUi` の空実装を置く**

`function setStaleBadge(on,failed){`（2158行）の直前に追加:

```javascript
// ★案A: 未送信の帯と印。中身はTask4で実装する（ここでは何もしない）。
function renderPendingUi(){}
```

- [ ] **Step 4: 送信係 `drainQueue` を書く**

`function refreshInBackground(after){`（2250行付近）の直前に追加:

```javascript
// ★案A 送信係: 箱から1件取り出して送る。例外を投げない。多重起動しない。
//
// 2回目以降の送信では、送る直前にGASから最新を読み「もう入っていないか」を
// 確認する（設計書§4.3）。2026-08-25の実測で「送信側にエラーが返ったのに
// 実際は書き込まれていた」が実際に発生しており、確認せず再送すると
// 予定が二重になる。確認はD1ではなくGASから行う（D1は最大5分古くなりうる）。
async function checkAlreadyLanded(item){
  const base=getGasUrl();
  const res=await fetch(base+(base.includes('?')?'&':'?')+'compact=1&t='+Date.now(),
    {signal:timeoutSignal(GAS_READ_TIMEOUT_MS)});
  if(!res.ok)throw new Error('確認のための読み取りに失敗しました('+res.status+')');
  const json=await res.json();
  if(!json||json.status!=='ok')throw new Error('確認のための読み取りに失敗しました');
  const rows=json.compact?expandCompactRows(json.headers,json.rows):json.rows;
  const seen=new Set((rows||[]).map(r=>String((r&&(r.ID!==undefined?r.ID:r.id))||'')));
  return (item.rows||[]).some(r=>seen.has(String(r.id||'')));
}

async function drainQueue(){
  if(drainRunning)return;
  drainRunning=true;
  try{
    for(;;){
      const item=sendQueue.nextDue(Date.now());
      if(!item)break;
      const claim=sendQueue.beginSend(item.id,Date.now());
      if(!claim)break; // 他タブが持っている／記録できなかった＝安全側で何もしない
      try{
        if(claim.wasRetry){
          const landed=await checkAlreadyLanded(item);
          if(landed){sendQueue.markSent(item.id,claim.token);renderPendingUi();continue;}
        }
        const res=await fetch(getGasUrl(),{method:'POST',
          body:JSON.stringify({action:'add',rows:item.rows,updatedBy:(item.rows[0]&&item.rows[0].updatedBy)||''}),
          headers:{'Content-Type':'text/plain'},signal:timeoutSignal(GAS_READ_TIMEOUT_MS)});
        const json=await res.json();
        if(res.ok&&json&&json.status==='ok'){
          sendQueue.markSent(item.id,claim.token);
          renderPendingUi();
          refreshInBackground();
        }else{
          sendQueue.markFailed(item.id,claim.token,(json&&json.message)||('HTTP '+res.status),Date.now());
          renderPendingUi();
          break;
        }
      }catch(err){
        sendQueue.markFailed(item.id,claim.token,String((err&&err.message)||err),Date.now());
        renderPendingUi();
        break;
      }
    }
  }finally{
    drainRunning=false;
    try{if(sendQueue.count()>0)scheduleDrain(5000);}catch(e){}
  }
}

function scheduleDrain(ms){
  if(drainTimer)clearTimeout(drainTimer);
  drainTimer=setTimeout(()=>{drainTimer=null;drainQueue();},Math.max(1000,ms||5000));
}
```

- [ ] **Step 5: 起動時とオンライン復帰で送信係を起こす**

`else{if(hydrateFromSnapshot(currentCompany))setStaleBadge(true);loadData();}`（4094行付近）の直後に追加:

```javascript
// ★案A: 前回送れなかった登録を、アプリを開いたときと電波が戻ったときに送る
try{
  if(sendQueue.count()>0){renderPendingUi();scheduleDrain(1000);}
  window.addEventListener('online',()=>scheduleDrain(1000));
}catch(e){}
```

- [ ] **Step 6: 構文チェックとテスト**

```bash
node --check send-queue.js
cd cf && npx vitest run
```
Expected: エラー無し／179件 PASS

- [ ] **Step 7: 見た目が変わっていないことを確認**

`index.html` を Chrome で `file:///` から開き、コンソールに `send-queue.js` のエラーが出ないこと、`sendQueue.count()` が `0` を返すことを確認する。**この時点で保存の挙動は従来どおり（押すと待つ）。**

- [ ] **Step 8: コミット**

```bash
git add index.html
git commit -m "feat(index.html): 送信係drainQueueを配線（submitNippoはまだ従来どおり＝見た目は無変化）"
```

---

### Task 4: `index.html` の `submitNippo` を楽観化＋未送信の表示

**Files:**
- Modify: `index.html`（`submitNippo` 1966-2100行付近 / `addLocalRows` 2237行 / `hydrateFromSnapshot` 2152行 / `loadData` 2408・2434行 / `renderPendingUi` / カード描画 2468・2615・2635行）
- Test: 手動（`node --check` ＋ ブラウザ）

**Interfaces:**
- Consumes: Task 3 の `sendQueue` / `drainQueue` / `scheduleDrain` / `renderPendingUi`
- Produces:
  - `addLocalRows(rows, opts)` — `opts.pending===true` のとき各行に `isPending:true` を付ける（`opts` 省略時は従来どおり＝後方互換）
  - `function applyPendingRows()` → `void`

- [ ] **Step 1: `addLocalRows` に `pending` を足す**

2237行の `function addLocalRows(rows){` を次に置き換える:

```javascript
function addLocalRows(rows,opts){
  const pending=!!(opts&&opts.pending);
  const added=(rows||[]).map(r=>({
    id:String(r.id||''),timestamp:'',
    date:r.date,genba:r.genba||'',loc:r.loc||'',name:r.name||'',role:r.role||'',
    start:formatTime(r.start),end:formatTime(r.end),
    kosu:Number(r.kosu)||0,memo:r.memo||'',
    yakin:!!r.yakin,yasumi:!!r.yasumi,yotei:!!r.yotei,souko:!!r.souko,
    company:r.company||currentCompany,updatedBy:r.updatedBy||'',color:r.color||'',
    division:String(r.jobNoDivision||''),jobNo:'',workType:r.workType||'',vehicle:r.vehicle||'',
    isPending:pending
  }));
  allNippos=allNippos.concat(added).concat(generateGhosts(added));
  rerenderAll();
}
```

- [ ] **Step 2: 未送信を載せ直す `applyPendingRows` を書く**

`addLocalRows` の直後に追加:

```javascript
// ★案A（設計書D1）: loadData / hydrateFromSnapshot は allNippos を丸ごと
// 差し替えるため、未送信の行が画面から消える。差し替えのたびに載せ直す。
// 未送信は「本人の予定」として人工の集計にも数える（利用者判断 2026-08-25）
// ＝isGhostのように除外はしない。
function applyPendingRows(){
  let rows;
  // ★設計書D6: loadDataは currentCompany で絞って取得しているため、
  // 他社の未送信をそのまま足すと画面に他社の予定が混ざる。表示中の会社だけ足す。
  // 「全社」表示のときは全部足す（登録ボタンは全社では無効なので、
  // 箱の中身に company==='全社' の項目は存在しない）。
  try{rows=sendQueue.pendingRows(currentCompany==='全社'?'':currentCompany);}catch(e){return;}
  if(!rows||!rows.length)return;
  const have=new Set(allNippos.map(n=>String(n.id)));
  const fresh=rows.filter(r=>!have.has(String(r.id||'')));
  if(!fresh.length)return;
  addLocalRows(fresh,{pending:true});
}
```

- [ ] **Step 3: `allNippos` を差し替えている3か所の直後で載せ直す**

- 2152行 `allNippos=base.concat(generateGhosts(base));` の直後に `applyPendingRows();`
- 2410行 `allNippos=allNippos.concat(generateGhosts(allNippos));` の直後に `applyPendingRows();`
- 2434行 `allNippos=[];allMembers=[];allGenbaMaster=[];allJobsites=[];` の直後に `applyPendingRows();`

- [ ] **Step 4: `submitNippo` の送信部分を差し替える**

2074行の `const btn=document.getElementById('submit-btn');` から2100行の `isSubmitting=false;btn.textContent='登録する';btn.disabled=(currentCompany==='全社');` までを次に置き換える。

**フォールバック条件（設計書D3・D4）:** 箱が使えない／上限超過のときは**従来どおり同期送信で待たせる**。無言で未送信を失うことを絶対にしない。

```javascript
  const btn=document.getElementById('submit-btn');
  const clearForm=()=>{
    document.getElementById('s-genba').value='';document.getElementById('s-genba-select').value='';document.getElementById('s-genba').style.display='none';document.getElementById('s-location').value='';
    document.getElementById('s-memo').value='';document.getElementById('s-leader').value='';
    document.getElementById('s-date-end').value='';
    const _v=document.getElementById('s-vehicle');if(_v)_v.value='';
    refreshVehicleSelect('s');
    const _wt=document.getElementById('s-work-type');if(_wt)_wt.value='';
    applyWorkTypeLock('s');
    const _ls=document.getElementById('s-location-select');if(_ls)_ls.value='';
    const _li=document.getElementById('s-location');if(_li){_li.value='';_li.style.display='none';}
    const _lse=document.getElementById('s-location-search');if(_lse)_lse.value='';
    const _lsim=document.getElementById('s-loc-similar');if(_lsim){_lsim.style.display='none';_lsim.innerHTML='';}
    document.getElementById('s-jobno-hint').style.display='none';
    document.getElementById('s-jobno-div-wrap').style.display='none';
    const _jnSel=document.getElementById('s-jobno-division');if(_jnSel){delete _jnSel.dataset.userSet;_jnSel.value='ICT';}
    applyMode('s','');selectedMembers=[];renderMemberChips('member-selector',selectedMembers,toggleMember);
  };

  // ★案A: 箱が使えるときだけ楽観化する。使えないとき（localStorage不可・
  // 上限超過・send-queue.js未読込）は従来どおり同期送信で待たせる＝安全側。
  const canQueue=sendQueue.isStorageUsable()
    &&sendQueue.enqueue({id:rows[0].id,rows:rows,company:currentCompany},Date.now());

  if(canQueue){
    hapticSuccess();
    showAlert(`✓ ${dates.length}日×${members.length}人分を登録しました！（${currentCompany}）`,'ok');
    clearForm();
    addLocalRows(rows,{pending:true});
    renderPendingUi();
    drainQueue();               // awaitしない＝ここで待たせない
    isSubmitting=false;btn.textContent='登録する';btn.disabled=(currentCompany==='全社');
    return;
  }

  isSubmitting=true;btn.textContent=`送信中（${dates.length}日分）...`;btn.disabled=true;
  try{
    const res=await fetch(getGasUrl(),{method:'POST',body:JSON.stringify({action:'add',rows}),headers:{'Content-Type':'text/plain'}});
    const json=await res.json();
    if(json.status==='ok'){
      hapticSuccess();
      showAlert(`✓ ${dates.length}日×${members.length}人分を登録しました！（${currentCompany}）`,'ok');
      clearForm();
      addLocalRows(rows);refreshInBackground();
    }else showAlert('エラー：'+json.message,'err');
  }catch(err){showAlert('送信エラー：'+err.message,'err');}
  isSubmitting=false;btn.textContent='登録する';btn.disabled=(currentCompany==='全社');
```

- [ ] **Step 5: 未送信の帯を実装する**

Task 3 で置いた `function renderPendingUi(){}` を次に置き換える:

```javascript
// ★案A: 未送信の帯。件数は全社分を出す（会社を切り替えていても
// 見落とさないため・設計書D6）。タップで諦めた分も含めて手動再送する。
function renderPendingUi(){
  let n=0,gave=0;
  try{n=sendQueue.count();gave=sendQueue.gaveUpCount();}catch(e){return;}
  let el=document.getElementById('cal-pending-badge');
  if(n===0){if(el)el.remove();rerenderAll();return;}
  if(!el){
    el=document.createElement('div');el.id='cal-pending-badge';
    el.style.cssText='position:sticky;top:0;z-index:60;padding:8px 12px;font-size:13px;text-align:center;cursor:pointer';
    document.body.insertBefore(el,document.body.firstChild);
    el.onclick=()=>{try{sendQueue.list().forEach(it=>sendQueue.retryNow(it.id,Date.now()));drainQueue();}catch(e){}};
  }
  if(gave>0){
    el.style.background='#ffdad6';el.style.color='#8c0009';el.style.fontWeight='700';
    el.textContent=`⚠ ${gave}件が送信できていません。タップして再送してください`;
  }else{
    el.style.background='#fff4d6';el.style.color='#7a5b00';el.style.fontWeight='400';
    el.textContent=`⏳ 未送信 ${n}件（自動で送信中）`;
  }
  rerenderAll();
}
```

- [ ] **Step 6: 予定カードに未送信の印を出す**

- 2468行の `isGhost:!!n.isGhost,originalId:n.originalId||''` を
  `isGhost:!!n.isGhost,isPending:!!n.isPending,originalId:n.originalId||''` にする
- 2615行 `const ghostMarker=g.isGhost?...;` の直後に追加:
  ```javascript
  const pendingMarker=g.isPending?`<span class="ghost-marker">⏳未送信</span>`:'';
  ```
- 2635行付近のカードHTMLで `${ghostMarker}` を出している箇所を `${ghostMarker}${pendingMarker}` にする

- [ ] **Step 7: テスト**

Run: `cd cf && npx vitest run`
Expected: 179件 PASS（既存を1件も壊していないこと）

- [ ] **Step 8: コミット**

```bash
git add index.html
git commit -m "feat(index.html): submitNippoを楽観化。未送信の保持・帯・印（案A本体）"
```

---

### Task 5: `admin.html` へ同一の変更を反映

**Files:**
- Modify: `admin.html`

**Interfaces:**
- Consumes: Task 4 までの全て
- Produces: なし（`index.html` と同一の挙動）

**このTaskの唯一の要件:** Task 3・4 で `index.html` に入れた変更を、**1文字も違わずに** `admin.html` へ入れる（行番号と、ログ出力のプレフィックス `[予定管理-admin]` だけが異なる）。過去のレビューで両画面の乖離が繰り返し重大指摘になっている。

`admin.html` の対応箇所:
- `<script src="sync-guard.js">` = 1150行付近
- 定数群（`SYNC_TIMEOUT_MS`）= 1206行付近
- `submitNippo` = 2428行
- `refreshInBackground` = 2113行付近
- `setStaleBadge` / `addLocalRows` / `loadData` / `hydrateFromSnapshot` は `index.html` と同じ関数名で存在する

- [ ] **Step 1: `index.html` に入れた差分を取り出す**

```bash
git diff HEAD~2 -- index.html
```

- [ ] **Step 2: 同じ内容を `admin.html` に適用する**

**Task 3 の Step 1〜5、および Task 4 の Step 1〜6 に載っているコードブロックを、
そのまま（1文字も変えずに）`admin.html` の対応箇所へ入れる。**
`admin.html` にも `getGasUrl` / `timeoutSignal` / `expandCompactRows` /
`refreshInBackground` / `addLocalRows` / `rerenderAll` / `generateGhosts` /
`hapticSuccess` / `showAlert` / `currentCompany` / `allNippos` は
同じ名前で存在するため、コードの中身を書き換える必要は無い。

- [ ] **Step 3: 両画面の該当コードが同数入っていることを確認**

```bash
for f in index.html admin.html; do
  printf "%s: " "$f"
  grep -c "sendQueue\|drainQueue\|applyPendingRows\|renderPendingUi\|isPending\|checkAlreadyLanded" "$f"
done
```
Expected: 2つの数が一致

- [ ] **Step 4: テスト**

Run: `cd cf && npx vitest run`
Expected: 179件 PASS

- [ ] **Step 5: コミット**

```bash
git add admin.html
git commit -m "feat(admin.html): index.htmlと同一の楽観的保存を反映"
```

---

### Task 6: 通し確認と記録

**Files:**
- Modify: `引き継ぎ.md`

- [ ] **Step 1: 自動テストの全件確認**

Run: `cd cf && npx vitest run`
Expected: 179件 PASS。**4回連続で実行し、不安定なテストが無いことを確認する。**

- [ ] **Step 2: 手で確かめる項目を洗い出して記録する**

以下は実ブラウザでしか確認できない。**確認していない項目を「確認済み」と書かない。**

| # | 確認すること | 期待 |
|---|---|---|
| 1 | 予定を1件登録する | **押した瞬間**に予定が出てフォームが空になる |
| 2 | 登録直後の予定 | ⏳未送信の印が付き、上に帯が出る |
| 3 | 数秒待つ | 印と帯が黙って消える |
| 4 | 登録直後にリロード | **自分の登録が見える**（read-your-own-writes） |
| 5 | 機内モードにして登録 | 予定は出る。帯が「未送信」のまま残る |
| 6 | 機内モードのままリロード | **未送信の予定が消えていない** |
| 7 | 機内モードを解除 | **自動で送信され、印と帯が消える** |
| 8 | 出面表（10日締め）を開く | 未送信の分も人工に数えられている |
| 9 | スマホのPWAで1・4を試す | 同じ結果 |
| 10 | 管理画面でも1〜4を試す | 同じ結果 |

- [ ] **Step 3: `引き継ぎ.md` を更新する**

「最終更新」を書き換え、案Aの節を足す（実装内容・戻し方・**未確認のまま残っている項目**）。

- [ ] **Step 4: コミット**

```bash
git add 引き継ぎ.md
git commit -m "docs: 案A（保存の楽観的表示）の記録と通し確認の結果"
```
