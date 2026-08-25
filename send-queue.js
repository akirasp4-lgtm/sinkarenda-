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
