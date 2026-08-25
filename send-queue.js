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

  // ★レビュー修正ラウンド1・Critical: readStateが読み込んだitemsの各要素を検疫する。
  // 捨てる条件はこの2つだけに限定する（厳しくしすぎて正常な項目を捨てないため）:
  //   - 項目がオブジェクトでない（null・数値・文字列等）
  //   - idが空文字、またはrowsが配列でない
  // これを怠ると、一度壊れた項目がstorageに乗った時点でlist()/pendingRows()が
  // 毎回例外を投げ続け、未送信が永久に送られなくなる（帯の表示・カレンダーへの
  // 合流・送信ループはすべてこの2つを呼ぶ）。
  function sanitizeItems(rawItems) {
    var out = [];
    for (var i = 0; i < rawItems.length; i++) {
      var it = rawItems[i];
      if (!it || typeof it !== 'object') continue;
      var id = typeof it.id === 'string' ? it.id : String(it.id || '');
      if (!id) continue;
      if (!Array.isArray(it.rows)) continue;
      out.push(it);
    }
    return out;
  }

  // ★レビュー修正ラウンド1・Important 2: rowsの中身を深く複製する。
  // rowsに積まれるのは文字列・数値・真偽値だけ（index.html:1961で日付は
  // YYYY-MM-DD文字列、1976-1977で時刻はinputのvalue＝文字列、人工は数値、
  // フラグは真偽値であることを確認済み）なのでJSON往復で複製できる。
  // 往復が失敗した場合だけslice()にフォールバックし、落ちないようにする。
  function cloneRows(rows) {
    try {
      return JSON.parse(JSON.stringify(rows));
    } catch (e) {
      return rows.slice();
    }
  }

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
            if (parsed && Array.isArray(parsed.items)) {
              st = { v: 1, items: sanitizeItems(parsed.items) };
            }
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

    // ★レビュー修正ラウンド1・Important 1: 状態を変更する処理は必ずこれを通す。
    // 直前に読み直してから変更して書くため、「読んでから書くまでの間に別タブが
    // 入れた内容」を巻き込まずに済む。fnの中でawaitを挟まないこと（挟むと窓が
    // 再び開く）。Task 2で追加するbeginSend/markSent/markFailed/retryNowも同じ
    // 入口を使う前提なので、_internalsとして外から使える形にしておく。
    // 限界: localStorageには比較交換（CAS）が無いため、2つのタブが同じ同期
    // ブロックで書いた場合の競合までは防げない（1人1端末が基本の業務アプリの
    // ため許容）。
    function mutate(fn) {
      var st = readState();          // 直前に読み直す
      var ret = fn(st);
      var ok = writeState(st);
      return { ok: ok, ret: ret };
    }

    function copyItems(items) {
      return items.map(function (it) {
        return {
          id: it.id, rows: cloneRows(it.rows || []), company: it.company || '',
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
        var accepted = false;
        mutate(function (st) {
          if (st.items.length >= maxItems) { accepted = false; return; }
          st.items.push({
            id: id, rows: cloneRows(rows), company: String(item.company || ''),
            createdAt: typeof now === 'number' ? now : 0,
            attempts: 0, nextAt: 0, lastError: '', gaveUp: false,
            owner: tabId, claimedAt: 0
          });
          accepted = true;
        });
        return accepted;
      },

      list: function () { return copyItems(readState().items); },
      count: function () { return readState().items.length; },

      pendingRows: function (company) {
        var want = typeof company === 'string' && company !== '' ? company : null;
        var out = [];
        readState().items.forEach(function (it) {
          if (want !== null && it.company !== want) return;
          cloneRows(it.rows || []).forEach(function (r) { out.push(r); });
        });
        return out;
      },

      // ★Task 2以降（beginSend/markSent/markFailed/retryNow等）が同じ入口を
      // 使えるように公開する。テストからも直接検証できる。
      _internals: {
        mutate: mutate,
        readState: readState,
        writeState: writeState
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
