// send-queue.js — 未送信の登録を貯める箱（案A: 楽観的表示＋裏送信）
//
// 設計書: docs/superpowers/specs/2026-08-25-optimistic-save-design.md
//
// この箱は「次にどれを送るか」だけを判断する純ロジックで、DOMにもfetchにも
// 触らない。だからNode/vitestからそのままテストできる（sync-guard.jsと同じ方針）。
// 画面側（index.html/admin.html）が実際の送信を担当する。
//
// ★修正ラウンド2・変更1: 保存の形を「1つのキーに全項目をまとめる」から
// 「1件＝1キー（storageKey + ':' + id）」に変えた。別々の項目を触る2つの
// タブがそもそも同じキーを書かなくなるため、「読み直してから丸ごと書き戻す」
// ことに伴う交錯（修正ラウンド1でCritical C-1として直しても窓を塞ぎ切れな
// かった問題）が構造的に起きなくなる。同じ項目を2つのタブが同時に触るケース
// （既存のリース・初回所有権が通常は防ぐ）だけ、beginSend内の読み直し確認で
// 引き続き二重に守る。
(function (root) {
  'use strict';

  var DEFAULT_KEY = 'yotei-pending-add-v1';
  var DEFAULT_MAX = 50;
  var DEFAULT_LEASE_MS = 30000;
  var DEFAULT_GIVE_UP = 10;
  var DEFAULT_BACKOFF = [5000, 15000, 45000, 120000, 300000];

  // 項目1件の形が正しいかを検査する。壊れていたら null を返すだけで、
  // 呼び出し側はそのキーを読み飛ばす。
  // ★修正ラウンド2・変更1: 消さない。職人が入力した予定そのものなので、
  // 壊れて読めないからといって勝手に消していい理由がない（変更2と同じ考え方）。
  function sanitizeItem(it) {
    if (!it || typeof it !== 'object') return null;
    var id = typeof it.id === 'string' ? it.id : String(it.id || '');
    if (!id) return null;
    if (!Array.isArray(it.rows)) return null;
    return it;
  }

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
    var prefix = storageKey + ':';
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

    // storageが使えない／使えなくなったときに使う控え（id → 項目）。
    // usableな間もopportunisticに更新するが、そこから読みに行くのはusableが
    // falseの間だけ（Task1の「memory」を1件単位のMapにした形）。
    var memory = new Map();

    // 1件読む。storageが使える間は毎回storageから読み直す（Task1由来の
    // 「毎回読み直す」方針を1件単位に引き継ぐ）。読めたらmemoryにも反映し、
    // 無ければmemoryからも消す（他タブが消した項目を抱え込み続けないため）。
    function readItem(id) {
      if (storage && usable) {
        try {
          var raw = storage.getItem(prefix + id);
          if (!raw) { memory.delete(id); return null; }
          var it = sanitizeItem(JSON.parse(raw));
          if (it) memory.set(id, it);
          return it;
        } catch (e) { return null; }
      }
      return memory.has(id) ? memory.get(id) : null;
    }

    // 1件書く。
    // ★修正ラウンド2・変更2: 失敗してもキーを消さない（sync-guard.jsの踏襲を
    // やめる）。この箱に入っているのは職人が入力した予定そのものであり、キーを
    // 消す＝未送信の全消去になる。usable=falseにしてメモリ運転へ移るだけにする。
    // storageに古い内容が残っても、以後readItem/listAllItemsはmemoryを見る
    // ので実害はない。むしろ次にページを開いてstorageが回復したとき、残って
    // いた古い内容がそのまま読まれて送信される＝救済になる。
    function writeItem(item) {
      memory.set(item.id, item);
      if (!storage || !usable) return false;
      try {
        storage.setItem(prefix + item.id, JSON.stringify(item));
        return true;
      } catch (e) {
        usable = false;
        return false;
      }
    }

    function deleteItem(id) {
      memory.delete(id);
      if (!storage || !usable) return;
      try { storage.removeItem(prefix + id); } catch (e) { /* noop */ }
    }

    // 全件を走査する。
    // ★修正ラウンド2・変更3: このオリジンには他アプリの大きなキャッシュ等も
    // 同居しているため、走査は prefix（storageKey + ':'）の前方一致に厳密に
    // 限定する。それ以外のキー（無関係なキー・旧形式の1キーまとめ）は読みも
    // 消しもしない。
    function listAllItems() {
      var out = [];
      if (storage && usable) {
        var len = 0;
        try { len = storage.length; } catch (e) { len = 0; }
        var seen = {};
        for (var i = 0; i < len; i++) {
          var k = null;
          try { k = storage.key(i); } catch (e) { k = null; }
          if (typeof k !== 'string' || k.indexOf(prefix) !== 0) continue;
          var id = k.slice(prefix.length);
          seen[id] = true;
          try {
            var raw = storage.getItem(k);
            if (!raw) continue;
            var it = sanitizeItem(JSON.parse(raw));
            if (it) { out.push(it); memory.set(id, it); }
          } catch (e) { /* 壊れていたら読み飛ばす（storageからは消さない） */ }
        }
        // storage側で見えなくなった項目はmemoryのミラーからも落とす（他タブが
        // 送り終えて消した項目を、usable=falseに落ちた後まで「未送信」として
        // 抱え続けないため）。
        memory.forEach(function (_v, mid) { if (!seen[mid]) memory.delete(mid); });
      } else {
        memory.forEach(function (it) { out.push(it); });
      }
      out.sort(function (a, b) { return (a.createdAt || 0) - (b.createdAt || 0); });
      return out;
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
      // 他タブが拾えるのはリース経過後（(now-createdAt)>=leaseMs）。attempts>=1
      // （＝enqueueしたタブが少なくとも1回は送信を試みている）で拾った場合は
      // wasRetryになり、存在確認（画面側のcheckAlreadyLanded）が働くため
      // 二重登録にならない。
      // ★修正ラウンド1・Minor-3: 「他タブはリース経過後にしか拾わない＝そのときは
      // 必ずwasRetryになる」という以前のコメントは不正確だった。attempts=0のまま
      // （＝enqueueしたタブがbeginSendを一度も呼べておらず、ネットワークには一度も
      // 出ていない）リースだけが経過したケースでは、拾った側もattempts=0からの
      // 開始になりwasRetryにならない。ただしこの場合は元のタブが一度も送信して
      // いないので「初回送信」がここで初めて行われるだけであり、動作としては
      // 安全（二重にはならない）。
      if ((it.attempts || 0) === 0 && it.owner && it.owner !== tabId
          && (now - (it.createdAt || 0)) < leaseMs) return false;
      return true;
    }

    return {
      isStorageUsable: function () { return usable; },

      enqueue: function (item, now) {
        if (!item) return false;
        var id = String(item.id || '');
        var rows = Array.isArray(item.rows) ? item.rows : [];
        if (!id || rows.length === 0) return false;
        // ★修正ラウンド3: 同一idの再投入を拒否する。受け入れると attempts が
        // 0に巻き戻り、送信開始済みの項目が「初回」扱いで再送されて予定が
        // 二重になる。呼び出し側は id を毎回新規採番する契約（submitNippo の
        // uuid()）だが、「到達しない前提」に頼らず構造で閉じる。
        if (readItem(id)) return false;
        if (listAllItems().length >= maxItems) return false;
        // ★修正ラウンド1・Important-2: 生成時のプローブ（usable判定）を通っていても、
        // この項目自身の書き込みでquota満杯等により初めて失敗することがある。
        // 従来はwriteItemの成否を見ずに常にtrueを返していたため、呼び出し側には
        // 「箱に入った」と伝わるのに実際はlocalStorageに書かれておらず、usableが
        // falseに落ちて以後beginSendが常にnullを返す＝タブを閉じた瞬間にこの
        // 未送信が消える事故になっていた（レビュー指摘・再現済み）。
        // usableが「この呼び出しの直前まではtrueだった」のに書き込みが失敗した
        // 場合だけ、この項目を受理しなかったことにする（memoryからも消して
        // falseを返す。残すとdrainQueueが後から拾って実体の無い送信を試みる）。
        // 呼び出し側（submitNippo）はfalseを見て従来どおり同期送信に倒れる。
        // ★一方、usableが最初（コンストラクタのプローブ）からfalseだったタブは
        // 従来どおりメモリ運転で受理してtrueを返す（Task1由来の既存動作を維持。
        // 実運用ではTask4のsubmitNippoがisStorageUsable()を先に見てそもそも
        // enqueueを呼ばない設計だが、直接呼び出す既存テスト・呼び出し側の
        // 後方互換のため残す）。
        var wasUsable = usable;
        var ok = writeItem({
          id: id, rows: cloneRows(rows), company: String(item.company || ''),
          createdAt: typeof now === 'number' ? now : 0,
          attempts: 0, nextAt: 0, lastError: '', gaveUp: false,
          owner: tabId, claimedAt: 0
        });
        if (!ok && wasUsable) { memory.delete(id); return false; }
        return true;
      },

      list: function () { return copyItems(listAllItems()); },
      count: function () { return listAllItems().length; },

      pendingRows: function (company) {
        var want = typeof company === 'string' && company !== '' ? company : null;
        var out = [];
        listAllItems().forEach(function (it) {
          if (want !== null && it.company !== want) return;
          cloneRows(it.rows || []).forEach(function (r) { out.push(r); });
        });
        return out;
      },

      nextDue: function (now) {
        var n = typeof now === 'number' ? now : 0;
        var items = listAllItems();
        for (var i = 0; i < items.length; i++) {
          if (isDue(items[i], n)) return copyItems([items[i]])[0];
        }
        return null;
      },

      // ★fetchの前に attempts を +1 して永続化する（設計書D9）。
      // 永続化できなければ null を返して送らせない（記録が残らない以上、
      // 次回に再送扱いへ倒せず、二重登録の危険が残るため）。
      beginSend: function (id, now) {
        var n = typeof now === 'number' ? now : 0;
        var it = readItem(id);
        if (!it || !isDue(it, n)) return null;
        var wasRetry = (it.attempts || 0) >= 1;
        var updated = {
          id: it.id, rows: cloneRows(it.rows || []), company: it.company || '',
          createdAt: it.createdAt || 0, attempts: (it.attempts || 0) + 1,
          nextAt: it.nextAt || 0, lastError: it.lastError || '',
          gaveUp: !!it.gaveUp, owner: tabId, claimedAt: n
        };
        var token = tokenOf(updated);
        if (!writeItem(updated)) return null;

        // ★修正ラウンド1・Critical C-1（設計書D2(b)）を1件キーの世界に引き継ぐ:
        // localStorageにはCASが無いため、同じキーへ他タブが同時に書き込んだ
        // 場合の競合はまだ残りうる（既存のリース・初回所有権で通常は起きない
        // 想定だが、二重に確かめる）。書いた直後にもう一度読み直し、自分が
        // 書いたはずの値がまだ残っているか確認する。巻き戻っていたら null を
        // 返して送らせない。
        var after = readItem(id);
        if (!after || after.owner !== tabId || (after.attempts || 0) !== updated.attempts) {
          return null;
        }
        return { token: token, wasRetry: wasRetry };
      },

      markSent: function (id, token) {
        var it = readItem(id);
        if (!it || tokenOf(it) !== String(token)) return false;
        deleteItem(id);
        return true;
      },

      markFailed: function (id, token, message, now) {
        var n = typeof now === 'number' ? now : 0;
        var it = readItem(id);
        if (!it || tokenOf(it) !== String(token)) return false;
        var idx = Math.min(Math.max((it.attempts || 1) - 1, 0), backoffMs.length - 1);
        writeItem({
          id: it.id, rows: cloneRows(it.rows || []), company: it.company || '',
          createdAt: it.createdAt || 0, attempts: it.attempts || 0,
          nextAt: n + backoffMs[idx], lastError: String(message || ''),
          gaveUp: !!it.gaveUp || (it.attempts || 0) >= giveUpAfter,
          owner: it.owner || '', claimedAt: 0
        });
        return true;
      },

      retryNow: function (id, now) {
        var it = readItem(id);
        if (!it) return false;
        writeItem({
          id: it.id, rows: cloneRows(it.rows || []), company: it.company || '',
          createdAt: it.createdAt || 0, attempts: it.attempts || 0,
          nextAt: 0, lastError: it.lastError || '',
          gaveUp: false, owner: tabId, claimedAt: 0
        });
        return true;
      },

      gaveUpCount: function () {
        return listAllItems().filter(function (it) { return !!it.gaveUp; }).length;
      },

      // ★テストから1件単位の入出力を直接検証できるように公開する。
      _internals: {
        readItem: readItem,
        writeItem: writeItem,
        deleteItem: deleteItem,
        listAllItems: listAllItems
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
