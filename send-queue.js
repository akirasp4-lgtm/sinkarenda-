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
      // ★レビュー修正ラウンド1・Important I-2: 一度usableがfalseに落ちたタブは、
      // その後storageへ一切書き込まない（メモリのみで動く）。usable===falseのまま
      // storageへの書き込みを試み続けると、このタブが持つ古い（storageと食い違った）
      // memoryの内容を、storageが回復した後に上書きしてしまい、その間に別タブが
      // 正しく書き込んだ未送信を消してしまう（実測済みの事故）。
      if (!storage || !usable) return false;
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
      var dirty = fn(st);
      // ★レビュー修正ラウンド1・Minor M-1: fnが実際に書き換えた（truthyを返した）
      // ときだけ書く。「見つからない・変化なし」の呼び出しでも無条件に書いていると、
      // 読んでから書くまでの間に別タブが割り込める窓（C-1で防ごうとしている交錯）を
      // 無駄に広げてしまう。beginSendの永続化ゲート（書けなかったらnull）は
      // dirtyがtrueのときのwriteStateの成否で判定されるため、この変更では壊れない。
      var ok = dirty ? writeState(st) : true;
      return { ok: ok, ret: dirty };
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

    return {
      isStorageUsable: function () { return usable; },

      enqueue: function (item, now) {
        if (!item) return false;
        var id = String(item.id || '');
        var rows = Array.isArray(item.rows) ? item.rows : [];
        if (!id || rows.length === 0) return false;
        var accepted = false;
        mutate(function (st) {
          if (st.items.length >= maxItems) { accepted = false; return false; }
          st.items.push({
            id: id, rows: cloneRows(rows), company: String(item.company || ''),
            createdAt: typeof now === 'number' ? now : 0,
            attempts: 0, nextAt: 0, lastError: '', gaveUp: false,
            owner: tabId, claimedAt: 0
          });
          accepted = true;
          return true;
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

      nextDue: function (now) {
        var n = typeof now === 'number' ? now : 0;
        var st = readState();
        for (var i = 0; i < st.items.length; i++) {
          if (isDue(st.items[i], n)) return copyItems([st.items[i]])[0];
        }
        return null;
      },

      // ★fetchの前に attempts を +1 して永続化する（設計書D9）。
      // Task 1で導入した mutate() を通す＝「直前に読み直してから変更して書く」ため、
      // 別タブが後から入れた未送信を巻き込まずに済む（fnの中でawaitは挟まない）。
      // 永続化できなければ null を返して送らせない（記録が残らない以上、
      // 次回に再送扱いへ倒せず、二重登録の危険が残るため。mutate().ok で判定する）。
      beginSend: function (id, now) {
        var n = typeof now === 'number' ? now : 0;
        var found = false;
        var token = null;
        var wasRetry = false;
        var expectedOwner = tabId;
        var expectedAttempts = 0;
        var m = mutate(function (st) {
          var it = findItem(st, id);
          if (!it || !isDue(it, n)) return false;
          found = true;
          wasRetry = (it.attempts || 0) >= 1;
          it.attempts = (it.attempts || 0) + 1;
          it.owner = tabId;
          it.claimedAt = n;
          token = tokenOf(it);
          expectedAttempts = it.attempts;
          return true;
        });
        if (!found || !m.ok) return null;

        // ★レビュー修正ラウンド1・Critical C-1（設計書D2(b)）: localStorageには
        // 比較交換（CAS）が無いため、上のmutateが書いた直後に、別タブが「読んで
        // からこの書き込みより前に取っていた古いスナップショット」をそのまま
        // 書き戻すと、attempts/ownerがこの書き込み以前の値に巻き戻ることがある。
        // 書いて終わりにせず、もう一度storageから読み直して「自分が書いたはずの
        // 値」がまだ残っているか確認する。巻き戻っていたら null を返して送らせない
        // （＝attemptsが0のまま拾われてwasRetry:falseになり、二重登録に至る事故を
        // 防ぐ）。
        var verify = readState();
        var after = findItem(verify, id);
        if (!after || after.owner !== expectedOwner || (after.attempts || 0) !== expectedAttempts) {
          return null;
        }
        return { token: token, wasRetry: wasRetry };
      },

      markSent: function (id, token) {
        var found = false;
        mutate(function (st) {
          var it = findItem(st, id);
          if (!it || tokenOf(it) !== String(token)) return false;
          found = true;
          st.items = st.items.filter(function (x) { return x !== it; });
          return true;
        });
        return found;
      },

      markFailed: function (id, token, message, now) {
        var n = typeof now === 'number' ? now : 0;
        var found = false;
        mutate(function (st) {
          var it = findItem(st, id);
          if (!it || tokenOf(it) !== String(token)) return false;
          found = true;
          var idx = Math.min(Math.max((it.attempts || 1) - 1, 0), backoffMs.length - 1);
          it.nextAt = n + backoffMs[idx];
          it.lastError = String(message || '');
          it.claimedAt = 0;
          if ((it.attempts || 0) >= giveUpAfter) it.gaveUp = true;
          return true;
        });
        return found;
      },

      retryNow: function (id, now) {
        var n = typeof now === 'number' ? now : 0;
        var found = false;
        mutate(function (st) {
          var it = findItem(st, id);
          if (!it) return false;
          found = true;
          it.gaveUp = false;
          it.nextAt = 0;
          it.claimedAt = 0;
          it.owner = tabId;
          return true;
        });
        return found;
      },

      gaveUpCount: function () {
        return readState().items.filter(function (it) { return !!it.gaveUp; }).length;
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
