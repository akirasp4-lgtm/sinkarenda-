// 引っ越し：古い住所で開いた人を、新しい住所へ自動で移す（2026-09-23・設計書D）
//
// 社員はURLを貼り替えなくていい。いつものブックマーク・ホーム画面のアイコンから開けば、
// 新しい住所 https://yotei.glorise.net へ移り、端末に覚えている設定（名前・会社・拠点）も持っていく。
//
// 守っていること:
//   1. 未送信の予定が端末に残っている間は移さない（古い住所の箱に置き去りにしない）。
//      送り終われば、次に開いたときに移る。
//   2. 持っていくのは設定だけ。URLの # の後ろに入れるので、サーバーには送られない。
//      受け取ったら # はすぐ消す。新しい住所で既に覚えている値は上書きしない。
//   3. 逃げ道：古い住所に ?stay=1 を付けて開くと、その回は移らない。
//      全体を止めるときは MOVE_TO_NEW = false にして push する（GitHub側の配信が変わる）。
//   4. 何かで失敗しても画面は止めない（全部 try の中）。
//   5. ★転送専用ページ（GitHub非公開化の後に古い住所へ置く抜け殻・window.MOVE_STUB=true）では、
//      未送信の予定も一緒に運ぶ（抜け殻には送る仕組みが無いため。新しい住所の箱に入れれば自動で送られる）。
(function (root) {
  'use strict';

  var MOVE_TO_NEW = true;
  var OLD_HOST = 'akirasp4-lgtm.github.io';
  var OLD_BASE = '/sinkarenda-/';
  var NEW_ORIGIN = 'https://yotei.glorise.net';
  var CARRY_KEYS = ['genba-c-param', 'genba-username', 'genba-company', 'genba-kyoten', 'admin-username'];
  var PENDING_PREFIXES = ['yotei-pending-add-v1:', 'pres-pending-add-v1:'];
  var HASH = '#mv=';

  function isPendingKey(k) {
    for (var j = 0; j < PENDING_PREFIXES.length; j++) {
      if (String(k).indexOf(PENDING_PREFIXES[j]) === 0) return true;
    }
    return false;
  }

  function pendingKeys(storage) {
    var out = [];
    for (var i = 0; i < storage.length; i++) {
      var k = storage.key(i) || '';
      if (isPendingKey(k)) out.push(k);
    }
    return out;
  }

  // /sinkarenda-/admin.html → /admin、/sinkarenda-/ と index.html → /
  // （新しい置き場は .html を外した形が正。付けたままだと308で1回余計に回る）
  function newPath(pathname) {
    var p = pathname.indexOf(OLD_BASE) === 0 ? pathname.slice(OLD_BASE.length) : pathname.replace(/^\//, '');
    p = p.replace(/\.html$/, '');
    if (p === 'index') p = '';
    return '/' + p;
  }

  // 古い住所で開いたとき：移し先のURLを返す（移さないなら null）
  function planMove(loc, storage, stub) {
    if (!MOVE_TO_NEW || loc.hostname !== OLD_HOST) return null;
    if (/[?&]stay=1(&|$)/.test(loc.search)) return null;
    var pend = pendingKeys(storage);
    if (pend.length && !stub) return null;   // 本物の画面なら、ここで送り終えてから移す
    var carry = {}, n = 0;
    for (var i = 0; i < CARRY_KEYS.length; i++) {
      var v = storage.getItem(CARRY_KEYS[i]);
      if (v) { carry[CARRY_KEYS[i]] = v; n++; }
    }
    for (var p = 0; p < pend.length; p++) { carry[pend[p]] = storage.getItem(pend[p]); n++; }
    return NEW_ORIGIN + newPath(loc.pathname) + loc.search +
      (n ? HASH + encodeURIComponent(JSON.stringify(carry)) : '');
  }

  // 新しい住所で開いたとき：持ってきた設定を覚える。受け取ったら true
  // referrer＝直前のページ。未送信の予定は「古い住所の抜け殻から来た」ときだけ受け取る
  //   （細工したリンクを踏ませて、その人の端末から偽の予定を送らせる事故を防ぐ）。
  function receive(loc, storage, referrer) {
    if (loc.hash.indexOf(HASH) !== 0) return false;
    var d = JSON.parse(decodeURIComponent(loc.hash.slice(HASH.length)));
    for (var i = 0; i < CARRY_KEYS.length; i++) {
      var k = CARRY_KEYS[i];
      if (typeof d[k] === 'string' && d[k] && !storage.getItem(k)) storage.setItem(k, d[k]);
    }
    // 抜け殻から運ばれた未送信の予定。中身がJSONとして読めるものだけ箱に入れる（壊れた物で箱を詰まらせない）
    var fromOld = String(referrer || '').indexOf('https://' + OLD_HOST + '/') === 0;
    for (var key in d) {
      if (!fromOld) break;
      if (!Object.prototype.hasOwnProperty.call(d, key) || !isPendingKey(key)) continue;
      if (typeof d[key] !== 'string' || storage.getItem(key)) continue;
      try { JSON.parse(d[key]); } catch (e) { continue; }
      storage.setItem(key, d[key]);
    }
    return true;
  }

  var api = { planMove: planMove, receive: receive, newPath: newPath, CARRY_KEYS: CARRY_KEYS };
  if (typeof module !== 'undefined' && module.exports) { module.exports = api; return; }

  try {
    var ls = root.localStorage;
    var to = planMove(root.location, ls, !!root.MOVE_STUB);
    if (to) { root.location.replace(to); return; }
    if (receive(root.location, ls, root.document && root.document.referrer)) {
      root.history.replaceState(null, '', root.location.pathname + root.location.search);
    }
  } catch (e) {}
})(typeof window !== 'undefined' ? window : this);
