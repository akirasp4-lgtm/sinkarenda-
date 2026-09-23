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
//   6. ★2026-09-24: 古いホーム画面アイコンから抜け殻を開くと、iOSは新しい住所を「アプリ内ブラウザ」
//      として開き、住所バー分だけ画面の上が隠れる（iOS側の挙動なので消せない）。抜け殻の standalone
//      起動だけを見分けて印（oldIcon）を運び、新しい住所で1回だけ案内バナーを出す（画面ごとの表示は
//      呼び出し側＝index.html 等が担当。MoveOldIcon / MoveOldIconBanner を見る）。
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
  // standalone＝ホーム画面のアイコンから開いた（抜け殻限定。2026-09-24・古いアイコン対策）
  function planMove(loc, storage, stub, standalone) {
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
    // ★2026-09-24: 古いホーム画面アイコン（抜け殻＝MOVE_STUB）から開くと、新しい住所は
    //   iOSの「アプリ内ブラウザ（Safari View Controller）」扱いになり、住所バーで画面の上が隠れる。
    //   その端末から来たと分かるよう、印を1つだけ載せて運ぶ（新しい住所で案内バナーを出すため）。
    if (stub && standalone) { carry.oldIcon = 1; n++; }
    return NEW_ORIGIN + newPath(loc.pathname) + loc.search +
      (n ? HASH + encodeURIComponent(JSON.stringify(carry)) : '');
  }

  // 新しい住所で開いたとき：# の中に oldIcon の印があるか（受け取り前に確かめる・副作用なし）
  function hasOldIcon(loc) {
    if (!loc || String(loc.hash || '').indexOf(HASH) !== 0) return false;
    try {
      var d = JSON.parse(decodeURIComponent(loc.hash.slice(HASH.length)));
      return !!(d && d.oldIcon);
    } catch (e) { return false; }
  }

  // 古いアイコンの案内バナー：1日に1回まで（本人が「閉じる」を押したら、その日はもう出さない）
  var OLD_ICON_DISMISS_KEY = 'old-icon-banner-dismiss-v1';
  var OLD_ICON_DISMISS_MS = 24 * 60 * 60 * 1000;
  function oldIconDismissed(storage) {
    try {
      var t = Number(storage.getItem(OLD_ICON_DISMISS_KEY));
      return !!t && (Date.now() - t) < OLD_ICON_DISMISS_MS;
    } catch (e) { return false; }
  }
  function dismissOldIcon(storage) {
    try { storage.setItem(OLD_ICON_DISMISS_KEY, String(Date.now())); } catch (e) {}
  }
  function shouldShowOldIconBanner(flag, storage) {
    return !!flag && !oldIconDismissed(storage);
  }
  var oldIconBanner = {
    shouldShow: shouldShowOldIconBanner,
    dismiss: dismissOldIcon,
    dismissed: oldIconDismissed,
    KEY: OLD_ICON_DISMISS_KEY,
  };

  // 新しい住所で開いたとき：持ってきた設定を覚える。受け取り終えたら true（呼ぶ側が住所の # を消す）
  // referrer＝直前のページ。未送信の予定は「細工したリンク」から受け取らない
  //   （細工したリンクを踏ませて、その人の端末から偽の予定を送らせる事故を防ぐ）。
  // ★2026-09-23 Codex検品（2回目）P1: referrer が**空**のとき（Referrer-Policy: no-referrer・noreferrer 付きのリンク等）は、
  //   古い住所から来たのか、細工したリンクなのか区別できない。そこで空のときは送信箱に入れず、
  //   別の置き場（move-quarantine-v1）に取っておく。画面を開いたときに中身（日付・現場や件名・名前）を見せて
  //   「送りますか？」と聞き、本人が OK を押したときだけ送信箱へ移す（promptQuarantine）。
  //   キャンセルなら「捨てますか？」を聞き、捨てなければ次に開いたときにもう一度聞く（予定を黙って消さない）。
  //   ・referrer が古い住所 … 今までどおり送信箱へ（自動で送る）
  //   ・referrer が空       … 取っておく（本人が確かめてから送る）
  //   ・referrer がそれ以外 … 受け取らない（設定だけ受け取る）
  // ★未送信の予定は、置き場（localStorage）に書いて**読み直して同じだったときだけ**受け取り済みにする。
  //   書けなかった（容量・使えない等）ときは false＝住所の # を残し、session（sessionStorage）にも控える
  //   （開き直せばもう一度受け取る。控えは行き先ごとに別のキー）。受け取り終えたら控えを消す。
  var SESSION_KEY = 'move-carry-v1';        // 送信箱へ入れる物（古い住所から）
  var SESSION_Q_KEY = 'move-carry-q-v1';    // 取っておく物（直前のページが分からない）
  var QUARANTINE_KEY = 'move-quarantine-v1';
  function referrerKind(referrer) {
    var r = String(referrer || '');
    if (!r) return 'empty';
    return r.indexOf('https://' + OLD_HOST + '/') === 0 ? 'old' : 'other';
  }
  function referrerOk(referrer) { return referrerKind(referrer) === 'old'; }

  function readQuarantine(storage) {
    try {
      var o = JSON.parse(storage.getItem(QUARANTINE_KEY) || '{}');
      return o && typeof o === 'object' && !Array.isArray(o) ? o : {};
    } catch (e) { return {}; }
  }
  // 書いて読み直して同じなら true。空になったら置き場ごと消す
  function writeQuarantine(storage, q) {
    var n = 0;
    for (var k in q) if (Object.prototype.hasOwnProperty.call(q, k)) n++;
    if (!n) { storage.removeItem(QUARANTINE_KEY); return true; }
    var text = JSON.stringify(q);
    storage.setItem(QUARANTINE_KEY, text);
    return storage.getItem(QUARANTINE_KEY) === text;
  }

  // mode: 'queue'（送信箱へ）/ 'quarantine'（取っておく）/ 'none'（未送信は受け取らない）
  function applyCarry(d, storage, mode) {
    for (var i = 0; i < CARRY_KEYS.length; i++) {
      var k = CARRY_KEYS[i];
      // 設定が書けなくても、未送信の予定の受け取りは続ける
      try { if (typeof d[k] === 'string' && d[k] && !storage.getItem(k)) storage.setItem(k, d[k]); } catch (e) {}
    }
    if (mode !== 'queue' && mode !== 'quarantine') return true;
    var allIn = true;
    var q = mode === 'quarantine' ? readQuarantine(storage) : null;
    var qAdded = false;
    for (var key in d) {
      if (!Object.prototype.hasOwnProperty.call(d, key) || !isPendingKey(key)) continue;
      if (typeof d[key] !== 'string' || storage.getItem(key)) continue;
      try { JSON.parse(d[key]); } catch (e) { continue; }   // 壊れた物で箱を詰まらせない
      if (q) {
        if (q[key] !== d[key]) { q[key] = d[key]; qAdded = true; }
        continue;
      }
      try { storage.setItem(key, d[key]); } catch (e) { /* 下の読み直しで分かる */ }
      var back = null;
      try { back = storage.getItem(key); } catch (e) { back = null; }
      if (back !== d[key]) allIn = false;
    }
    if (q && qAdded) {
      try { allIn = writeQuarantine(storage, q); } catch (e) { allIn = false; }
    }
    return allIn;
  }
  function receive(loc, storage, referrer, session) {
    if (loc.hash.indexOf(HASH) === 0) {
      var raw = loc.hash.slice(HASH.length);
      var d = JSON.parse(decodeURIComponent(raw));
      var kind = referrerKind(referrer);
      var mode = kind === 'old' ? 'queue' : kind === 'empty' ? 'quarantine' : 'none';
      var sk = mode === 'queue' ? SESSION_KEY : mode === 'quarantine' ? SESSION_Q_KEY : '';
      if (sk && session) { try { session.setItem(sk, raw); } catch (e) {} }
      var ok = applyCarry(d, storage, mode);
      if (ok && sk && session) { try { session.removeItem(sk); } catch (e) {} }
      return ok;
    }
    // # はもう消えたが、前回受け取りきれなかった控えが残っている
    //   （控えは行き先を決めてから書いた物なので、確かめ直さずにその行き先へ）
    if (session) {
      var pairs = [[SESSION_KEY, 'queue'], [SESSION_Q_KEY, 'quarantine']];
      for (var i = 0; i < pairs.length; i++) {
        var r2 = '';
        try { r2 = session.getItem(pairs[i][0]) || ''; } catch (e) { r2 = ''; }
        if (!r2) continue;
        var d2 = null;
        try { d2 = JSON.parse(decodeURIComponent(r2)); } catch (e) { d2 = null; }
        if (!d2 || applyCarry(d2, storage, pairs[i][1])) { try { session.removeItem(pairs[i][0]); } catch (e) {} }
      }
    }
    return false;
  }

  // ------------------------------------------------------------
  // 取っておいた未送信（move-quarantine-v1）を画面で確かめる
  // ------------------------------------------------------------
  // prefixes＝その画面の送信箱の頭（例 'yotei-pending-add-v1:'）。ほかの画面の分は触らない
  function matches(key, prefixes) {
    for (var j = 0; j < prefixes.length; j++) if (String(key).indexOf(prefixes[j]) === 0) return true;
    return false;
  }
  function short(v, n) {
    var t = String(v == null ? '' : v).replace(/\s+/g, ' ').trim();
    return t.length > n ? t.slice(0, n) + '…' : t;
  }
  // 1件の中身（日付・現場や件名・名前）。読めない物も「中身が読めない予定」として数える（黙って消さない）
  function describeItem(value) {
    var it = null;
    try { it = JSON.parse(value); } catch (e) { it = null; }
    var rows = it && Array.isArray(it.rows) ? it.rows : [];
    // ★2026-09-23 検品3回目の指摘：最初の行だけを見せて、見えていない行まで送っていた。
    //   すべての行の「日付・現場（件名）」を重ねずに並べ、名前もすべて出す（送る物＝見せた物）。
    var parts = [], seen = {}, names = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i] && typeof rows[i] === 'object' ? rows[i] : {};
      var p = [short(r.date || r.startDate || '', 20), short(r.title || r.genba || r.loc || '', 40)].filter(Boolean).join(' ');
      if (p && !seen[p]) { seen[p] = 1; parts.push(p); }
      var nm = r.name ? short(r.name, 20) : '';
      if (nm && names.indexOf(nm) < 0) names.push(nm);
    }
    return { parts: parts, names: names };
  }
  function quarantineList(storage, prefixes) {
    var q = readQuarantine(storage);
    var out = [];
    for (var k in q) {
      if (!Object.prototype.hasOwnProperty.call(q, k) || !matches(k, prefixes)) continue;
      var info = describeItem(q[k]);
      info.key = k;
      out.push(info);
    }
    return out;
  }
  // 送信箱へ移す（本人が OK を押したとき）。移せた件数を返す。書けなかった物は置き場に残す
  // onlyKeys を渡したときは、その番号だけを移す（確認画面に出した物だけ＝見せていない物は送らない）
  function releaseQuarantine(storage, prefixes, onlyKeys) {
    var q = readQuarantine(storage);
    var n = 0;
    for (var k in q) {
      if (!Object.prototype.hasOwnProperty.call(q, k) || !matches(k, prefixes)) continue;
      if (onlyKeys && onlyKeys.indexOf(k) < 0) continue;
      var have = null;
      try { have = storage.getItem(k); } catch (e) { have = null; }
      if (have == null) {
        try { storage.setItem(k, q[k]); } catch (e) {}
        var back = null;
        try { back = storage.getItem(k); } catch (e) { back = null; }
        if (back !== q[k]) continue;   // 書けなかった＝置き場に残す（次に開いたときにもう一度聞く）
        n++;
      }
      delete q[k];   // 送信箱に入った（同じ番号が既に入っていた物は、そちらが送る）
    }
    try { writeQuarantine(storage, q); } catch (e) {}
    return n;
  }
  function discardQuarantine(storage, prefixes) {
    var q = readQuarantine(storage);
    for (var k in q) {
      if (Object.prototype.hasOwnProperty.call(q, k) && matches(k, prefixes)) delete q[k];
    }
    try { writeQuarantine(storage, q); } catch (e) {}
  }
  var QUARANTINE_SHOW_MAX = 10;   // 1回に確認する件数。残りは次に開いたときにもう一度聞く
  function quarantineMessage(list, total) {
    var all = total == null ? list.length : total;
    var head = all > list.length
      ? '前の住所に残っていた未送信の予定が' + all + '件あります。まず下の' + list.length + '件を送りますか？（残りは次に開いたときに聞きます）'
      : '前の住所に残っていた未送信の予定が' + list.length + '件あります。送りますか？';
    var lines = [head, ''];
    for (var i = 0; i < list.length; i++) {
      var it = list[i];
      var body = it.parts.length ? it.parts.join('／') : '中身が読めない予定';
      if (it.names.length) body += '（' + it.names.join('、') + '）';
      lines.push('・' + body);
    }
    lines.push('', '自分で入れた覚えのない予定なら「キャンセル」を押してください。');
    return lines.join('\n');
  }
  var DISCARD_MESSAGE = 'この未送信の予定を捨てますか？\n\n［OK］捨てる\n［キャンセル］捨てずに残す（次に開いたときにもう一度聞きます）';
  // 戻り値: 'none'（何も無い）/ 'sent'（送信箱へ移した）/ 'discarded'（捨てた）/ 'kept'（残した）
  function promptQuarantine(storage, prefixes, confirmFn) {
    var all = quarantineList(storage, prefixes);
    if (!all.length) return 'none';
    var shown = all.slice(0, QUARANTINE_SHOW_MAX);
    var keys = shown.map(function (it) { return it.key; });
    if (confirmFn(quarantineMessage(shown, all.length))) {
      releaseQuarantine(storage, prefixes, keys);
      return 'sent';
    }
    if (confirmFn(DISCARD_MESSAGE)) {
      discardQuarantine(storage, prefixes);
      return 'discarded';
    }
    return 'kept';
  }

  var quarantine = {
    list: quarantineList, release: releaseQuarantine, discard: discardQuarantine, message: quarantineMessage,
    prompt: function (storage, prefixes, confirmFn) {
      return promptQuarantine(storage, prefixes, confirmFn || function (m) { return root.confirm(m); });
    },
    KEY: QUARANTINE_KEY,
    DISCARD_MESSAGE: DISCARD_MESSAGE,
  };
  var api = {
    planMove: planMove, receive: receive, newPath: newPath, CARRY_KEYS: CARRY_KEYS, quarantine: quarantine,
    hasOldIcon: hasOldIcon, oldIconBanner: oldIconBanner,
  };
  if (typeof module !== 'undefined' && module.exports) { module.exports = api; return; }
  // 送信箱のある画面が開いたときに呼ぶ: MoveQuarantine.prompt(localStorage, [その画面の箱の頭])
  root.MoveQuarantine = quarantine;
  // 古いアイコンの案内バナーが呼ぶ: MoveOldIconBanner.shouldShow(window.MoveOldIcon, localStorage)
  root.MoveOldIconBanner = oldIconBanner;

  try {
    var ls = root.localStorage;
    // ホーム画面のアイコン（またはPWAインストール済み）から開いたか。install-banner と同じ判定を使う
    var standalone = !!(root.navigator && root.navigator.standalone) ||
      !!(root.matchMedia && root.matchMedia('(display-mode: standalone)').matches);
    var to = planMove(root.location, ls, !!root.MOVE_STUB, standalone);
    if (to) { root.location.replace(to); return; }
    // ★2026-09-24: # を受け取る（＝消す）前に、oldIcon の印を読んでおく。
    //   印があるときだけ <html> に in-app-view を付ける（iOSのアプリ内ブラウザ用の余白調整。他の人には付けない）。
    var oldIcon = hasOldIcon(root.location);
    root.MoveOldIcon = oldIcon;
    if (oldIcon) {
      try { root.document.documentElement.classList.add('in-app-view'); } catch (e) {}
    }
    var ss = null;
    try { ss = root.sessionStorage; } catch (e) { ss = null; }
    if (receive(root.location, ls, root.document && root.document.referrer, ss)) {
      root.history.replaceState(null, '', root.location.pathname + root.location.search);
    }
  } catch (e) {}
})(typeof window !== 'undefined' ? window : this);
