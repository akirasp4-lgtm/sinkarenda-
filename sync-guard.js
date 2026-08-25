// sync-guard.js
// index.html / admin.html の refreshInBackground / loadData / loadSettingsMembers が使う
// 「今回・しばらくの間GASを優先すべきか」「単価セルの値が妥当か」の判定ロジックを、
// DOM・fetchに依存しない純粋な形で切り出したもの。
//
// ★なぜ切り出したか（4回目レビュー・Fable 5・Codex両者が独立に指摘）:
//   このロジックはこれまでindex.html/admin.htmlの巨大な<script>内に直接書かれており、
//   「該当するクライアント側の自動テストがリポジトリに存在しない」との指摘を受けた
//   （台帳に「テストで確認済み」と書いていたのは誤りだった）。ここに切り出すことで、
//   ブラウザではそのままグローバル関数として動きつつ、Node/vitestからも同じファイル・
//   同じコードをrequireしてテストできる（実装とテストが同じコードを見る＝
//   テストが「実装と乖離した別物になる」事故を防ぐ）。
//
// ★あえてESモジュール(type="module")にしていない理由:
//   index.html/admin.htmlの既存スクリプトは全て素の<script>（非モジュール）で、
//   file://直接オープンでも動く前提のワンフォルダ構成（single-folder-pwa方針）。
//   type="module"にするとfile://配下でCORSエラーになる環境があり、ロード順序
//   （moduleはdeferと同様に遅延実行される）も既存の同期実行スクリプトと混ざると
//   事故りやすい。このファイル1つのためだけにアプリ全体の読み込み方式を変える
//   のは今回の修正範囲を超えるため、あえてUMD風（<script src>ならグローバルに
//   生える／Node/vitestからはmodule.exports経由でrequireできる）にしてある。
//
// ★2026-08-24 5回目レビュー（Fable 5・Codex 両者が独立に一致）で、createPreferGasTracker
// の設計そのものに重大な欠陥2件が見つかり、作り直した：
//   欠陥1（重大）: mark()は「until=now+5分」を置くだけで、同期の確実成功を一度も
//     確認せずに時間経過だけで失効していた。次のCron巡回（最大5分後＋所要3.9〜56秒）
//     はstickyの5分より遅れうるため、stickyの方が先に切れる。その時点でD1の
//     成功記録は約6分前＝15分の鮮度ガードを通過するため、更新ボタン・会社切替の
//     loadDataが「書き込み前のD1」を正常として読んでしまう。
//     → 直した: 時間だけでは絶対に解除しない。「確実な同期成功を観測した」ときだけ
//     解除する（clear）。時間（PREFER_GAS_CAP_MS）はあくまで上限であり、上限に
//     達しても無条件には解除せず、呼び出し側が/api/healthを見て「markした時刻より
//     後に確実な書き込みがあったか」を確認してから解除する（resolveWithHealthEvidence）。
//   欠陥2（重大）: trackerの状態（until）はページのJSメモリ上だけにあり、
//     タブ間共有も永続化もされていなかった。リロード・新規タブでは常にuntil=0から
//     始まるため、あるタブで書き込み直後の同期が失敗しても、別タブ・再読み込み後の
//     ページはその失敗を知らず、D1が直近成功していれば「書き込み前のD1」を正常と
//     読んでしまう。
//     ★私が旧実装に書いたコメント「タブ間非共有でも実害はD1を早めに疑いすぎる方向＝
//     安全側」は方向が逆だった（両レビュアーが同じ箇所を指摘）。正しくは、
//     markを知らない側のタブがD1を信じすぎる方向に倒れる（＝危険側）。
//     → 直した: 状態をlocalStorageへ永続化し、同一オリジンの全タブ・全ページ
//     （index.html/admin.html）で共有する。読み書きはtry/catchで囲み、
//     localStorageが使えない環境（プライベートブラウジング等）でもページ内メモリの
//     フォールバックで壊れずに動く（ただしその場合はタブ間共有は効かない＝
//     欠陥2を完全には解消できない。これは環境側の制約であり、既知の限界として
//     正直に書いておく）。
//   欠陥3（中・Codex）: 並行するrefreshInBackground呼び出しでは、開始順ではなく
//     完了順の最後の結果が状態を決めていた。同期A（先に開始・低速）が進行中に
//     同期B（後に開始・skippedで即mark）の結果が先に反映され、遅れて返ったAの
//     「確実成功」がBのmarkを消してしまう経路があった。
//     → 直した: mark/clearに「その試行が開始した時刻」（beginAttemptの返り値）を
//     持たせ、clear()は「自分が開始した後に、より新しいmarkが記録されていないか」を
//     確認してからでないと反映しない（＝古い試行の成功が新しい試行のmarkを
//     消せない）。mark()は常に反映する（安全側に倒すだけなので、順序を気にする
//     必要が無い）。
//
// ★2026-08-25 6回目レビュー（Fable 5・Codex 両者が収束・「切り替えてよい」の
// 手順条件として指摘された2件）:
//   修正1（高）: capを超えた後の解除経路resolveWithHealthEvidence()を廃止した。
//     /api/healthのsnapshotAt（GAS取り込みの完了時刻・サーバー時計）とmarkedAt
//     （ブラウザ時計）を直接比較しており、(a)「取得開始がmarkより前だが完了が
//     markより後」の同期でも誤って解除できる、(b) 端末の時計ずれで解除の窓が
//     広がる、という2つの穴があった。代わりに、capに達したときも通常のclear
//     経路と同じ仕組み（実際に/api/syncをPOSTし、decideSyncOutcome().confirmed
//     で判定してからclear(attemptStartedAt)を呼ぶ）に一本化した。これにより
//     時計比較そのものが無くなる（比較する2つの時刻がどちらもブラウザ側の
//     Date.now()になるため）。あわせて、cf/src/sync.jsの「変更なし（ハッシュ
//     一致）によるスキップ」にskipReason:'unchanged'を追加し、decideSyncOutcome
//     がこれも確実成功として扱うようにした（GASが回復した後もデータが一切
//     変化しない静穏期の間ずっとブロックが解けない、という副作用も同時に解消）。
//   修正2（中・Fable 5）: writeState()がsetItem失敗時に何もしていなかったため、
//     「以前は書けていた古い値（例:{markedAt:null}）がstorageに残ったまま」に
//     なり、readState()が常にstorageを優先する結果、同じタブ内の直後の
//     readState()すら最新のmark（メモリには反映済み）を見失っていた（安全側の
//     はずが危険側に反転）。setItem失敗時はstorageのキーを消し、readState()を
//     メモリへフォールバックさせるよう直した。
(function (root) {
  'use strict';

  // ★5回目レビュー修正1: 「時間が経てば無条件解除」をやめたため、この定数は
  // もう「stickyの持続時間」ではなく「時間だけでは絶対に超えて良いブロック期間の
  // 上限（cap）」という意味に変わった。上限に達しても自動では解除されない。
  // ★6回目レビュー修正3（記録の訂正）: capは「ブロックを終わらせる上限」ではない。
  // 呼び出し側（index.html/admin.htmlのloadData）がcapに達したことをきっかけに
  // 「本当に確実成功したか」の再確認（recheck。★6回目レビュー修正1で/api/syncを
  // POSTして直接確かめる方式に変更。旧実装は/api/healthのsnapshotAtを見るだけ
  // だった）を1回試みるだけのトリガーであり、確認できなければ無期限にブロックが
  // 続く。「sticky最大15分」という言い方は、この「無期限にもなりうる」実態を
  // 誤解させるため使わない。
  // 15分＝D1側の鮮度ガード（cf/src/read.jsのFRESHNESS_THRESHOLD_MS）と同じ値。
  // 「D1がまだ正常返却できる程度に新しい」とみなされる期間と揃えてあるだけで、
  // これ自体が解除の根拠にはならない。
  var PREFER_GAS_CAP_MS = 15 * 60 * 1000;

  // localStorageに永続化するときのキー。index.html/admin.htmlの両方から同じキーで
  // 参照するため、片方のページで登録が確実成功しなかった直後にもう片方（別タブ・
  // 別ページ）を開いても、そのブロック状態を引き継げる。
  var STORAGE_KEY = 'yotei-cf-prefer-gas-v1';

  // ★4回目レビュー修正1（必須・両レビュアーが独立に一致）:
  // backend.json自体が取得できなかった（fetch失敗・非OK・JSON壊れ）ときにcfg=nullで
  // 呼ぶ。backendが実際に'd1'かどうか「不明」なため、安全側でtrue（GAS優先）を返す。
  // 旧実装は無条件にfalse（GASを優先しない）を返していたため、loadData側が
  // backend.jsonを再取得し直して2回目だけ成功しbackend:'d1'なら、同期を一度も
  // 実行していないのにD1（書き込み前の古い内容）を読んでしまっていた。
  //
  // cfgが取得できてbackend==='d1'かつworkerUrlありのときはnullを返す＝
  // 「/api/syncの結果を見てから呼び出し側（decideSyncOutcome）が判断すること」を示す。
  //
  // ★5回目レビュー修正8（中・Codex）: 「backend.jsonがtruthyだが未知の形式（{}・
  // 未知のbackend値等）」のとき、旧実装は無条件にfalse（＝「GASと確定」扱い）を
  // 返していた。しかしbackend.jsonが壊れている・部分的にしか書けていない・
  // 将来値が増える等で{}や未知の文字列が来ることはありうり、そのときは本当は
  // 「backendが何なのか不明」というだけで「GASに確定」ではない。
  // → backendが'gas'または'd1'と明確に読み取れた場合だけ「確定」とし、
  // それ以外（undefined・{}・未知の文字列等）は「不明」としてtrue（GAS優先）を返す。
  function forceGasWhenBackendUnknown(cfg) {
    if (!cfg) return true;
    if (cfg.backend === 'd1' && cfg.workerUrl) return null;
    if (cfg.backend === 'd1') return false; // d1指定だがworkerUrl欠落＝実質使えない。D1を試みないので無関係。
    if (cfg.backend === 'gas') return false; // 明確に'gas'と確定。D1は使わないので無関係。
    return true; // backendが'gas'にも'd1'にも明確に読み取れない＝不明。安全側でGAS優先に倒す。
  }

  // /api/syncのHTTPレスポンスとJSON本体から、今回forceGasにすべきか
  // （＝今回は「確実に成功した」と言えないか）を判定する。
  //   resOk: HTTPステータスが2xxだったか
  //   json:  レスポンスボディ（パース失敗・未取得ならnullを渡す）
  // 「確実成功」の定義は3回目レビュー修正2（HTTP 200・JSON.status==='ok'・
  // skippedでない）が基本だが、
  // ★6回目レビュー修正1（高・両者一致・Fable 5推奨）: skipped:trueでも
  // json.skipReason==='unchanged'（GASを実際に取得した結果、既存のD1と
  // ハッシュが完全一致したためのスキップ＝cf/src/sync.js参照）は確実成功として
  // 扱う。GASへ実際に取得しに行き、D1が既に同一内容であることまで確認できて
  // いるため、これを「確実成功でない」として扱い続けるとGAS復旧後もデータが
  // 一切変化しない静穏期の間ずっとブロックが解けない欠陥になる
  // （旧実装のindex.html/admin.htmlのrecheck分岐がこの欠陥を持っていた）。
  // 「進行中のためスキップ」（GASへ一度も取得しに行っていない）にはskipReasonが
  // 付かないため、これまでどおり確実成功として扱わない（安全側を維持）。
  function decideSyncOutcome(resOk, json) {
    var okStatus = !!(resOk && json && json.status === 'ok');
    var confirmedBySkip = !!(json && json.skipped && json.skipReason === 'unchanged');
    var confirmed = okStatus && (!(json && json.skipped) || confirmedBySkip);
    return { forceGas: !confirmed, confirmed: confirmed };
  }

  // localStorageが実際に読み書きできるかを一度だけ試す（存在チェックだけでは
  // 不十分：プライベートブラウジング等では存在するのに書き込みで例外を投げる
  // 実装があるため、実際に試し書きして確認する）。使えなければnullを返す。
  function detectLocalStorage(ls) {
    try {
      if (!ls) return null;
      var testKey = '__sync_guard_test__';
      ls.setItem(testKey, '1');
      ls.removeItem(testKey);
      return ls;
    } catch (e) {
      return null;
    }
  }

  // 「確実に同期が成功するまでD1を読まない」を状態として持続させるための
  // 小さなトラッカー。index.html/admin.htmlそれぞれがこれを1つ持つ。
  //
  // ★5回目レビュー修正2（重大・両者一致）: 状態はlocalStorage（既定のストレージ）へ
  // 永続化する。同一オリジンの全タブ・index.html/admin.htmlの両方で共有される。
  // localStorageが使えない環境ではページ内メモリだけのフォールバックになる
  // （タブ間共有・リロード後の永続化は効かないが、少なくとも例外でページが
  // 壊れることは無い。読み書きは必ずtry/catchで囲む）。
  //
  // ★5回目レビュー修正1（重大・両者一致）: 状態が持つのは「until」ではなく
  // 「markedAt（直近、確実成功でなかった時刻。ブロックしていなければnull）」。
  // isActive/statusは時間だけでfalseへは戻らない。解除は必ずclear()（確実成功を
  // 観測したとき）を経由する。
  // ★6回目レビュー修正1（高・両者一致）: capを超えた後（status()==='recheck'）の
  // 解除経路は、旧実装は専用のresolveWithHealthEvidence()（/api/healthのsnapshotAt
  // ＝サーバー時計と、markedAt＝ブラウザ時計を直接比較）を持っていたが削除した。
  // 「取得完了時刻」と「取得開始時刻」を取り違えると誤解除しうる欠陥があったため
  // （詳細はこの下のPREFER_GAS_CAP_MSコメント）。今は呼び出し側が実際に/api/syncを
  // POSTして直接確かめ、decideSyncOutcome().confirmedがtrueのときだけ通常のclear()
  // を呼ぶ設計に一本化した（index.html/admin.htmlのloadData参照）。clear()自体は
  // markedAtともattemptStartedAtとも「ブラウザ側のDate.now()」だけで比較するため、
  // サーバー時計との突き合わせが原理的に発生しない。
  //
  // ★5回目レビュー修正7（中・Codex）: clear()は「この呼び出しが対応する試行が
  // 開始した時刻」（beginAttemptの返り値）を受け取り、その時刻より後に記録された
  // markがあれば（＝この試行の後により新しい問題が観測されていれば）反映しない。
  // mark()は常に無条件で反映する（安全側に倒すだけなので順序を気にする必要が無い）。
  function createPreferGasTracker(opts) {
    opts = opts || {};
    var capMs = typeof opts.capMs === 'number' ? opts.capMs : PREFER_GAS_CAP_MS;
    var storageKey = opts.storageKey || STORAGE_KEY;
    var storage = opts.storage;
    if (storage === undefined) {
      var candidate = null;
      try {
        if (typeof localStorage !== 'undefined') candidate = localStorage;
      } catch (e) {
        candidate = null;
      }
      storage = detectLocalStorage(candidate);
    } else if (storage) {
      storage = detectLocalStorage(storage);
    }
    var memoryState = { markedAt: null }; // localStorageが無い/壊れているときのフォールバック

    function readState() {
      if (storage) {
        try {
          var raw = storage.getItem(storageKey);
          if (raw) {
            var parsed = JSON.parse(raw);
            if (parsed && (parsed.markedAt === null || typeof parsed.markedAt === 'number')) return parsed;
          }
        } catch (e) {
          // 壊れている・読めない→メモリ側にフォールバック
        }
      }
      return memoryState;
    }
    function writeState(state) {
      memoryState = state;
      if (storage) {
        try {
          storage.setItem(storageKey, JSON.stringify(state));
        } catch (e) {
          // ★6回目レビュー修正2（中・Fable 5）: 旧実装はここで失敗を握りつぶすだけ
          // だった。すると、以前は書けていた古い値（例: 過去にclear()できたときの
          // {markedAt:null}）がstorageに残ったままになり、readState()は常に
          // storageを優先するため、直後のreadState()（同じタブ内の次のloadData等）
          // が古い値を読んでmemoryStateの最新のmark（危険を示す情報）を見失う
          // （安全側のはずが危険側に反転する）。このアプリは同じlocalStorageに
          // 全データキャッシュ（実測700KB≒UTF-16で1.4MB）も書いており、他アプリと
          // 共有するquota（5MB前後）が埋まって setItem だけ失敗するのは現実的に
          // 起こりうる。storageのキーを消してreadState()をメモリへフォールバック
          // させる。removeItem自体が失敗する（quota超過ではなくstorageそのものが
          // 使えない）場合は、以後このtrackerインスタンスではstorageを使わず
          // メモリのみで動かす（このタブ内では、この時点から正しく動く）。
          try {
            storage.removeItem(storageKey);
          } catch (e2) {
            storage = null;
          }
        }
      }
    }

    return {
      // 同期処理を開始する直前に必ず呼ぶ。返り値（開始時刻）をmark/clearへそのまま
      // 渡すこと（欠陥3対策・並行refresh対応）。
      beginAttempt: function (now) {
        return typeof now === 'number' ? now : Date.now();
      },
      // 「確実成功ではなかった」ときに呼ぶ。常に反映する（安全側なので早い者勝ちの
      // 制約は要らない。最悪でも「必要以上に少し長くブロックする」だけで、
      // 危険側には倒れない）。
      mark: function (now) {
        var n = typeof now === 'number' ? now : Date.now();
        writeState({ markedAt: n });
      },
      // 「確実成功だった」ときに呼ぶ。attemptStartedAt（beginAttemptの返り値）を
      // 渡すこと。自分（この試行）が開始した後に記録されたmarkがあれば、この
      // clearは古い情報に基づく可能性があるため反映しない（欠陥3対策）。
      // 戻り値: 実際に解除できたか（true/false）。
      clear: function (attemptStartedAt) {
        var state = readState();
        if (state.markedAt !== null && typeof attemptStartedAt === 'number' && state.markedAt > attemptStartedAt) {
          return false; // より新しいmarkが既にある→この（古い試行の）成功では解除しない
        }
        if (state.markedAt === null) return true; // 既に解除済み
        writeState({ markedAt: null });
        return true;
      },
      // 現在「D1を信用してよいか」を3値で返す:
      //   'trust'   : ブロックしていない。D1を読んでよい。
      //   'block'   : ブロック中（cap未満）。D1を読まずGASを使う。
      //   'recheck' : cap（上限）に達した。時間だけでは解除しない。呼び出し側は
      //               実際に/api/syncをPOSTして確かめ、decideSyncOutcome().confirmed
      //               がtrueのときだけ通常のclear(attemptStartedAt)を呼ぶこと
      //               （★6回目レビュー修正1。確認できなければ実質'block'と同じに
      //               扱ってよい＝呼び出し側は何もしなければ自動的にそうなる）。
      status: function (now) {
        var n = typeof now === 'number' ? now : Date.now();
        var state = readState();
        if (state.markedAt === null) return 'trust';
        var elapsed = n - state.markedAt;
        if (elapsed < capMs) return 'block';
        return 'recheck';
      },
      // 呼び出し側の多くは3値を区別する必要が無いので、単純な真偽値も用意する。
      // 'trust'以外は常にtrue（＝D1を信用しない）を返す＝安全側のデフォルト。
      isActive: function (now) {
        return this.status(now) !== 'trust';
      },
      // ★6回目レビュー修正1（高・両者一致）: 旧実装はここに resolveWithHealthEvidence()
      // （/api/healthのsnapshotAtとmarkedAtを比較して解除する専用メソッド）を持って
      // いたが削除した。理由（2点。詳細はfix-round6-report.md）:
      //   1. snapshotAtは「書き込み完了時刻」であり、GAS取得の「開始時刻」ではない。
      //      「取得開始がmarkより前（＝利用者の書き込みを含まない内容）だが完了が
      //      markより後」という同期でも証拠として通ってしまう欠陥があった。
      //   2. markedAtはブラウザの時計、snapshotAtはサーバーの時計であり、
      //      異なる時計を直接比較していた（端末の時計が遅れているだけで誤解除の
      //      窓が広がる）。
      // 解除はどちらの経路（cap未満のclear／capを超えた後のrecheck）でも同じ
      // clear(attemptStartedAt)だけを使う。呼び出し側は実際に/api/syncをPOSTし、
      // decideSyncOutcome().confirmedがtrueのときだけclear()を呼ぶ（index.html/
      // admin.htmlのloadData参照）。beginAttempt/markedAt/clearはすべてブラウザ側の
      // Date.now()だけで完結するため、サーバー時計との比較が原理的に発生しない。
      markedAtValue: function () { return readState().markedAt; },
      capMsValue: function () { return capMs; }
    };
  }

  // ★4回目レビュー修正6: 単価セルの値が「妥当な数値」かどうかを判定する。
  // Number('   ') は 0 になるため、素朴な Number.isFinite(Number(raw)) だけだと
  // 空白文字だけのセルが「正当な0円」として受理されてしまう（＝取得できなかったのに
  // 編集可能になり、管理者が触ると0が保存される穴）。trimしてから空文字判定する。
  function isRateCellOk(raw) {
    var v = typeof raw === 'string' ? raw.trim() : raw;
    return v !== '' && v !== null && v !== undefined && Number.isFinite(Number(v));
  }

  var api = {
    PREFER_GAS_CAP_MS: PREFER_GAS_CAP_MS,
    forceGasWhenBackendUnknown: forceGasWhenBackendUnknown,
    decideSyncOutcome: decideSyncOutcome,
    createPreferGasTracker: createPreferGasTracker,
    isRateCellOk: isRateCellOk
  };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = api;
  } else {
    for (var k in api) {
      if (Object.prototype.hasOwnProperty.call(api, k)) root[k] = api[k];
    }
  }
})(typeof window !== 'undefined' ? window : (typeof globalThis !== 'undefined' ? globalThis : this));
