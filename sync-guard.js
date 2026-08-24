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
(function (root) {
  'use strict';

  // ★4回目レビュー修正2（強く推奨）: forceGasはrefreshInBackground直後の1回だけしか
  // 効かない引数だった。同期が失敗した後、次に確実な同期成功（Cronは最大5分間隔）が
  // 起きるまでの間に「更新」ボタン・プル更新・会社切替が挟まると、それらが呼ぶ素の
  // loadData()（forceGas無し）がD1（書き込み前の古い内容）を読んでしまう。
  // 5分＝Cronの巡回間隔。次の巡回で確実成功するまで持てば十分という意図。
  var PREFER_GAS_STICKY_MS = 5 * 60 * 1000;

  // ★4回目レビュー修正1（必須・両レビュアーが独立に一致）:
  // backend.json自体が取得できなかった（fetch失敗・非OK・JSON壊れ）ときにcfg=nullで
  // 呼ぶ。backendが実際に'd1'かどうか「不明」なため、安全側でtrue（GAS優先）を返す。
  // 旧実装は無条件にfalse（GASを優先しない）を返していたため、loadData側が
  // backend.jsonを再取得し直して2回目だけ成功しbackend:'d1'なら、同期を一度も
  // 実行していないのにD1（書き込み前の古い内容）を読んでしまっていた。
  //
  // cfgが取得できてbackend==='d1'かつworkerUrlありのときはnullを返す＝
  // 「/api/syncの結果を見てから呼び出し側（decideSyncOutcome）が判断すること」を示す。
  // それ以外（backendがd1でないと確定・またはd1指定でもworkerUrl欠落で実質使えない）は
  // false＝loadData側もD1を候補に入れないのでforceGasは無関係。
  function forceGasWhenBackendUnknown(cfg) {
    if (!cfg) return true;
    if (cfg.backend === 'd1' && cfg.workerUrl) return null;
    return false;
  }

  // /api/syncのHTTPレスポンスとJSON本体から、今回forceGasにすべきか
  // （＝今回は「確実に成功した」と言えないか）を判定する。
  //   resOk: HTTPステータスが2xxだったか
  //   json:  レスポンスボディ（パース失敗・未取得ならnullを渡す）
  // 「確実成功」の定義は3回目レビュー修正2から変わっていない
  // （HTTP 200・JSON.status==='ok'・skippedでない）。
  function decideSyncOutcome(resOk, json) {
    var confirmed = !!(resOk && json && json.status === 'ok' && !json.skipped);
    return { forceGas: !confirmed, confirmed: confirmed };
  }

  // 「確実に同期が成功するまでD1を読まない」を状態として持続させるための
  // 小さなトラッカー。index.html/admin.htmlそれぞれが自分専用のインスタンスを1つ持つ。
  // ★複数タブについて: このトラッカーはページ（タブ）ごとのJSインスタンスであり、
  // タブ間では共有されない。あるタブで同期が失敗しても、別タブは自分がloadData/
  // refreshInBackgroundを呼ぶまでその失敗を知らない。これは旧forceGas（引数）でも
  // 同じ制約だったため、今回の変更による新しい劣化ではないが、解消もしていない
  // （既知の限界。localStorage等でタブ間共有する案もあるが、書き込みは常にGASであり
  // 読み取り専用の安全弁のため、タブごとの独立動作でも実害は「D1を早めに疑いすぎる」
  // 方向にしか倒れない＝安全側）。
  function createPreferGasTracker(stickyMs) {
    var ms = typeof stickyMs === 'number' ? stickyMs : PREFER_GAS_STICKY_MS;
    var until = 0;
    return {
      mark: function (now) { until = (typeof now === 'number' ? now : Date.now()) + ms; },
      clear: function () { until = 0; },
      isActive: function (now) { return (typeof now === 'number' ? now : Date.now()) < until; },
      untilValue: function () { return until; }
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
    PREFER_GAS_STICKY_MS: PREFER_GAS_STICKY_MS,
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
