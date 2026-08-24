import { describe, it, expect } from 'vitest';
import { createRequire } from 'node:module';

// ★4回目レビュー: index.html/admin.htmlのrefreshInBackground/loadData/
// loadSettingsMembersが使う判定ロジックは、これまでリポジトリに自動テストが
// 存在しなかった（レビュアー指摘・台帳の「テストで確認済み」記載は誤りだった）。
// 修正1・2・6でこのロジックをプロジェクト直下の sync-guard.js に切り出した
// （ブラウザでは素の<script src>でグローバル関数として動く・Node/vitestからは
// module.exports経由でrequireできるUMD風の作り）。ここではNode側からその同じ
// ファイルを直接requireしてテストする＝実装とテストが同じコードを見ているため、
// 「実装を変えたのにテストは古いまま」というズレが原理的に起きない。
//
// cfはpackage.jsonで"type":"module"だが、sync-guard.js自体はimport/export構文を
// 使わない素のスクリプトなので、createRequireで普通にrequireできる。
const require = createRequire(import.meta.url);
const SG = require('../../sync-guard.js');

describe('sync-guard.js: forceGasWhenBackendUnknown（4回目レビュー修正1・必須）', () => {
  it('backend.json自体が取得できなかった（cfg=null＝fetch失敗・非OK・JSON壊れ）ときはtrue（GAS優先）を返す', () => {
    // ★これが今回の必須修正の本体。旧実装はここでfalseを返しており、loadData側が
    // backend.jsonを再取得し直して2回目だけ成功しbackend:'d1'なら、同期を一度も
    // 実行していないのにD1（書き込み前の古い内容）を読んでしまう欠陥があった。
    expect(SG.forceGasWhenBackendUnknown(null)).toBe(true);
  });

  it('backendが\'d1\'でないと確定した場合はfalseを返す（loadData側もD1を読みに行かないので無関係）', () => {
    expect(SG.forceGasWhenBackendUnknown({ backend: 'gas' })).toBe(false);
    expect(SG.forceGasWhenBackendUnknown({})).toBe(false);
  });

  it('backend===\'d1\'だがworkerUrlが無い（実質使えない）場合もfalseを返す（loadData側もD1を候補に入れない）', () => {
    expect(SG.forceGasWhenBackendUnknown({ backend: 'd1' })).toBe(false);
  });

  it('backend===\'d1\'かつworkerUrlありのときはnullを返す（＝/api/syncの結果で呼び出し側が判断する合図）', () => {
    expect(SG.forceGasWhenBackendUnknown({ backend: 'd1', workerUrl: 'https://example.test' })).toBeNull();
  });
});

describe('sync-guard.js: decideSyncOutcome（同期結果からforceGas/confirmedを決める）', () => {
  it('HTTP 200・status:ok・skippedでない＝確実成功。forceGas:false, confirmed:true', () => {
    const out = SG.decideSyncOutcome(true, { status: 'ok', skipped: false });
    expect(out).toEqual({ forceGas: false, confirmed: true });
  });

  it('skipped:trueのときは確実成功ではない（forceGas:true）', () => {
    const out = SG.decideSyncOutcome(true, { status: 'ok', skipped: true });
    expect(out.confirmed).toBe(false);
    expect(out.forceGas).toBe(true);
  });

  it('status:errorのときは確実成功ではない', () => {
    const out = SG.decideSyncOutcome(true, { status: 'error', message: '急減ガードで拒否' });
    expect(out.confirmed).toBe(false);
    expect(out.forceGas).toBe(true);
  });

  it('HTTPエラー（res.ok:false）のときは確実成功ではない', () => {
    const out = SG.decideSyncOutcome(false, { status: 'error' });
    expect(out.confirmed).toBe(false);
    expect(out.forceGas).toBe(true);
  });

  it('通信例外・JSONパース失敗（jsonがnull）のときも確実成功ではない', () => {
    const out = SG.decideSyncOutcome(false, null);
    expect(out.confirmed).toBe(false);
    expect(out.forceGas).toBe(true);
  });
});

describe('sync-guard.js: createPreferGasTracker（4回目レビュー修正2・強く推奨）', () => {
  it('mark()するとその直後はisActive()がtrueになる', () => {
    const tracker = SG.createPreferGasTracker(5 * 60 * 1000);
    const now = 1_000_000;
    tracker.mark(now);
    expect(tracker.isActive(now)).toBe(true);
    expect(tracker.isActive(now + 1000)).toBe(true);
  });

  it('sticky期間（例:5分）を過ぎるとisActive()がfalseに戻る＝Cronの次巡回で自然に解除される', () => {
    const tracker = SG.createPreferGasTracker(5 * 60 * 1000);
    const now = 1_000_000;
    tracker.mark(now);
    expect(tracker.isActive(now + 5 * 60 * 1000 - 1)).toBe(true);
    expect(tracker.isActive(now + 5 * 60 * 1000)).toBe(false);
  });

  it('clear()すると即座にisActive()がfalseになる（＝確実成功を検知したとき）', () => {
    const tracker = SG.createPreferGasTracker(5 * 60 * 1000);
    const now = 1_000_000;
    tracker.mark(now);
    expect(tracker.isActive(now)).toBe(true);
    tracker.clear();
    expect(tracker.isActive(now)).toBe(false);
  });

  it('mark()を繰り返す（同期が連続して失敗する）と保護期間が都度延長される', () => {
    const tracker = SG.createPreferGasTracker(5 * 60 * 1000);
    let now = 1_000_000;
    tracker.mark(now);
    now += 4 * 60 * 1000; // 4分後、まだ有効
    expect(tracker.isActive(now)).toBe(true);
    tracker.mark(now); // ここでまた失敗＝延長
    expect(tracker.isActive(now + 4 * 60 * 1000)).toBe(true); // 延長前なら切れていたはずの時刻でもまだ有効
  });

  it('「preferGasが立っている間はforceGas引数無しの呼び出しもGASを使う」を再現する（loadDataのeffectiveForceGas算出と同じ式）', () => {
    // index.html/admin.htmlのloadData()は
    //   const effectiveForceGas = !!forceGas || preferGasTracker.isActive();
    // という1行でこのトラッカーを参照する。「更新」ボタン・プル更新・switchCompany等は
    // forceGas引数を渡さない（undefined）が、trackerがアクティブな間は
    // effectiveForceGasがtrueになる＝D1を候補から外しGASを使う、という点をここで確認する。
    const tracker = SG.createPreferGasTracker(5 * 60 * 1000);
    const now = 2_000_000;
    tracker.mark(now); // 直前の同期が確実成功でなかった

    const forceGasArgFromButton = undefined; // 「更新」ボタン等はforceGasを渡さない
    const effectiveForceGas = !!forceGasArgFromButton || tracker.isActive(now + 1000);
    expect(effectiveForceGas).toBe(true);

    tracker.clear(); // 次のCron巡回で確実成功した
    const effectiveForceGasAfterClear = !!forceGasArgFromButton || tracker.isActive(now + 2000);
    expect(effectiveForceGasAfterClear).toBe(false);
  });
});

describe('sync-guard.js: isRateCellOk（4回目レビュー修正6）', () => {
  it('空白文字だけのセルは妥当な値として受理しない（Number(\'   \')===0になる罠の対策）', () => {
    expect(SG.isRateCellOk('   ')).toBe(false);
    expect(SG.isRateCellOk('　')).toBe(false); // 全角スペースもtrim対象
  });

  it('本当に空のセル（空文字・null・undefined）も受理しない', () => {
    expect(SG.isRateCellOk('')).toBe(false);
    expect(SG.isRateCellOk(null)).toBe(false);
    expect(SG.isRateCellOk(undefined)).toBe(false);
  });

  it('0は妥当な数値として受理する（無給ではなく「未取得」だけを区別する）', () => {
    expect(SG.isRateCellOk(0)).toBe(true);
    expect(SG.isRateCellOk('0')).toBe(true);
  });

  it('通常の単価（数値・前後空白付き数値文字列）は受理する', () => {
    expect(SG.isRateCellOk(1500)).toBe(true);
    expect(SG.isRateCellOk('1500')).toBe(true);
    expect(SG.isRateCellOk(' 1500 ')).toBe(true);
  });

  it('非数値の文字列は受理しない', () => {
    expect(SG.isRateCellOk('未設定')).toBe(false);
    expect(SG.isRateCellOk('abc')).toBe(false);
  });
});
