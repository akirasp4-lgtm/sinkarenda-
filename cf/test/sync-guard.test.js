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

describe('sync-guard.js: forceGasWhenBackendUnknown（4回目レビュー修正1・必須／5回目レビュー修正8）', () => {
  it('backend.json自体が取得できなかった（cfg=null＝fetch失敗・非OK・JSON壊れ）ときはtrue（GAS優先）を返す', () => {
    // ★これが今回の必須修正の本体。旧実装はここでfalseを返しており、loadData側が
    // backend.jsonを再取得し直して2回目だけ成功しbackend:'d1'なら、同期を一度も
    // 実行していないのにD1（書き込み前の古い内容）を読んでしまう欠陥があった。
    expect(SG.forceGasWhenBackendUnknown(null)).toBe(true);
  });

  it('backendが明確に\'gas\'と読み取れた場合はfalseを返す（loadData側もD1を読みに行かないので無関係）', () => {
    expect(SG.forceGasWhenBackendUnknown({ backend: 'gas' })).toBe(false);
  });

  it('5回目レビュー修正8（Codex・中）: backend.jsonがtruthyだが未知の形式（{}・未知のbackend値）はtrue（不明＝GAS優先）を返す。旧実装はここでfalse（＝GASと確定扱い）を返していた', () => {
    expect(SG.forceGasWhenBackendUnknown({})).toBe(true);
    expect(SG.forceGasWhenBackendUnknown({ backend: 'unknown-future-value' })).toBe(true);
    expect(SG.forceGasWhenBackendUnknown({ backend: null })).toBe(true);
  });

  it('backend===\'d1\'だがworkerUrlが無い（実質使えない）場合はfalseを返す（loadData側もD1を候補に入れないので無関係）', () => {
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

// ★5回目レビュー修正1・2・7（重大2件・両者一致／中1件・Codex）で createPreferGasTracker
// を作り直した。以下のテストは旧実装（時間だけで自動解除・タブ間非共有・並行refreshで
// 逆転しうる）のテストを置き換えたもの。詳細はsync-guard.js冒頭のコメント参照。
function makeFakeStorage() {
  const map = new Map();
  return {
    getItem: (k) => (map.has(k) ? map.get(k) : null),
    setItem: (k, v) => { map.set(k, String(v)); },
    removeItem: (k) => { map.delete(k); }
  };
}
function makeThrowingStorage() {
  return {
    getItem: () => { throw new Error('storage blocked'); },
    setItem: () => { throw new Error('storage blocked'); },
    removeItem: () => { throw new Error('storage blocked'); }
  };
}

describe('sync-guard.js: createPreferGasTracker（5回目レビュー修正1: 時間だけでは解除しない）', () => {
  it('mark()するとその直後はisActive()がtrue・status()が\'block\'になる', () => {
    const tracker = SG.createPreferGasTracker({ capMs: 15 * 60 * 1000, storage: makeFakeStorage() });
    const now = 1_000_000;
    tracker.mark(now);
    expect(tracker.isActive(now)).toBe(true);
    expect(tracker.status(now)).toBe('block');
  });

  it('★重大: capMs（上限）を過ぎても、確実成功を観測していなければisActive()はtrueのまま（旧実装の「時間で勝手に切れる」バグの直接的な再発防止テスト）', () => {
    const tracker = SG.createPreferGasTracker({ capMs: 5 * 60 * 1000, storage: makeFakeStorage() });
    const now = 1_000_000;
    tracker.mark(now);
    // 5分だけでなく、1時間・1日経ってもmark()もclear()も呼ばれていなければブロックしたまま。
    expect(tracker.isActive(now + 5 * 60 * 1000)).toBe(true);
    expect(tracker.isActive(now + 60 * 60 * 1000)).toBe(true);
    expect(tracker.isActive(now + 24 * 60 * 60 * 1000)).toBe(true);
  });

  it('capMsを過ぎるとstatus()が\'block\'から\'recheck\'に変わる（呼び出し側にhealth確認を促す合図）', () => {
    const tracker = SG.createPreferGasTracker({ capMs: 5 * 60 * 1000, storage: makeFakeStorage() });
    const now = 1_000_000;
    tracker.mark(now);
    expect(tracker.status(now + 5 * 60 * 1000 - 1)).toBe('block');
    expect(tracker.status(now + 5 * 60 * 1000)).toBe('recheck');
  });

  it('mark()前（一度も失敗していない）はstatus()が\'trust\'でisActive()がfalse', () => {
    const tracker = SG.createPreferGasTracker({ storage: makeFakeStorage() });
    expect(tracker.status()).toBe('trust');
    expect(tracker.isActive()).toBe(false);
  });

  it('clear()すると即座にisActive()がfalse・status()が\'trust\'になる（＝確実成功を検知したとき）', () => {
    const tracker = SG.createPreferGasTracker({ storage: makeFakeStorage() });
    const now = 1_000_000;
    const attemptStartedAt = tracker.beginAttempt(now);
    tracker.mark(now);
    expect(tracker.isActive(now)).toBe(true);
    tracker.clear(attemptStartedAt);
    expect(tracker.isActive(now)).toBe(false);
    expect(tracker.status(now)).toBe('trust');
  });

  it('resolveWithHealthEvidence: markedAtより後の確実な証拠があれば、capに達していなくても解除できる', () => {
    const tracker = SG.createPreferGasTracker({ capMs: 15 * 60 * 1000, storage: makeFakeStorage() });
    const markedAt = 1_000_000;
    tracker.mark(markedAt);
    expect(tracker.isActive(markedAt + 1000)).toBe(true);

    const resolved = tracker.resolveWithHealthEvidence(true, markedAt - 500); // markedAtより前の証拠→不十分
    expect(resolved).toBe(false);
    expect(tracker.isActive(markedAt + 1000)).toBe(true);

    const resolved2 = tracker.resolveWithHealthEvidence(true, markedAt + 1); // markedAtより後の証拠
    expect(resolved2).toBe(true);
    expect(tracker.isActive(markedAt + 2000)).toBe(false);
  });

  it('resolveWithHealthEvidence: evidenceOk:falseや証拠なしでは解除しない（確認できなければブロック継続＝安全側）', () => {
    const tracker = SG.createPreferGasTracker({ storage: makeFakeStorage() });
    const markedAt = 1_000_000;
    tracker.mark(markedAt);
    expect(tracker.resolveWithHealthEvidence(false, markedAt + 999999)).toBe(false);
    expect(tracker.resolveWithHealthEvidence(true, undefined)).toBe(false);
    expect(tracker.resolveWithHealthEvidence(true, NaN)).toBe(false);
    expect(tracker.isActive(markedAt + 1000)).toBe(true);
  });
});

describe('sync-guard.js: createPreferGasTracker（5回目レビュー修正2: localStorageへの永続化・タブ間共有）', () => {
  it('同じstorageを共有する2つのtrackerインスタンス（＝2つのタブを模す）で、一方のmark()がもう一方からも見える', () => {
    const sharedStorage = makeFakeStorage();
    const tabA = SG.createPreferGasTracker({ storage: sharedStorage });
    const tabB = SG.createPreferGasTracker({ storage: sharedStorage });
    const now = 1_000_000;

    expect(tabB.isActive(now)).toBe(false); // まだ何も無い
    tabA.mark(now); // タブAで書き込み直後の同期が確実成功しなかった
    expect(tabB.isActive(now)).toBe(true); // タブBが後から読んでも（新規タブ・リロード相当）ブロックが見える
  });

  it('一方のclear()ももう一方から見える（Cronが確実成功したことをどちらのタブでも認識できる）', () => {
    const sharedStorage = makeFakeStorage();
    const tabA = SG.createPreferGasTracker({ storage: sharedStorage });
    const tabB = SG.createPreferGasTracker({ storage: sharedStorage });
    const now = 1_000_000;
    tabA.mark(now);
    expect(tabB.isActive(now)).toBe(true);

    const attemptStartedAt = tabA.beginAttempt(now + 500);
    tabA.clear(attemptStartedAt);
    expect(tabB.isActive(now + 1000)).toBe(false);
  });

  it('別のstorageキー（別インスタンスの独立領域）を使えば状態は共有されない（回帰確認・意図しない相互汚染をしないことの確認）', () => {
    const storage = makeFakeStorage();
    const trackerA = SG.createPreferGasTracker({ storage, storageKey: 'a' });
    const trackerB = SG.createPreferGasTracker({ storage, storageKey: 'b' });
    const now = 1_000_000;
    trackerA.mark(now);
    expect(trackerA.isActive(now)).toBe(true);
    expect(trackerB.isActive(now)).toBe(false);
  });

  it('localStorageが使えない（setItem等が例外を投げる）環境でも、例外を投げずページ内メモリで動く', () => {
    const tracker = SG.createPreferGasTracker({ storage: makeThrowingStorage() });
    const now = 1_000_000;
    expect(() => tracker.mark(now)).not.toThrow();
    expect(tracker.isActive(now)).toBe(true); // メモリ内では正しく動作する
    const attemptStartedAt = tracker.beginAttempt(now);
    expect(() => tracker.clear(attemptStartedAt)).not.toThrow();
    expect(tracker.isActive(now)).toBe(false);
  });

  it('storageを明示的に渡さない場合はlocalStorageが無い環境（Node）でも例外にならず、ページ内メモリで動く', () => {
    // ブラウザではopts.storageを省略すると自動でwindow.localStorageを試みるが、
    // Node/vitest環境にはlocalStorageが無いため、自動的にメモリのみへフォールバックする。
    const tracker = SG.createPreferGasTracker();
    const now = 1_000_000;
    expect(() => tracker.mark(now)).not.toThrow();
    expect(tracker.isActive(now)).toBe(true);
  });
});

describe('sync-guard.js: createPreferGasTracker（5回目レビュー修正7: 並行refreshで古い成功が新しいmarkをclearしない）', () => {
  it('同期A（先に開始）が進行中に同期B（後に開始）がmarkし、その後Aが確実成功で遅れて返ってきても、Bのmarkは消えない', () => {
    const tracker = SG.createPreferGasTracker({ storage: makeFakeStorage() });

    const attemptA = tracker.beginAttempt(1000); // Aが先に開始
    const attemptB = tracker.beginAttempt(2000); // Bが後から開始（Aはまだ進行中）

    tracker.mark(3000); // Bが先に完了し、skipped（確実成功でない）でmark
    expect(tracker.isActive(4000)).toBe(true);

    const cleared = tracker.clear(attemptA); // 遅れて返ってきたAの「確実成功」
    expect(cleared).toBe(false); // Aより後に開始したBの結果（＝Aより後のmark）があるため反映されない
    expect(tracker.isActive(4000)).toBe(true); // ブロックされたまま（安全側）
  });

  it('後から開始した同期が確実成功で先に完了した場合は、先に開始した（まだ結果が返っていない）同期の存在に関わらず正常にclearできる', () => {
    const tracker = SG.createPreferGasTracker({ storage: makeFakeStorage() });
    tracker.beginAttempt(1000); // Aが先に開始（まだ進行中。結果はまだ来ない＝mark/clearどちらも未呼び出し）
    const attemptB = tracker.beginAttempt(2000); // Bが後から開始

    // Bが確実成功で先に完了。この時点で何もmarkされていないので、普通にclearできる。
    const clearedByB = tracker.clear(attemptB);
    expect(clearedByB).toBe(true);
    expect(tracker.isActive(3000)).toBe(false);
  });

  it('同期A（先に開始・確実成功）の結果が、Aの後に何のmarkも無いまま先に返ってくる通常のケースでは、普通にclearできる', () => {
    const tracker = SG.createPreferGasTracker({ storage: makeFakeStorage() });
    tracker.mark(500); // 直前の失敗
    const attemptA = tracker.beginAttempt(1000);
    expect(tracker.clear(attemptA)).toBe(true);
    expect(tracker.isActive(2000)).toBe(false);
  });

  it('mark()は常に反映される（安全側なので、古い試行からのmarkが新しい試行のclearを妨げても実害は「余分にブロックする」だけ）', () => {
    const tracker = SG.createPreferGasTracker({ storage: makeFakeStorage() });
    const attemptA = tracker.beginAttempt(1000);
    tracker.clear(attemptA); // 何も無い状態でのclearは常に成立
    expect(tracker.isActive(2000)).toBe(false);

    tracker.mark(3000); // 別の（古いかもしれない）試行がmark
    expect(tracker.isActive(4000)).toBe(true); // 常に反映される
  });
});

describe('sync-guard.js: createPreferGasTracker（loadDataのeffectiveForceGas算出との結線）', () => {
  it('「preferGasがブロック中の間はforceGas引数無しの呼び出しもGASを使う」を再現する（loadDataのeffectiveForceGas算出と同じ式）', () => {
    // index.html/admin.htmlのloadData()は
    //   const effectiveForceGas = !!forceGas || preferGasTracker.isActive();
    // という式でこのトラッカーを参照する。「更新」ボタン・プル更新・switchCompany等は
    // forceGas引数を渡さない（undefined）が、trackerがブロック中の間は
    // effectiveForceGasがtrueになる＝D1を候補から外しGASを使う、という点をここで確認する。
    const tracker = SG.createPreferGasTracker({ storage: makeFakeStorage() });
    const now = 2_000_000;
    tracker.mark(now); // 直前の同期が確実成功でなかった

    const forceGasArgFromButton = undefined; // 「更新」ボタン等はforceGasを渡さない
    const effectiveForceGas = !!forceGasArgFromButton || tracker.isActive(now + 1000);
    expect(effectiveForceGas).toBe(true);

    const attemptStartedAt = tracker.beginAttempt(now + 1500);
    tracker.clear(attemptStartedAt); // 次のCron巡回で確実成功した
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
