-- ★2026-08-24 最終レビュー（Fable 5 / Codex 両者）で「切り替え不可」の判定を受け、
-- D1の持ち方を「行ごとのテーブル」から「スナップショット1行」へ全面的に変更した。
--
-- 変更前の設計（nippo/members/genba/jobsitesの4テーブル・全件DELETE+500文ずつINSERT）には
-- 独立に発見された2つの重大欠陥があった：
--   1) 原子性: batch()の原子性はチャンク単位のみで、その隙間の読み取りが中途半端な
--      D1をstatus:'ok'で返してしまう（Codexが実装を使って再現済み）。
--   2) 費用: 1回の同期=5,765行の書き込み×288回/日=166万行/日。D1無料枠10万行/日の
--      16.6倍（実測: 稼働3.8時間で667,466行）。
--
-- 解決策: GAS応答（compact形式）をJSON文字列にして snapshot テーブルの1行へまるごと
-- 格納する。書き込みは INSERT OR REPLACE の1文・1行のみになるため、
--   - 原子的（単一文なので中途半端な状態が外部から見える瞬間が原理的に無い）
--   - 1回の同期が1行の書き込みになる（288回/日でも288行/日。無料枠の0.3%）
-- の両方が同時に解決する。読み取り側（cf/src/read.js）は1行SELECT→JSON.parseし、
-- 会社での絞り込みはWorker内(JS)で行う。
--
-- ★2026-08-24 再レビュー（Fable 5 / Codex 両者）で、上記スナップショット方式のままでも
-- なお「切り替え不可」の判定を受け、さらに以下2点を修正した（cf/src/read.js・sync.js参照）：
--   3) 鮮度: snapshotが「存在するだけ」で正常返却していた（同期が失敗し続けても、最後に
--      成功した内容を永久に「正常」として返し続ける）。→ sync_logの直近成功時刻を見て、
--      古すぎればreadSchedule自身がstatus:'error'を返す（追加のテーブルは不要。sync_logの
--      既存の行を使う）。
--   4) 世代の逆転: sync_lockの取得が「SELECTで確認→INSERTで取得」の2文に分かれており、
--      2つの同期が両方取得成功しうる。先に取得した古いGAS応答が遅れて完了すると、新しい
--      スナップショットを古い内容で上書きできてしまう（Codexが再現）。
--      → sync_lockの取得を「INSERT ... ON CONFLICT ... WHERE」の単一文にし、D1の
--      meta.changesで取得可否を判定する（2文の間に割り込む余地を無くす）。
--      → さらに本質的な防御として、snapshotに「GASへの取得を開始した時刻」
--      (fetch_started_at) を持たせ、書き込みそのものも「保存済みより古い取得時刻の
--      結果では上書きしない」という条件付き単一文にする。ロックは「無駄な二重実行を
--      減らす」ためのbest-effortに過ぎず、正しさの最終防衛はこのWHERE条件が担う。
--   5) マスタの全消え: 半減チェックが日報(rows)にしかなく、members/genbaMaster/jobsites
--      は空配列でも受け入れていた。→ snapshotに各マスタの件数
--      (members_count/genba_count/jobsites_count) を持たせ、日報と同じ半減チェックを
--      3つのマスタにも適用する（cf/src/sync.js）。

-- ★上記3)〜5)で列構成が変わるため、snapshot/sync_lockは一度DROPしてから作り直す。
-- D1はあくまでGASの派生コピー（sync.js冒頭のコメント参照）で、まだ本番切替前
-- （backend.jsonはgasのまま）のため、安全に作り直せる。壊れても次の同期で全件戻る。
DROP TABLE IF EXISTS snapshot;
DROP TABLE IF EXISTS sync_lock;

CREATE TABLE IF NOT EXISTS snapshot (
  id               INTEGER PRIMARY KEY CHECK (id = 1),  -- 常に1行だけ（CHECKで強制）
  payload          TEXT NOT NULL,     -- GASのcompact応答をJSON文字列にしたもの（単価は除去済み。給料情報はD1へ持ち込まない）
  hash             TEXT NOT NULL,     -- 中身が変わったかの判定用（SHA-256）。夜間・休日の無変化時は書き込みをスキップする
  rows             INTEGER NOT NULL,  -- 日報(rows)の行数（健全性確認・急減検知用）
  members_count    INTEGER NOT NULL,  -- 職人マスタの件数（急減・全消え検知用。修正3）
  genba_count      INTEGER NOT NULL,  -- 元請マスタの件数（同上）
  jobsites_count   INTEGER NOT NULL,  -- 現場マスタの件数（同上）
  bytes            INTEGER NOT NULL,  -- payloadのUTF-8バイト数（サイズガード用。1行上限2,000,000バイトに対する余裕の確認用）
  fetch_started_at INTEGER NOT NULL,  -- GASへの取得(fetch)を開始した時刻のepoch ms。修正2:
                                       -- これより古い取得時刻の結果では上書きしない、という
                                       -- WHERE条件の比較対象（世代の逆転を防ぐ本体）
  at               TEXT NOT NULL      -- 書き込み時刻（ISO8601）
);

-- 同時実行の抑止（修正2）。/api/sync が並行して複数走るのを防ぐための簡易ロック。
-- 常に1行だけで、locked_at が直近（cf/src/sync.jsのLOCK_STALE_MS。既定90秒以内）
-- ならロック中とみなしスキップする。それより古ければ「前回が異常終了して
-- 解放されなかった」とみなして上書きする（永久に固まらないための安全弁）。
-- ★再レビュー修正: 取得(SELECT)と確保(INSERT)を1文にまとめた（cf/src/sync.jsのtryAcquireLock）。
-- ただしこのロックは「無駄な二重実行を減らす」ためのbest-effortであり、正しさの最終防衛は
-- 上のsnapshot.fetch_started_atのWHERE条件が担う。
CREATE TABLE IF NOT EXISTS sync_lock (
  id        INTEGER PRIMARY KEY CHECK (id = 1),
  locked_at TEXT   -- ロック取得時刻のepoch文字列。NULLなら未ロック
);

-- 取り込みの記録。最後にいつ・何行・成功したか失敗したかを残す（障害調査用）。
-- ★スナップショット方式でも sync_log は維持する：
--   - 障害調査用の履歴として（いつ・なぜ失敗したか）
--   - ★再レビュー修正: 読み取り側（read.js）の鮮度ガードが「直近の成功時刻」を
--     ここから読む。ハッシュ一致で書き込みをスキップした場合も ok=1 で記録するため
--     （sync.jsのsyncAll）、「変更が無いだけ」を「古い」と誤判定しない。
--   - ★再レビュー修正: 急減ガードの自己回復（修正7）が、直近の拒否が何回連続したかを
--     ここから数える（cf/src/sync.jsのrecentConsecutiveShrinkRejections）。
CREATE TABLE IF NOT EXISTS sync_log (
  at       TEXT PRIMARY KEY,
  rows     INTEGER,
  ok       INTEGER,
  message  TEXT
);

-- ★旧設計（行ごとのテーブル）は廃止。もう使わない。
DROP TABLE IF EXISTS nippo;
DROP TABLE IF EXISTS members;
DROP TABLE IF EXISTS genba;
DROP TABLE IF EXISTS jobsites;
