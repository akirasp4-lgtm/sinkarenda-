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

CREATE TABLE IF NOT EXISTS snapshot (
  id       INTEGER PRIMARY KEY CHECK (id = 1),  -- 常に1行だけ（CHECKで強制）
  payload  TEXT NOT NULL,     -- GASのcompact応答をJSON文字列にしたもの（単価は除去済み。給料情報はD1へ持ち込まない）
  hash     TEXT NOT NULL,     -- 中身が変わったかの判定用（SHA-256）。夜間・休日の無変化時は書き込みをスキップする
  rows     INTEGER NOT NULL,  -- 日報(rows)の行数（健全性確認・急減検知用）
  bytes    INTEGER NOT NULL,  -- payloadのUTF-8バイト数（サイズガード用。1行上限2,000,000バイトに対する余裕の確認用）
  at       TEXT NOT NULL      -- 書き込み時刻（ISO8601）
);

-- 同時実行の抑止（修正2）。/api/sync が並行して複数走るのを防ぐための簡易ロック。
-- 常に1行だけで、locked_at が直近（cf/src/sync.jsのLOCK_STALE_MS。既定90秒以内）
-- ならロック中とみなしスキップする。それより古ければ「前回が異常終了して
-- 解放されなかった」とみなして上書きする（永久に固まらないための安全弁）。
CREATE TABLE IF NOT EXISTS sync_lock (
  id        INTEGER PRIMARY KEY CHECK (id = 1),
  locked_at TEXT   -- ロック取得時刻のepoch文字列。NULLなら未ロック
);

-- 取り込みの記録。最後にいつ・何行・成功したか失敗したかを残す（障害調査用）。
-- ★スナップショット方式でも sync_log は維持する：読み取り側は使わなくなったが
-- （snapshotが存在すること自体が「直近の成功状態」を保証するため）、
-- 「いつ・なぜ失敗したか」の履歴はCronの障害調査に必要。
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
