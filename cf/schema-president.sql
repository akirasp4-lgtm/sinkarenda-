-- 社長予定をD1へ載せるための追加スキーマ（2026-08-26）
--
-- ★このファイルは cf/schema.sql とは別に用意した。理由:
--   cf/schema.sql は先頭で snapshot / sync_lock / sync_log を DROP する破壊的な作りで、
--   本番切替後（backend.json が "d1" の今）に再適用すると社員用のD1が一瞬空になる。
--   社長予定のテーブルを足すためだけにその危険を冒す必要は無いので、
--   「追加するものだけ」をこのファイルに分けた。DROPは1つも書かない。
--
-- 適用:
--   npx wrangler d1 execute yotei --remote --file=schema-president.sql
--   （何度流しても安全＝IF NOT EXISTS のみ）

-- ★なぜ社員用の snapshot に相乗りしないか:
--   snapshot は CHECK (id = 1) で1行しか持てない設計のため、そもそも相乗りできない。
CREATE TABLE IF NOT EXISTS pres_snapshot (
  id               INTEGER PRIMARY KEY CHECK (id = 1),  -- 常に1行だけ
  payload          TEXT NOT NULL,     -- GASの pres_list 応答の rows をJSON文字列にしたもの
  hash             TEXT NOT NULL,     -- 中身が変わったかの判定用（SHA-256）
  rows             INTEGER NOT NULL,  -- 予定の件数（急減・全消え検知用）
  bytes            INTEGER NOT NULL,  -- payloadのUTF-8バイト数（サイズガード用）
  fetch_started_at INTEGER NOT NULL,  -- GASへの取得を開始した時刻のepoch ms。
                                       -- これより古い取得結果では上書きしない、という
                                       -- WHERE条件の比較対象（世代の逆転を防ぐ本体）。
                                       -- 社員用と違いロックテーブルは作らない。社員用の
                                       -- cf/schema.sql 自身が「ロックはbest-effortに過ぎず、
                                       -- 正しさの最終防衛はこのWHERE条件が担う」と明記して
                                       -- おり、社長予定は件数が2桁と小さく取得も速いため、
                                       -- 守りの本体であるこの条件だけを採用した。
  at               TEXT NOT NULL      -- 書き込み時刻（ISO8601）
);

-- ★なぜ社員用の sync_log に相乗りしてはいけないか（最重要）:
--   社員用の鮮度判定 cf/src/read.js の getLastSuccessAt は
--     SELECT at FROM sync_log WHERE ok = 1 ORDER BY at DESC LIMIT 1
--   で「直近の成功」を見ている。ここに社長予定の同期結果を混ぜると、
--   **社員用の同期が失敗し続けていても、社長予定の同期が成功しただけで
--   「社員用のデータは新しい」と誤判定**され、古いデータを正常として返してしまう。
--   さらに cf/src/index.js の isSyncRateLimited が直近1分間の sync_log 行数で
--   レート制限を判定しているため、行数も歪む。
--   → 社長予定は必ずこの専用テーブルにだけ書く。sync_log には1行も書かない。
CREATE TABLE IF NOT EXISTS pres_sync_log (
  at            TEXT PRIMARY KEY,
  rows          INTEGER,
  ok            INTEGER,
  message       TEXT,
  payload_hash  TEXT  -- 今回取得した内容のSHA-256。取得自体が失敗したときはNULL。
                       -- 急減ガードの自己回復（同じ内容の拒否が30分続いているか）に使う。
);
