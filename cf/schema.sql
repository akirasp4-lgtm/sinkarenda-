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
--
-- ★2026-08-24 3回目レビュー（Fable 5 / Codex）で、なお以下2件（重大・高）が残っている
-- 判定を受け、sync_logに列を1つ追加した（cf/src/sync.js参照）：
--   6) 急減ガードの自己回復が「回数」だけで判定しており、Codexが「毎回まったく別の
--      欠損を連発しても3回連続で自動受入されてしまう」ことを再現した。→ 拒否した
--      取得内容のハッシュ(payload_hash)をsync_logに記録し、「同一ハッシュの拒否が
--      最初の拒否から30分以上続いている」ときだけ自己回復するように変更した。
--
-- ついでに /api/sync のレート制限（cf/src/index.js、修正5）もsync_logの直近1分間の
-- 行数を数えて判定する形にした（新しいテーブルを増やさず、既存の「実際にGAS/D1へ
-- 負荷をかけた回数」の記録をそのまま使う）。

-- ★上記6)でsync_logの列構成も変わるため、snapshot/sync_lock/sync_logをすべて
-- 一度DROPしてから作り直す。
--
-- ★注意（この一括DROPは「まだ本番切替前だから安全」という前提に依存している）:
-- backend.jsonが"gas"のまま（＝画面がまだD1を読みに行っていない）今のうちは、
-- 一度空にしても実害が無い（D1はあくまでGASの派生コピー。次のCronの成功で全件戻る）。
-- しかし本番切替後（backend.jsonが"d1"になった後）にこのファイルを再適用すると、
-- その瞬間 snapshot / sync_log が空になる。read.jsは「まだ取り込みが行われていません」
-- を返し、画面側は自動でGASへフォールバックするため利用者に実害は無いが、次のCron
-- （最大5分後）が成功するまでD1経由の読み取りが一時的に使えなくなる。列追加などで
-- 切替後にこのファイルを再適用する必要が生じたときは、深夜・早朝などアクセスの
-- 少ない時間帯に行うこと（毎回DROPする設計そのものを変えない限り、この注意点は
-- 消えない）。
DROP TABLE IF EXISTS snapshot;
DROP TABLE IF EXISTS sync_lock;
DROP TABLE IF EXISTS sync_log;

CREATE TABLE IF NOT EXISTS snapshot (
  id               INTEGER PRIMARY KEY CHECK (id = 1),  -- 常に1行だけ（CHECKで強制）
  payload          TEXT NOT NULL,     -- GASのcompact応答をJSON文字列にしたもの（単価は除去済み。給料情報はD1へ持ち込まない）
  hash             TEXT NOT NULL,     -- 中身が変わったかの判定用（SHA-256）。夜間・休日の無変化時は書き込みをスキップする
  raw_hash         TEXT,              -- GASの生の応答そのもののSHA-256。前回と同じなら
                                       -- JSONの解析も組み直しも丸ごと省く（CronのCPU上限対策・
                                       -- 2026-08-31の本番障害）。古い行にはNULLが入るので
                                       -- 「NULLなら省かない」＝安全側に倒れる
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
--   - ★3回目レビュー修正3: 急減ガードの自己回復が、直近の拒否が「同じ内容
--     (payload_hash)のまま」何分続いているかをここから遡って数える
--     （cf/src/sync.jsのsameHashShrinkRejectStreak）。回数だけでなく内容の一致も
--     見るようにしたのは、Codexが「毎回別の内容で拒否させても回数だけで自動受入
--     されてしまう」ことを再現したため。
--   - ★3回目レビュー修正5: /api/syncのレート制限（cf/src/index.js）が、直近1分間の
--     行数を「実際にGAS/D1へ負荷をかけた回数」の実測値として使う。
-- ★30日より古い行はCronのたびに掃除する（cf/src/sync.jsのcleanupSyncLog。修正8）。
-- 無限に増え続けるのを防ぐ。
CREATE TABLE IF NOT EXISTS sync_log (
  at            TEXT PRIMARY KEY,
  rows          INTEGER,
  ok            INTEGER,
  message       TEXT,
  payload_hash  TEXT  -- 今回取得した内容のSHA-256。取得自体が失敗し内容が無いときはNULL。
                       -- 急減ガードの自己回復（同一内容の拒否が続いているか）の判定に使う。
);

-- ★旧設計（行ごとのテーブル）は廃止。もう使わない。
DROP TABLE IF EXISTS nippo;
DROP TABLE IF EXISTS members;
DROP TABLE IF EXISTS genba;
DROP TABLE IF EXISTS jobsites;

-- AI（要件5の候補者の順位付け）の呼び出し記録。
-- ★目的は「1日に何回呼んだか」を数えて課金の上限を守ること。
--   数えられないときは呼ばない側に倒す（cf/src/suggest.js の overDailyLimit）。
CREATE TABLE IF NOT EXISTS ai_log (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  at TEXT NOT NULL,     -- ISO8601
  ok INTEGER NOT NULL   -- 1=成功 0=失敗
);
CREATE INDEX IF NOT EXISTS idx_ai_log_at ON ai_log(at);
