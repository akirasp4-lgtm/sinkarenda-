-- 日報データ。スプレッドシートの19列をそのまま持つ。
-- id はスプレッドシートの「ID」列（同じ予定グループで同一値）。
-- ★2026-08-24 設計変更：スプレッドシートは「行を一意に識別できる列」を
-- 持たない、ただの記録の羅列。氏名が空の正当な行（車検期限リマインダー等、
-- 人が出る予定ではなく車両の期限を置いてあるだけの行）が実在するため、
-- (id, sagyoubi, shimei) を主キーにして氏名やIDが空の行を捨てると、
-- GASが返す行がD1側で欠落する（本番データ突き合わせで14行欠落を確認）。
-- D1はGAS応答の忠実な写しとし、複合主キー・NOT NULL・重複排除はやめ、
-- 連番(seq)だけを主キーにする。並び順は seq の昇順＝取り込み順で保つ。
CREATE TABLE IF NOT EXISTS nippo (
  seq           INTEGER PRIMARY KEY AUTOINCREMENT,
  id            TEXT,           -- スプレッドシートの「ID」列
  touroku       TEXT,           -- 登録日時
  sagyoubi      TEXT,           -- 作業日 YYYY-MM-DD
  motoukr       TEXT,           -- 元請名
  genba         TEXT,           -- 現場名
  shimei        TEXT,           -- 氏名（空が正当なケースあり＝車検期限リマインダー等）
  yakuwari      TEXT,           -- 役割
  shukkin       TEXT,           -- 出勤 HH:MM
  taikin        TEXT,           -- 退勤 HH:MM
  kosu          REAL DEFAULT 0, -- 人工
  memo          TEXT,
  yakin         TEXT,           -- '夜勤'/'休み'/'予定'/'倉庫'/''
  kaisha        TEXT,           -- 会社
  koushinsha    TEXT,           -- 更新者
  iro           TEXT,           -- 色
  jigyoubu      TEXT,           -- 事業部
  kouban        TEXT,           -- 工番
  sagyou_kubun  TEXT,           -- 作業区分
  sharyou       TEXT            -- 車両
);
CREATE INDEX IF NOT EXISTS idx_nippo_sagyoubi ON nippo(sagyoubi);
CREATE INDEX IF NOT EXISTS idx_nippo_kaisha   ON nippo(kaisha);

-- 職人マスタ。★給料情報は意図的に持たない（D1へ持ち込まない）。
-- 同名・同会社の重複行（元データの重複）もそのまま持つ。一意制約はしない。
CREATE TABLE IF NOT EXISTS members (
  seq      INTEGER PRIMARY KEY AUTOINCREMENT,
  name     TEXT,
  company  TEXT,
  division TEXT
);

CREATE TABLE IF NOT EXISTS genba (
  seq     INTEGER PRIMARY KEY AUTOINCREMENT,
  name    TEXT,
  company TEXT
);

CREATE TABLE IF NOT EXISTS jobsites (
  seq            INTEGER PRIMARY KEY AUTOINCREMENT,
  genba          TEXT,
  loc            TEXT,
  jobNo          TEXT,
  completed      INTEGER DEFAULT 0,
  billingMethod  TEXT
);

-- 取り込みの記録。最後にいつ・何行取り込んだかを残す（障害調査用）。
CREATE TABLE IF NOT EXISTS sync_log (
  at       TEXT PRIMARY KEY,
  rows     INTEGER,
  ok       INTEGER,
  message  TEXT
);
