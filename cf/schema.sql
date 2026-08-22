-- 日報データ。スプレッドシートの19列をそのまま持つ。
-- id はスプレッドシートの「ID」列（同じ予定グループで同一値）。
-- 行の一意キーは (id, sagyoubi, shimei)。同じIDで日と人が異なる行が並ぶため。
CREATE TABLE IF NOT EXISTS nippo (
  id            TEXT NOT NULL,
  touroku       TEXT,           -- 登録日時
  sagyoubi      TEXT NOT NULL,  -- 作業日 YYYY-MM-DD
  motoukr       TEXT,           -- 元請名
  genba         TEXT,           -- 現場名
  shimei        TEXT NOT NULL,  -- 氏名
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
  sharyou       TEXT,           -- 車両
  PRIMARY KEY (id, sagyoubi, shimei)
);
CREATE INDEX IF NOT EXISTS idx_nippo_sagyoubi ON nippo(sagyoubi);
CREATE INDEX IF NOT EXISTS idx_nippo_kaisha   ON nippo(kaisha);

-- 職人マスタ。★給料情報は意図的に持たない（D1へ持ち込まない）。
CREATE TABLE IF NOT EXISTS members (
  name     TEXT NOT NULL,
  company  TEXT NOT NULL,
  division TEXT,
  PRIMARY KEY (name, company)
);

CREATE TABLE IF NOT EXISTS genba (
  name    TEXT NOT NULL,
  company TEXT,
  PRIMARY KEY (name, company)
);

CREATE TABLE IF NOT EXISTS jobsites (
  genba          TEXT NOT NULL,
  loc            TEXT NOT NULL,
  jobNo          TEXT,
  completed      INTEGER DEFAULT 0,
  billingMethod  TEXT,
  PRIMARY KEY (genba, loc)
);

-- 取り込みの記録。最後にいつ・何行取り込んだかを残す（障害調査用）。
CREATE TABLE IF NOT EXISTS sync_log (
  at       TEXT PRIMARY KEY,
  rows     INTEGER,
  ok       INTEGER,
  message  TEXT
);
