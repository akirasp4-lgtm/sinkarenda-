# 予定管理アプリ Cloudflare移行 フェーズ1（読み取り経路）実装計画

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** アプリを開くときの読み取りを Cloudflare Workers + D1 に移し、9.7秒→0.3秒台にする。書き込みは一切触らない。

**Architecture:** D1 は**スプレッドシートの読み取り専用コピー**として作る。Worker が既存のGAS `doGet` を定期的に呼んで D1 へ取り込み、アプリは D1 から読む。書き込みは今までどおりGAS→スプレッドシート。**スプレッドシートが唯一の正**であり続けるので、D1 が壊れても消えるデータは無い。切り替えは設定ファイル1行、失敗時は自動でGASへ落ちる。

**Tech Stack:** Cloudflare Workers（JavaScript / ESM）, Cloudflare D1（SQLite）, wrangler CLI, 既存の GitHub Pages（index.html / admin.html）

**Spec:** `引き継ぎ.md` の「6. 既知の課題・残課題 → ② GASから離れるか」（Cloudflare採用の理由・料金・無停止切替の設計要件）

## Global Constraints

- **利用者に気づかせない。** アプリURL・アイコン・合言葉・localStorageの設定はすべて据え置き。利用者が気づくのは「速くなった」ことだけ。
- **切り替えと切り戻しは設定ファイル1行**（`backend.json`）。数十秒で往復できること。
- **元に戻せない一方通行の作業を1つも作らない。** このフェーズでは日報データを一切書き換えない。
- **スプレッドシートが唯一の正。** D1 は派生コピー。D1 を消しても再取り込みで完全復元できること。
- **GAS側のコードは1行も変更しない。** 38アクションのうち34（集計生成・Excel出力・マスタ管理・社長カレンダー・**ラーテルLINEボット連携** `vehicle_res_*` / `warehouse_today`）はGASのまま。`ラインボット/` は本番稼働中につき触らない（CLAUDE.md §2）。
- **金額・単価をD1に入れない。** 職人マスタの `rate`（単価＝給料情報）はD1へ取り込まない。2026-06-11の給料漏れ対策の線引きを維持する。
- **費用は0円。** Cloudflare 無料枠内（Workers 10万req/日・D1 500万行読/日・書10万行/日・5GB）。超過時は課金ではなく停止するため、使用量ログを残すこと。
- 本番デプロイ（wrangler deploy / GitHub Pages push）は**実行前に利用者へ確認**する。2026-08-21にGASの誤デプロイで本番を20分停止させた前例がある。

---

## ファイル構成

| ファイル | 役割 |
|---|---|
| `cf/wrangler.toml` | Worker と D1 の設定（DB名・バインディング・Cron） |
| `cf/schema.sql` | D1 のテーブル定義（日報 + マスタ3種） |
| `cf/src/index.js` | Worker 本体。ルーティングのみ |
| `cf/src/sync.js` | GAS `doGet` から取り込んで D1 へ upsert する |
| `cf/src/read.js` | `/api/schedule` の組み立て（doGetと同じJSON形状を返す） |
| `cf/test/sync.test.js` | 取り込みの単体テスト |
| `cf/test/read.test.js` | 応答形状の単体テスト |
| `backend.json` | 切り替えスイッチ（GitHub Pages 直下） |
| `index.html` / `admin.html` | 読み取り先の切り替えと自動フォールバック |

---

## 前提の確認（着手前に必ず読む）

現行 `doGet` が返す形（実測・2026-08-22）:

```json
{"status":"ok","compact":1,
 "headers":["登録日時","作業日","元請名","現場名","氏名","役割","出勤","退勤","人工","メモ","夜勤","会社","ID","更新者","色","事業部","工番","作業区分","車両"],
 "rows":[[...19個の値...], ...],
 "members":[{"name":"","company":"","division":"","rate":0}],
 "genbaMaster":[{"name":"","company":""}],
 "jobsites":[{"genba":"","loc":"","jobNo":"","completed":false,"billingMethod":"応援"}]}
```

- `?compact=1` を付けないと `rows` はキー付きオブジェクトの配列になる（後方互換のため両方残っている）。**Worker は必ず `compact=1` で取る。**
- `?company=` を空にすると全社。
- 実測: 全社で約2,600行 / 約1.0MB / 3.9〜56秒、**5回に1回 HTTP404 を返す**（GAS側の不安定さ。§3.0b）。取り込みは必ずリトライすること。

---

### Task 1: D1 のテーブルを作り、スキーマを固める

**Files:**
- Create: `cf/wrangler.toml`
- Create: `cf/schema.sql`
- Create: `cf/package.json`
- Test: `cf/test/schema.test.js`

**Interfaces:**
- Consumes: なし（最初のタスク）
- Produces: D1テーブル `nippo` / `members` / `genba` / `jobsites`。列名は後続タスクがそのまま使う。

- [ ] **Step 1: 失敗するテストを書く**

`cf/test/schema.test.js`:

```javascript
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';

describe('schema.sql', () => {
  const sql = readFileSync(new URL('../schema.sql', import.meta.url), 'utf8');

  it('日報テーブルに19列そろっている', () => {
    const cols = ['id','touroku','sagyoubi','motoukr','genba','shimei','yakuwari',
                  'shukkin','taikin','kosu','memo','yakin','kaisha','koushinsha',
                  'iro','jigyoubu','kouban','sagyou_kubun','sharyou'];
    for (const c of cols) expect(sql).toContain(c);
  });

  it('作業日に索引がある（期間検索を速くするため）', () => {
    expect(sql).toMatch(/CREATE INDEX .*nippo.*sagyoubi/i);
  });

  it('単価(rate)は保存しない（給料情報をD1へ持ち込まない）', () => {
    expect(sql).not.toMatch(/\brate\b/i);
  });
});
```

- [ ] **Step 2: テストが落ちることを確認**

Run: `cd cf && npx vitest run test/schema.test.js`
Expected: FAIL（`schema.sql` が存在しない）

- [ ] **Step 3: schema.sql を書く**

```sql
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

-- 職人マスタ。★単価(rate)は意図的に持たない（給料情報をD1へ入れない）。
CREATE TABLE IF NOT EXISTS members (
  name     TEXT PRIMARY KEY,
  company  TEXT,
  division TEXT
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
```

- [ ] **Step 4: wrangler.toml と package.json を書く**

`cf/wrangler.toml`:

```toml
name = "yotei-api"
main = "src/index.js"
compatibility_date = "2026-08-01"

[[d1_databases]]
binding = "DB"
database_name = "yotei"
database_id = "PLACEHOLDER_実際のIDはwrangler d1 createの出力で置き換える"

# 5分ごとにスプレッドシートから取り込む
[triggers]
crons = ["*/5 * * * *"]

[vars]
GAS_URL = "https://script.google.com/macros/s/AKfycbxp2eUcpIjCj0ZWyAPPD9m3egJrKdWmXRK2AVnFrmBm4iO1QHCk-FZEH5LFFv7OloqcjQ/exec"
```

`cf/package.json`:

```json
{
  "name": "yotei-api",
  "private": true,
  "type": "module",
  "scripts": {
    "test": "vitest run",
    "dev": "wrangler dev",
    "deploy": "wrangler deploy"
  },
  "devDependencies": {
    "vitest": "^2.0.0",
    "wrangler": "^3.0.0"
  }
}
```

- [ ] **Step 5: テストが通ることを確認**

Run: `cd cf && npm install && npx vitest run test/schema.test.js`
Expected: PASS（3件）

- [ ] **Step 6: コミット**

```bash
git add cf/wrangler.toml cf/schema.sql cf/package.json cf/test/schema.test.js
git commit -m "feat(cf): D1のスキーマとWorkerの雛形を追加（単価は持たない）"
```

---

### Task 2: GASから取り込んでD1へ入れる

**Files:**
- Create: `cf/src/sync.js`
- Test: `cf/test/sync.test.js`

**Interfaces:**
- Consumes: Task 1 のテーブル定義
- Produces:
  - `parseGasPayload(json) -> {nippo: Row[], members, genba, jobsites}` — GASのcompact応答をD1の行形へ変換する純関数
  - `fetchWithRetry(url, tries) -> Promise<object>` — GASの404対策つき取得
  - `syncAll(env) -> Promise<{ok:boolean, rows:number, message:string}>`

- [ ] **Step 1: 失敗するテストを書く**

`cf/test/sync.test.js`:

```javascript
import { describe, it, expect, vi } from 'vitest';
import { parseGasPayload, fetchWithRetry } from '../src/sync.js';

const HEADERS = ['登録日時','作業日','元請名','現場名','氏名','役割','出勤','退勤','人工',
                 'メモ','夜勤','会社','ID','更新者','色','事業部','工番','作業区分','車両'];

describe('parseGasPayload', () => {
  it('compact形式の1行をD1の列名へ移し替える', () => {
    const json = {
      status: 'ok', compact: 1, headers: HEADERS,
      rows: [['2026-05-01T04:23:04.000Z','2026-05-02','NGS','大阪','川端（達）','代表',
              '09:00','18:00',1,'','','グローライズ','abc-1','森','#1D9E75','ICT',
              'INF-26-041','現場作業','']],
      members: [], genbaMaster: [], jobsites: []
    };
    const out = parseGasPayload(json);
    expect(out.nippo).toHaveLength(1);
    expect(out.nippo[0]).toMatchObject({
      id: 'abc-1', sagyoubi: '2026-05-02', motoukr: 'NGS', genba: '大阪',
      shimei: '川端（達）', kosu: 1, kaisha: 'グローライズ', kouban: 'INF-26-041'
    });
  });

  it('職人マスタから単価(rate)を落とす', () => {
    const json = { status:'ok', compact:1, headers:HEADERS, rows:[],
      members:[{name:'森',company:'GRHD',division:'ICT',rate:18000}],
      genbaMaster:[], jobsites:[] };
    const out = parseGasPayload(json);
    expect(out.members[0]).toEqual({name:'森',company:'GRHD',division:'ICT'});
    expect(JSON.stringify(out.members)).not.toContain('18000');
  });

  it('compactでない応答は受け付けない（形が変わると壊れるため明示的に落とす）', () => {
    expect(() => parseGasPayload({status:'ok', rows:[{'ID':'x'}]}))
      .toThrow(/compact/);
  });

  it('IDが空の行は捨てる（主キーにできないため）', () => {
    const row = new Array(19).fill('');
    row[1] = '2026-05-02'; row[4] = '森';
    const json = { status:'ok', compact:1, headers:HEADERS, rows:[row],
                   members:[], genbaMaster:[], jobsites:[] };
    expect(parseGasPayload(json).nippo).toHaveLength(0);
  });
});

describe('fetchWithRetry', () => {
  it('1回目が404でも2回目で成功すれば結果を返す（GASは5回に1回404を返す）', async () => {
    const calls = [];
    global.fetch = vi.fn(async (u) => {
      calls.push(u);
      if (calls.length === 1) return { ok:false, status:404, text: async () => '<html>' };
      return { ok:true, status:200, json: async () => ({status:'ok'}) };
    });
    const out = await fetchWithRetry('https://example.test/', 3);
    expect(out).toEqual({status:'ok'});
    expect(calls).toHaveLength(2);
  });

  it('回数を使い切ったら投げる', async () => {
    global.fetch = vi.fn(async () => ({ ok:false, status:404, text: async () => '<html>' }));
    await expect(fetchWithRetry('https://example.test/', 2)).rejects.toThrow(/404/);
  });
});
```

- [ ] **Step 2: テストが落ちることを確認**

Run: `cd cf && npx vitest run test/sync.test.js`
Expected: FAIL（`src/sync.js` が無い）

- [ ] **Step 3: sync.js を実装する**

```javascript
// GASの doGet(compact=1) からスプレッドシートの内容を取り込み、D1へ入れる。
// ★ここは「読むだけ」。スプレッドシートには何も書かない。
// ★D1はあくまで派生コピー。壊れても全件取り込み直せば完全に戻る。

const H = ['登録日時','作業日','元請名','現場名','氏名','役割','出勤','退勤','人工',
           'メモ','夜勤','会社','ID','更新者','色','事業部','工番','作業区分','車両'];

const COL = ['touroku','sagyoubi','motoukr','genba','shimei','yakuwari','shukkin','taikin',
             'kosu','memo','yakin','kaisha','id','koushinsha','iro','jigyoubu','kouban',
             'sagyou_kubun','sharyou'];

export function parseGasPayload(json) {
  if (!json || json.compact !== 1 || !Array.isArray(json.headers)) {
    throw new Error('compact形式の応答ではありません（?compact=1 を付けて取得すること）');
  }
  // ヘッダの並びがGAS側で変わっても壊れないよう、名前で位置を引く
  const pos = {};
  json.headers.forEach((h, i) => { pos[h] = i; });

  const nippo = [];
  for (const row of (json.rows || [])) {
    const rec = {};
    H.forEach((h, i) => { rec[COL[i]] = row[pos[h]]; });
    if (!rec.id) continue;                       // 主キーにできない行は捨てる
    if (!rec.sagyoubi || !rec.shimei) continue;  // 同上
    rec.kosu = Number(rec.kosu) || 0;
    for (const k of COL) if (rec[k] == null) rec[k] = '';
    nippo.push(rec);
  }

  // ★単価(rate)は落とす。給料情報をD1へ持ち込まない（2026-06-11の方針）。
  const members = (json.members || []).map(m => ({
    name: String(m.name || ''), company: String(m.company || ''), division: String(m.division || '')
  })).filter(m => m.name);

  const genba = (json.genbaMaster || []).map(g => ({
    name: String(g.name || ''), company: String(g.company || '')
  })).filter(g => g.name);

  const jobsites = (json.jobsites || []).map(j => ({
    genba: String(j.genba || ''), loc: String(j.loc || ''), jobNo: String(j.jobNo || ''),
    completed: j.completed ? 1 : 0, billingMethod: String(j.billingMethod || '')
  })).filter(j => j.genba && j.loc);

  return { nippo, members, genba, jobsites };
}

export async function fetchWithRetry(url, tries = 3) {
  let last = null;
  for (let i = 0; i < tries; i++) {
    try {
      const res = await fetch(url);
      if (!res.ok) { last = new Error('HTTP ' + res.status); continue; }
      return await res.json();   // HTMLが返ると例外になる＝リトライ対象
    } catch (e) { last = e; }
  }
  throw last || new Error('取得に失敗しました');
}

export async function syncAll(env) {
  const url = env.GAS_URL + '?compact=1&company=&t=' + Date.now();
  let parsed;
  try {
    parsed = parseGasPayload(await fetchWithRetry(url, 3));
  } catch (e) {
    return { ok: false, rows: 0, message: String(e.message || e) };
  }

  const stmts = [];
  // 全件入れ替え。日報は削除も起きるため差分ではなく総入れ替えにする。
  stmts.push(env.DB.prepare('DELETE FROM nippo'));
  const ins = env.DB.prepare(
    `INSERT OR REPLACE INTO nippo (${COL.join(',')}) VALUES (${COL.map(() => '?').join(',')})`
  );
  for (const r of parsed.nippo) stmts.push(ins.bind(...COL.map(c => r[c])));

  stmts.push(env.DB.prepare('DELETE FROM members'));
  const im = env.DB.prepare('INSERT OR REPLACE INTO members (name,company,division) VALUES (?,?,?)');
  for (const m of parsed.members) stmts.push(im.bind(m.name, m.company, m.division));

  stmts.push(env.DB.prepare('DELETE FROM genba'));
  const ig = env.DB.prepare('INSERT OR REPLACE INTO genba (name,company) VALUES (?,?)');
  for (const g of parsed.genba) stmts.push(ig.bind(g.name, g.company));

  stmts.push(env.DB.prepare('DELETE FROM jobsites'));
  const ij = env.DB.prepare(
    'INSERT OR REPLACE INTO jobsites (genba,loc,jobNo,completed,billingMethod) VALUES (?,?,?,?,?)');
  for (const j of parsed.jobsites) stmts.push(ij.bind(j.genba, j.loc, j.jobNo, j.completed, j.billingMethod));

  await env.DB.batch(stmts);   // batchは全部入るか全部入らないか
  const at = new Date().toISOString();
  await env.DB.prepare('INSERT OR REPLACE INTO sync_log (at,rows,ok,message) VALUES (?,?,?,?)')
    .bind(at, parsed.nippo.length, 1, '').run();
  return { ok: true, rows: parsed.nippo.length, message: '' };
}
```

- [ ] **Step 4: テストが通ることを確認**

Run: `cd cf && npx vitest run test/sync.test.js`
Expected: PASS（6件）

- [ ] **Step 5: コミット**

```bash
git add cf/src/sync.js cf/test/sync.test.js
git commit -m "feat(cf): GASからD1へ取り込む（404リトライ・単価は取り込まない）"
```

---

### Task 3: 読み取りAPI（doGetと同じ形で返す）

**Files:**
- Create: `cf/src/read.js`
- Create: `cf/src/index.js`
- Test: `cf/test/read.test.js`

**Interfaces:**
- Consumes: Task 1 のテーブル、Task 2 の `syncAll`
- Produces: `GET /api/schedule?company=&compact=1` が現行 `doGet` と**同じ形のJSON**を返す。`buildResponse(rows, members, genba, jobsites) -> object`

- [ ] **Step 1: 失敗するテストを書く**

`cf/test/read.test.js`:

```javascript
import { describe, it, expect } from 'vitest';
import { buildResponse, HEADERS } from '../src/read.js';

describe('buildResponse', () => {
  const row = {
    touroku:'2026-05-01T04:23:04.000Z', sagyoubi:'2026-05-02', motoukr:'NGS', genba:'大阪',
    shimei:'川端（達）', yakuwari:'代表', shukkin:'09:00', taikin:'18:00', kosu:1, memo:'',
    yakin:'', kaisha:'グローライズ', id:'abc-1', koushinsha:'森', iro:'#1D9E75',
    jigyoubu:'ICT', kouban:'INF-26-041', sagyou_kubun:'現場作業', sharyou:''
  };

  it('GASと同じ19個のヘッダを同じ順で返す', () => {
    expect(HEADERS).toEqual(['登録日時','作業日','元請名','現場名','氏名','役割','出勤','退勤',
      '人工','メモ','夜勤','会社','ID','更新者','色','事業部','工番','作業区分','車両']);
  });

  it('rowsはヘッダの順に並んだ値の配列になる', () => {
    const out = buildResponse([row], [], [], []);
    expect(out.status).toBe('ok');
    expect(out.compact).toBe(1);
    expect(out.rows[0][HEADERS.indexOf('作業日')]).toBe('2026-05-02');
    expect(out.rows[0][HEADERS.indexOf('ID')]).toBe('abc-1');
    expect(out.rows[0][HEADERS.indexOf('工番')]).toBe('INF-26-041');
    expect(out.rows[0]).toHaveLength(19);
  });

  it('職人マスタには単価を含めない', () => {
    const out = buildResponse([], [{name:'森',company:'GRHD',division:'ICT'}], [], []);
    expect(out.members[0]).toEqual({name:'森',company:'GRHD',division:'ICT'});
  });

  it('現場マスタのcompletedは真偽値に戻す（画面が真偽値で判定するため）', () => {
    const out = buildResponse([], [], [], [{genba:'A',loc:'B',jobNo:'',completed:1,billingMethod:'応援'}]);
    expect(out.jobsites[0].completed).toBe(true);
  });
});
```

- [ ] **Step 2: テストが落ちることを確認**

Run: `cd cf && npx vitest run test/read.test.js`
Expected: FAIL（`src/read.js` が無い）

- [ ] **Step 3: read.js と index.js を実装する**

`cf/src/read.js`:

```javascript
// 現行GASの doGet と "同じ形" を返す。画面側を書き換えずに差し替えるため、
// キー名も順番も1つも変えない。
export const HEADERS = ['登録日時','作業日','元請名','現場名','氏名','役割','出勤','退勤',
  '人工','メモ','夜勤','会社','ID','更新者','色','事業部','工番','作業区分','車両'];

const COL = ['touroku','sagyoubi','motoukr','genba','shimei','yakuwari','shukkin','taikin',
             'kosu','memo','yakin','kaisha','id','koushinsha','iro','jigyoubu','kouban',
             'sagyou_kubun','sharyou'];

export function buildResponse(nippo, members, genba, jobsites) {
  return {
    status: 'ok',
    compact: 1,
    headers: HEADERS,
    rows: nippo.map(r => COL.map(c => (r[c] == null ? '' : r[c]))),
    members: members.map(m => ({ name: m.name, company: m.company, division: m.division })),
    genbaMaster: genba.map(g => ({ name: g.name, company: g.company })),
    jobsites: jobsites.map(j => ({
      genba: j.genba, loc: j.loc, jobNo: j.jobNo,
      completed: !!j.completed, billingMethod: j.billingMethod
    }))
  };
}

export async function readSchedule(env, company) {
  const filter = company && company !== '全社';
  const nippo = filter
    ? await env.DB.prepare('SELECT * FROM nippo WHERE kaisha = ?').bind(company).all()
    : await env.DB.prepare('SELECT * FROM nippo').all();
  const members = filter
    ? await env.DB.prepare('SELECT * FROM members WHERE company = ?').bind(company).all()
    : await env.DB.prepare('SELECT * FROM members').all();
  const genba = await env.DB.prepare('SELECT * FROM genba').all();
  const jobsites = await env.DB.prepare('SELECT * FROM jobsites').all();

  const allowed = new Set(genba.results.filter(g => !filter || !g.company || g.company === company)
                                       .map(g => g.name));
  return buildResponse(
    nippo.results, members.results,
    genba.results.filter(g => !filter || !g.company || g.company === company),
    jobsites.results.filter(j => !filter || allowed.has(j.genba))
  );
}
```

`cf/src/index.js`:

```javascript
import { readSchedule } from './read.js';
import { syncAll } from './sync.js';

const CORS = {
  'Access-Control-Allow-Origin': 'https://akirasp4-lgtm.github.io',
  'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type'
};

const json = (obj, status = 200) =>
  new Response(JSON.stringify(obj), {
    status, headers: { 'Content-Type': 'application/json; charset=utf-8', ...CORS }
  });

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    if (request.method === 'OPTIONS') return new Response(null, { headers: CORS });

    if (url.pathname === '/api/schedule') {
      try {
        return json(await readSchedule(env, url.searchParams.get('company') || ''));
      } catch (e) {
        // 画面側は status!=='ok' を見てGASへ落ちる
        return json({ status: 'error', message: String(e.message || e) }, 500);
      }
    }

    // 書き込み直後に画面から呼ぶ。取り込んでから返す。
    if (url.pathname === '/api/sync' && request.method === 'POST') {
      const r = await syncAll(env);
      return json({ status: r.ok ? 'ok' : 'error', rows: r.rows, message: r.message });
    }

    if (url.pathname === '/api/health') {
      const last = await env.DB.prepare('SELECT * FROM sync_log ORDER BY at DESC LIMIT 1').all();
      const cnt = await env.DB.prepare('SELECT COUNT(*) AS c FROM nippo').all();
      return json({ status: 'ok', rows: cnt.results[0].c, lastSync: last.results[0] || null });
    }

    return json({ status: 'error', message: 'not found' }, 404);
  },

  async scheduled(event, env, ctx) {
    ctx.waitUntil(syncAll(env));
  }
};
```

- [ ] **Step 4: テストが通ることを確認**

Run: `cd cf && npx vitest run`
Expected: PASS（全13件）

- [ ] **Step 5: コミット**

```bash
git add cf/src/read.js cf/src/index.js cf/test/read.test.js
git commit -m "feat(cf): doGetと同形のJSONを返す読み取りAPIとCron取り込み"
```

---

### Task 4: Cloudflareへ配置して、GASと1行ずつ突き合わせる

**Files:**
- Modify: `cf/wrangler.toml`（database_id を実物へ）
- Create: `cf/test/compare.mjs`（本番同士の突き合わせスクリプト）

**Interfaces:**
- Consumes: Task 1〜3 のすべて
- Produces: 稼働中の Worker URL `https://yotei-api.<account>.workers.dev`

- [ ] **Step 1: D1 を作る**

⚠️ ここから本番リソースを作る。**実行前に利用者へ確認すること。**

```bash
cd cf && npx wrangler d1 create yotei
```

出力された `database_id` を `wrangler.toml` の PLACEHOLDER と置き換える。

- [ ] **Step 2: スキーマを流し込む**

```bash
cd cf && npx wrangler d1 execute yotei --remote --file=./schema.sql
```

- [ ] **Step 3: Worker を配置する**

```bash
cd cf && npx wrangler deploy
```

- [ ] **Step 4: 取り込みを1回走らせる**

```bash
curl -s -X POST "https://yotei-api.<account>.workers.dev/api/sync"
```

Expected: `{"status":"ok","rows":2600前後}`

- [ ] **Step 5: 突き合わせスクリプトを書いて実行する**

`cf/test/compare.mjs`:

```javascript
// GASの応答とWorkerの応答を1行ずつ突き合わせる。
// 使い方: node cf/test/compare.mjs <GAS_URL> <WORKER_URL>
const [gasUrl, workerUrl] = process.argv.slice(2);
const g = await (await fetch(gasUrl + '?compact=1&company=&t=' + Date.now())).json();
const w = await (await fetch(workerUrl + '/api/schedule?company=')).json();

const key = (h, r) => [r[h.indexOf('ID')], r[h.indexOf('作業日')], r[h.indexOf('氏名')]].join('|');
const norm = (h, r) => h.map((_, i) => String(r[i] ?? ''));

const gm = new Map(g.rows.map(r => [key(g.headers, r), norm(g.headers, r)]));
const wm = new Map(w.rows.map(r => [key(w.headers, r), norm(w.headers, r)]));

const onlyGas = [...gm.keys()].filter(k => !wm.has(k));
const onlyWorker = [...wm.keys()].filter(k => !gm.has(k));
const diff = [...gm.keys()].filter(k => wm.has(k) && JSON.stringify(gm.get(k)) !== JSON.stringify(wm.get(k)));

console.log('GAS行数    :', g.rows.length);
console.log('Worker行数 :', w.rows.length);
console.log('GASのみ    :', onlyGas.length, onlyGas.slice(0, 5));
console.log('Workerのみ :', onlyWorker.length, onlyWorker.slice(0, 5));
console.log('中身が違う :', diff.length, diff.slice(0, 5));
console.log('ヘッダ一致 :', JSON.stringify(g.headers) === JSON.stringify(w.headers));
console.log('職人数     :', g.members.length, w.members.length);
console.log('元請数     :', g.genbaMaster.length, w.genbaMaster.length);
console.log('現場数     :', g.jobsites.length, w.jobsites.length);
console.log(onlyGas.length === 0 && onlyWorker.length === 0 && diff.length === 0
  ? '=> 完全一致' : '=> 不一致あり。移行を進めないこと');
```

Run:

```bash
node cf/test/compare.mjs "https://script.google.com/macros/s/AKfycbxp2eUcpIjCj0ZWyAPPD9m3egJrKdWmXRK2AVnFrmBm4iO1QHCk-FZEH5LFFv7OloqcjQ/exec" "https://yotei-api.<account>.workers.dev"
```

Expected: `=> 完全一致`。**一致しない間は次のタスクへ進まない。**

- [ ] **Step 6: 速度を測る**

```bash
curl -s -o /dev/null -w "%{time_total}s %{size_download}bytes\n" "https://yotei-api.<account>.workers.dev/api/schedule?company=%E3%82%B0%E3%83%AD%E3%83%BC%E3%83%A9%E3%82%A4%E3%82%BA"
```

Expected: 1秒未満（GASは3.9〜56秒）。この数字を引き継ぎ.mdに記録する。

- [ ] **Step 7: コミット**

```bash
git add cf/wrangler.toml cf/test/compare.mjs
git commit -m "chore(cf): D1作成とWorker配置、GASとの全行突き合わせスクリプト"
```

---

### Task 5: 画面に切り替えスイッチと自動フォールバックを入れる

**Files:**
- Create: `backend.json`
- Modify: `index.html`（`loadData` 内のリトライ部分）
- Modify: `admin.html`（同じ箇所）

**Interfaces:**
- Consumes: Task 4 の Worker URL
- Produces: `backend.json` の1行で読み取り先が切り替わる。Worker が失敗したら自動でGASへ落ちる。

- [ ] **Step 1: backend.json を作る（最初はGASのまま＝何も変わらない）**

```json
{"backend":"gas","workerUrl":"https://yotei-api.<account>.workers.dev","note":"backendを d1 にすると読み取りだけWorkerへ切り替わる。戻すときは gas に戻す。書き込みは常にGAS。"}
```

- [ ] **Step 2: 失敗するテストを書く（ブラウザ上で確認する手順テスト）**

自動テストの土台が無い画面なので、**確認手順を先に書いて、それが通らないことを確認する**:

```
確認1: backend.json が d1 のとき、開発者ツールのネットワークに
       workers.dev への /api/schedule が出る（GASのexecは出ない）
確認2: Worker をわざと止めた状態でも画面が出る（GASへ自動で落ちる）
確認3: backend.json を gas に戻すと、次に開いたとき workers.dev を叩かない
確認4: どちらでも allNippos の件数が同じ
```

- [ ] **Step 3: loadData の取得部分を差し替える**

`index.html` / `admin.html` の `loadData` 内、2026-08-21に入れた自動リトライのループを次に置き換える:

```javascript
    const companyParam=requestCompany&&requestCompany!=='全社'?requestCompany:'';
    // 2026-08-22: 読み取り先を backend.json で切り替える。
    // ★書き込みは常にGAS。ここで切り替わるのは「読むだけ」。
    // ★Workerが失敗したら黙ってGASへ落ちる。利用者にはエラーを見せない。
    let json=null,lastErr=null,usedBackend='gas';
    let cfg={backend:'gas'};
    try{
      const c=await fetch('backend.json?t='+Date.now());
      if(c.ok)cfg=await c.json();
    }catch(e){/* 取れなければGASのまま＝安全側 */}

    const tryUrls=[];
    if(cfg.backend==='d1'&&cfg.workerUrl)
      tryUrls.push({kind:'d1',url:cfg.workerUrl+'/api/schedule?company='+encodeURIComponent(companyParam)});
    tryUrls.push({kind:'gas',url:GAS_URL+'?t='+Date.now()+'&compact=1&company='+encodeURIComponent(companyParam)});
    tryUrls.push({kind:'gas',url:GAS_URL+'?t='+Date.now()+'&compact=1&company='+encodeURIComponent(companyParam)});

    for(const t of tryUrls){
      try{
        const res=await fetch(t.url);
        const j=await res.json();
        if(j.status!=='ok')throw new Error(j.message||'予定データの取得に失敗しました');
        json=j;usedBackend=t.kind;lastErr=null;break;
      }catch(err){lastErr=err;}
    }
    if(requestSeq!==dataLoadSeq||requestCompany!==currentCompany)return;
    if(lastErr)throw lastErr;
    try{console.log('[予定管理] 読み取り元:',usedBackend);}catch(e){}
```

- [ ] **Step 4: 書き込み直後は取り込みを待ってから読み直す**

`refreshInBackground` を次に置き換える（D1が書き込み直後に古いままになるのを防ぐ）:

```javascript
function refreshInBackground(after){
  // 書き込みはGASへ行くので、D1へ取り込ませてから読み直す。
  // 取り込みに失敗しても loadData 側でGASへ落ちるので実害は無い。
  const done=()=>loadData().then(()=>{if(typeof after==='function'){try{after();}catch(e){}}}).catch(()=>{});
  fetch('backend.json?t='+Date.now()).then(r=>r.ok?r.json():null).then(cfg=>{
    if(cfg&&cfg.backend==='d1'&&cfg.workerUrl){
      return fetch(cfg.workerUrl+'/api/sync',{method:'POST'}).catch(()=>{});
    }
  }).catch(()=>{}).then(done);
}
```

- [ ] **Step 5: backend.json が gas のまま反映して、何も変わらないことを確認する**

⚠️ GitHub Pages への push は**実行前に利用者へ確認**。

```bash
git add backend.json index.html admin.html
git commit -m "feat: 読み取り先の切り替えスイッチと自動フォールバックを追加（既定はGASのまま）"
git push origin main
```

反映後、ブラウザで開いて確認1〜4のうち「確認3・4」が通ること（まだ `gas` なので workers.dev を叩かない）。

- [ ] **Step 6: 利用者の端末だけ先に切り替えて数日試す**

`backend.json` を `{"backend":"d1", ...}` にして push。確認1〜4をすべて通す。

問題が出たら `"backend":"gas"` に戻して push。**それだけで元通り。**

- [ ] **Step 7: コミット**

```bash
git add backend.json
git commit -m "chore: 読み取り先をD1へ切り替え（戻すときは backend を gas に）"
```

---

## このフェーズで達成すること / しないこと

**する**
- アプリを開くときの読み取りが 3.9〜56秒 → 1秒未満になる
- GASが5回に1回返す404の影響を受けなくなる（Workerが落ちてもGASへ自動フォールバック）

**しない（フェーズ2以降）**
- 書き込み（登録・編集・削除）は今までどおりGAS経由で約4.4秒のまま
- 集計生成・Excel出力・マスタ管理・社長カレンダー・ラーテルLINEボット連携（`vehicle_res_*` / `warehouse_today`）は**GASのまま1行も触らない**
- スプレッドシートは唯一の正であり続ける

**この設計だと、フェーズ1で失われるデータはゼロ。** D1はスプレッドシートの読み取り専用コピーなので、丸ごと消しても `/api/sync` で完全に戻る。

---

## フェーズ2の予告（着手時に別の計画書を書く）

書き込みをD1へ移す。そのとき必要になる検討事項:
- 工番の発番（現場マスタを読んで採番し書き戻す）をD1側で原子的にやる方法
- D1→スプレッドシートの書き戻し（`waitUntil` で非同期＋失敗分をCronで再送）
- 書き戻しが遅れている間、LINEボットの `warehouse_today` が当日の倉庫作業を取りこぼさないための方式
- 管理操作（`merge_genba` / `reassign_jobno` / `cleanup_orphan_jobnos`）がGAS側でシートを書き換えたあと、D1へ取り込み直す導線
