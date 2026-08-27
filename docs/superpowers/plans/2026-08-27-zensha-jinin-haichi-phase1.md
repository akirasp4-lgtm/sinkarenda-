# 全社横断 人員・案件配置管理 フェーズ1（土台）実装計画

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 予定1件ごとに「部隊」を持たせ、案件に8段階のステータスを持たせ、変更前の値が残る変更履歴を作る。フェーズ2以降（絞り込み・重複チェック・空き人員・AI提案・ダッシュボード）が乗る土台を用意する。

**Architecture:** 原本はGoogleスプレッドシートのまま。日報データに21列目「部隊」、職人マスタに「既定部隊」「有効」、現場マスタに「ステータス」を足し、新シート「変更履歴」を作る。Cloudflare Workerは読み取りの写しなので、列を通すだけ。既存20列の位置は1つも動かさない。

**Tech Stack:** Google Apps Script (gas.js) / Cloudflare Workers + D1 (cf/src) / 素のHTML+JS (index.html, admin.html) / vitest

**Spec:** `docs/superpowers/specs/2026-08-27-zensha-jinin-haichi-design.md`

## Global Constraints

- **既存20列の位置を1つも動かさない。** 動かすと集計エンジン約1,500行が全部壊れる
- **`代表` / `同行` という保存値は変えない。** 変えるのは画面の表示ラベルだけ（設計書 §3.5）
- **`president.html` と `cf/src/pres-*.js` は触らない。** 社長予定は別シートで部隊も拠点も持たない
- **`sync-guard.js` / `send-queue.js` は触らない。** 社員用・社長用の共有部品
- **デプロイ順は Worker → GAS → データ → 画面。** 画面をGASより先に出すと利用者の選択が消えて復元不能（設計書 §5）
- **既存の自動テスト270件を1件も落とさない**（ベースライン確認済み 2026-08-27）
- 部隊の値は `1部隊` `2部隊` `3部隊` `4部隊` と空欄のみ
- 案件ステータスは `見積中` `受注` `準備中` `施工中` `残工事` `完工` `延期` `中止` の8つのみ
- テスト実行: `cd cf && npx vitest run`

---

### Task 1: Worker が21列目「部隊」を受け入れられるようにする（最初に出す・無影響）

現状 `cf/src/sync.js` は「先頭19列が一致 ＋ 20列目があるなら必ず `拠点`」を検査している。
21列目 `部隊` が来ると **`20列目が「拠点」ではありません` と誤判定して取り込みが止まる**。
GASより先にここを緩める。この時点ではGASがまだ20列なので**挙動は変わらない**。

さらに `sanitizeForStorage` は職人マスタの項目を**明示的に列挙**しており（`rate` を意図的に落としている）、
ここに追記しないと `既定部隊` と `有効` が**黙って消える**。

**Files:**
- Modify: `cf/src/sync.js:36-38`（`EXPECTED_HEADERS` の下に `OPTIONAL_HEADERS` を追加）
- Modify: `cf/src/sync.js:140-146`（20列目の検査を可変長に）
- Modify: `cf/src/sync.js:161-172`（`sanitizeForStorage` の members）
- Test: `cf/test/sync.test.js`（末尾に追記）

**Interfaces:**
- Produces: `OPTIONAL_HEADERS = ['拠点','部隊']`（`cf/src/sync.js` からexport）。Task 2 のGASが書き出す21列目の名前と**必ず一致させる**
- Produces: `sanitizeForStorage` が返す member の形 `{name, company, division, butai, active}`。Task 6 の画面がこの形を読む

- [ ] **Step 1: 失敗するテストを書く**

`cf/test/sync.test.js` の末尾に追記。

⚠️ **★Codexレビュー[P3]#9: 実物の名前に合わせること。**
検査関数は `validate` ではなく **`validateGasPayload`**（`cf/src/sync.js:125`）。
テストファイル側は `EXPECTED_HEADERS` を import しておらず、
**自前の `HEADERS` という19列の定数**を先頭で定義している（`cf/test/sync.test.js:4`）。
間違えると意図した「21列で失敗」ではなく `ReferenceError` になり、
テストが通ったつもりで実装へ進んでしまう。

ファイル先頭の import に `OPTIONAL_HEADERS` を足す:

```js
import { validateGasPayload, sanitizeForStorage, fetchWithRetry, syncAll, OPTIONAL_HEADERS } from '../src/sync.js';
```

```js
describe('21列目 部隊（フェーズ1）', () => {
  const H19 = HEADERS;                    // テストファイル冒頭の自前定数（19列）
  const H20 = [...H19, '拠点'];
  const H21 = [...H19, '拠点', '部隊'];
  const base = (headers) => ({
    status: 'ok', compact: 1, headers,
    rows: [], members: [], genbaMaster: [], jobsites: []
  });

  it('21列（拠点＋部隊）を受け入れる', () => {
    expect(validateGasPayload(base(H21)).ok).toBe(true);
  });

  it('20列（拠点のみ・移行中）も受け入れる', () => {
    expect(validateGasPayload(base(H20)).ok).toBe(true);
  });

  it('19列ちょうど（さらに古い）も受け入れる', () => {
    expect(validateGasPayload(base([...H19])).ok).toBe(true);
  });

  it('20列目が拠点でなければ拒否する', () => {
    const r = validateGasPayload(base([...H19, '部隊']));
    expect(r.ok).toBe(false);
    expect(r.message).toContain('20列目');
  });

  it('21列目が部隊でなければ拒否する', () => {
    const r = validateGasPayload(base([...H19, '拠点', '班']));
    expect(r.ok).toBe(false);
    expect(r.message).toContain('21列目');
  });

  it('22列目以降は想定外として拒否する', () => {
    expect(validateGasPayload(base([...H21, '何か'])).ok).toBe(false);
  });

  it('OPTIONAL_HEADERSの順番は拠点→部隊', () => {
    expect(OPTIONAL_HEADERS).toEqual(['拠点', '部隊']);
  });

  it('sanitizeForStorage が既定部隊と有効を残す', () => {
    const out = sanitizeForStorage({
      headers: H21, rows: [], genbaMaster: [], jobsites: [],
      members: [{ name: '元', company: 'グローライズ', division: 'INF',
                  butai: '2部隊', active: false, rate: 25000 }]
    });
    expect(out.members[0]).toEqual({
      name: '元', company: 'グローライズ', division: 'INF',
      butai: '2部隊', active: false
    });
  });

  it('sanitizeForStorage は単価を落とし続ける', () => {
    const out = sanitizeForStorage({
      headers: H21, rows: [], genbaMaster: [], jobsites: [],
      members: [{ name: '元', company: 'グローライズ', division: '', rate: 25000 }]
    });
    expect(out.members[0].rate).toBeUndefined();
  });

  it('activeが無い古いGAS応答は全員 有効=true とみなす', () => {
    const out = sanitizeForStorage({
      headers: H20, rows: [], genbaMaster: [], jobsites: [],
      members: [{ name: '元', company: 'グローライズ', division: '' }]
    });
    expect(out.members[0].active).toBe(true);
    expect(out.members[0].butai).toBe('');
  });

  it('jobsitesのステータスはそのまま素通しする', () => {
    const out = sanitizeForStorage({
      headers: H21, rows: [], genbaMaster: [], members: [],
      jobsites: [{ genba: 'きんでん', loc: 'A', kyoten: '本社', status: '施工中' }]
    });
    expect(out.jobsites[0].status).toBe('施工中');
  });
});
```

- [ ] **Step 2: テストを実行して失敗を確認する**

```bash
cd cf && npx vitest run test/sync.test.js
```

想定: `OPTIONAL_HEADERS is not defined` で collect エラー、または 21列テストが `20列目が「拠点」ではありません` で FAIL。

- [ ] **Step 3: 最小の実装を書く**

`cf/src/sync.js:38` の直後（`EXPECTED_HEADERS` の定義の下）に追加:

```js
// ★2026-08-27 フェーズ1: 19列より後ろに増えてよい列と、その順番。
//   ここに書いた順番どおりでなければ取り込みを止める。
//   「増えた列が何であっても受け入れる」にすると、GASが別の列を足したとき
//   同期は成功しているのに画面はその列を見つけられず、静かな誤表示になる。
export const OPTIONAL_HEADERS = ['拠点', '部隊'];
```

`cf/src/sync.js:140-146` の20列目チェックを丸ごと置き換える:

```js
  // ★2026-08-27 フェーズ1: 19列より後ろは OPTIONAL_HEADERS の順番どおりであること。
  //   （旧: 20列目が「拠点」かどうかだけを見ていた。21列目 部隊 が来ると誤って拒否していた）
  const extraHeaders = json.headers.slice(EXPECTED_HEADERS.length);
  for (let i = 0; i < extraHeaders.length; i++) {
    const want = OPTIONAL_HEADERS[i];
    const colNo = EXPECTED_HEADERS.length + i + 1;
    if (!want) {
      return { ok: false, message: colNo + '列目は想定外の列です: ' + JSON.stringify(extraHeaders.slice(i)) };
    }
    if (extraHeaders[i] !== want) {
      return { ok: false, message: colNo + '列目が「' + want + '」ではありません: ' + JSON.stringify(extraHeaders.slice(i)) };
    }
  }
```

`cf/src/sync.js:166-168` の members を置き換える:

```js
    members: json.members.map(m => ({
      name: String(m.name || ''), company: String(m.company || ''), division: String(m.division || ''),
      // ★2026-08-27 フェーズ1: 既定部隊と有効フラグ。ここに書かないと黙って消える
      //   （rate を意図的に落としているのと同じ仕組みのため）。
      //   activeが無い＝まだ列を足していない古いGAS応答 → 全員 有効 とみなす。
      butai: String(m.butai || ''), active: m.active !== false
    })),
```

- [ ] **Step 4: テストを実行して通ることを確認する**

```bash
cd cf && npx vitest run
```

想定: 281件（既存270＋新規11）すべて PASS。

- [ ] **Step 5: コミット**

```bash
git add cf/src/sync.js cf/test/sync.test.js
git commit -m "feat(cf/sync): 21列目 部隊 を受け入れる／職人マスタの既定部隊・有効を保持する"
```

---

### Task 2: GAS — 日報データに21列目「部隊」を足す

**Files:**
- Modify: `gas.js:15`（`HEADERS`）
- Modify: `gas.js:1383-1395`（`getOrCreateMemberSheet_` を4列→6列）
- Modify: `gas.js:473-518`（`buildDailyValues_`）
- Modify: `gas.js:1292-1297`（`doGet` の members 出力）
- Create: `cf/test/gas-phase1.test.js`

**Interfaces:**
- Consumes: Task 1 の `OPTIONAL_HEADERS = ['拠点','部隊']`。21列目の名前は `'部隊'` で**完全一致**させる
- Produces: `normalizeButai_(v) -> string`（`1部隊`〜`4部隊` か空文字）
- Produces: `resolveButai_(row, memberDefault) -> string`
- Produces: `getMemberButaiMap_(ss) -> {氏名: 既定部隊}`
- Produces: doGet の member に `butai` `active` が載る。Task 1 の `sanitizeForStorage` がこの名前を読む

- [ ] **Step 1: 失敗するテストを書く**

`cf/test/gas-phase1.test.js` を新規作成。`gas.js` は Apps Script 用なので、`vm` に
最低限のグローバルを差してから読み込み、純粋な関数だけを取り出して試験する
（`cf/test/pres-*.test.js` が `president.html` に対して使っているのと同じやり方）。

```js
import { describe, it, expect, beforeAll } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const GAS_PATH = join(here, '..', '..', 'gas.js');

// ★2026-08-27 実測で判明した注意点（この方式でないと動かない）:
//   vm に読み込んでも `const HEADERS = ...` は**コンテキストの属性にならない**
//   （const/let は字句束縛でグローバルオブジェクトに載らない。var と function だけが載る）。
//   そのため gas.js の末尾に「同じ字句スコープのまま外へ出す」1行を足してから実行する。
//   ここに列挙し忘れた名前はテストから見えないので、関数を足したら必ずここにも足すこと。
const EXPORT_SNIPPET = `
;globalThis.__gas = {
  HEADERS, BUTAI_VALUES, SITE_STATUSES, SITE_STATUS_DONE,
  HISTORY_SHEET, HISTORY_HEADERS, HISTORY_SKIP_FIELDS, HISTORY_MAX_ROWS,
  KNOWN_COMPANIES,
  normalizeButai_, resolveButai_, normalizeMemberActive_,
  normalizeSiteStatus_, isSiteStatusDone_,
  diffDailyRows_, rowSummary_, sortHistoryRows_, historyTimeValue_,
  fixMojibakeCompany_, mergeMemberRows_
};`;

let ctx;   // ctx.__gas から取り出す
beforeAll(() => {
  const code = readFileSync(GAS_PATH, 'utf8');
  // Apps Script のグローバルを最低限だけ用意する（純粋関数の試験が目的）
  const sandbox = vm.createContext({
    SpreadsheetApp: { getActiveSpreadsheet: () => null },
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    Utilities: {}, ContentService: {}, PropertiesService: {},
    UrlFetchApp: {}, Logger: { log() {} }, console
  });
  vm.runInContext(code + EXPORT_SNIPPET, sandbox, { filename: 'gas.js' });
  ctx = sandbox.__gas;
});

describe('HEADERS', () => {
  it('21列で、21列目が部隊', () => {
    expect(ctx.HEADERS.length).toBe(21);
    expect(ctx.HEADERS[20]).toBe('部隊');
  });

  it('先頭19列は1つも動いていない', () => {
    expect(ctx.HEADERS.slice(0, 19)).toEqual([
      '登録日時','作業日','元請名','現場名','氏名','役割','出勤','退勤',
      '人工','メモ','夜勤','会社','ID','更新者','色','事業部','工番','作業区分','車両'
    ]);
  });

  it('20列目は拠点のまま', () => {
    expect(ctx.HEADERS[19]).toBe('拠点');
  });
});

describe('normalizeButai_', () => {
  it('1〜4部隊はそのまま通す', () => {
    ['1部隊','2部隊','3部隊','4部隊'].forEach(v =>
      expect(ctx.normalizeButai_(v)).toBe(v));
  });

  it('前後の空白を落とす', () => {
    expect(ctx.normalizeButai_('  2部隊 ')).toBe('2部隊');
  });

  it('知らない値は空にする', () => {
    ['5部隊','部隊','A班','1', 1, null, undefined, ''].forEach(v =>
      expect(ctx.normalizeButai_(v)).toBe(''));
  });
});

describe('resolveButai_', () => {
  it('画面が値を送ってきたらそれを使う', () => {
    expect(ctx.resolveButai_({ butai: '3部隊' }, '1部隊')).toBe('3部隊');
  });

  it('★画面が「空欄」を送ってきたら空欄のまま（既定値で上書きしない）', () => {
    // 事務所・休みなど「部隊に属さない」を明示できるようにするため。
    // 拠点で起きたバグ（手で消した値が既定値に戻る）を繰り返さない。
    expect(ctx.resolveButai_({ butai: '' }, '1部隊')).toBe('');
  });

  it('画面が項目そのものを送ってこなければ職人マスタの既定部隊を使う', () => {
    expect(ctx.resolveButai_({}, '1部隊')).toBe('1部隊');
  });

  it('既定部隊も無ければ空', () => {
    expect(ctx.resolveButai_({}, '')).toBe('');
    expect(ctx.resolveButai_({}, undefined)).toBe('');
  });

  it('既定部隊が壊れた値でも空にする', () => {
    expect(ctx.resolveButai_({}, '9部隊')).toBe('');
  });

  it('画面が送ってきた値が壊れていれば空（既定値へは戻さない）', () => {
    expect(ctx.resolveButai_({ butai: 'A班' }, '1部隊')).toBe('');
  });
});
```

- [ ] **Step 2: テストを実行して失敗を確認する**

```bash
cd cf && npx vitest run test/gas-phase1.test.js
```

想定: `HEADERS.length` が 20 で FAIL、`normalizeButai_ is not a function` で FAIL。

- [ ] **Step 3: 最小の実装を書く**

`gas.js:15` の `HEADERS` に `'部隊'` を足す:

```js
const HEADERS = ['登録日時','作業日','元請名','現場名','氏名','役割','出勤','退勤','人工','メモ','夜勤','会社','ID','更新者','色','事業部','工番','作業区分','車両','拠点','部隊'];
```

`gas.js` の `resolveKyoten_` の定義のすぐ下に追加:

```js
// ★2026-08-27 フェーズ1: 部隊（1〜4部隊）
//   社長決定（2026-08-24）の「1部隊・2部隊・3部隊・4部隊」。
//   ★人（職人マスタ）に固定で持たせず、予定1件ごとに持つ。
//     理由: GRは固定班ではなく案件ごとに人員を組み替える運用のため、
//     人に固定すると「同じ責任者が1部隊と2部隊を同時に持つ」が表現できない。
//     職人マスタの「既定部隊」は入力の初期値を入れるためだけに使う。
const BUTAI_VALUES = ['1部隊', '2部隊', '3部隊', '4部隊'];

function normalizeButai_(v) {
  const s = String(v == null ? '' : v).trim();
  return BUTAI_VALUES.indexOf(s) >= 0 ? s : '';
}

// ★拠点で起きたバグを繰り返さないための設計:
//   「画面が項目を送ってきたか」で分岐する。空欄を送ってきたら空欄のまま扱う。
//   （空文字を「未指定」とみなして既定値で上書きすると、
//     利用者が事務所・休みで部隊を消しても勝手に戻ってしまう）
function resolveButai_(row, memberDefault) {
  if (row && Object.prototype.hasOwnProperty.call(row, 'butai')) {
    return normalizeButai_(row.butai);
  }
  return normalizeButai_(memberDefault);
}

// 職人マスタの「既定部隊」を 氏名→部隊 の対応表にする
function getMemberButaiMap_(ss) {
  const sheet = getOrCreateMemberSheet_(ss);
  const data = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][0] || '').trim();
    if (!name) continue;
    const b = normalizeButai_(data[i][4]);
    if (b) map[name] = b;
  }
  return map;
}
```

`gas.js:1383-1395` の `getOrCreateMemberSheet_` を6列に拡張:

```js
function getOrCreateMemberSheet_(ss) {
  let sheet = ss.getSheetByName(MEMBER_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(MEMBER_SHEET);
    sheet.appendRow(['氏名', '会社', '事業部', '単価', '既定部隊', '有効']);
  } else {
    ensureColumns_(sheet, 6);
    const headers = sheet.getRange(1, 1, 1, 6).getValues()[0];
    if (String(headers[2] || '').trim() !== '事業部') sheet.getRange(1, 3).setValue('事業部');
    if (String(headers[3] || '').trim() !== '単価') sheet.getRange(1, 4).setValue('単価');
    // ★2026-08-27 フェーズ1
    if (String(headers[4] || '').trim() !== '既定部隊') sheet.getRange(1, 5).setValue('既定部隊');
    if (String(headers[5] || '').trim() !== '有効') sheet.getRange(1, 6).setValue('有効');
  }
  return sheet;
}
```

`gas.js:475` の直後（`const kyotenMap = ...` の下）に1行足す:

```js
  const butaiMap = getMemberButaiMap_(ss);
```

`gas.js:515` の `resolveKyoten_(...)` の行の**後ろに**、カンマ区切りで1要素追加:

```js
      resolveKyoten_(row.kyoten, kyotenMap[String(row.loc || '').trim()], row.company),
      // ★2026-08-27 フェーズ1: 部隊。画面が明示した値 > 職人マスタの既定部隊。
      resolveButai_(row, butaiMap[String(row.name || '').trim()])
```

**★自分の事前確認で見つけた欠陥（アーカイブのヘッダが伸びない）**

`archiveOldData_`（`gas.js:3123`）は元シートに `ensureHeaders_(sheet)` を呼ぶが、
**アーカイブシートには呼んでいない**（`gas.js:3132`）。列は
`insertColumnsAfter` で21列に伸びるものの、**21列目の見出しセルが空のまま**になる。

`sheetToRecords`（`gas.js:1563-1569`）は `headers.forEach((h,j)=>colIdx[h]=j)` と
**見出しの文字で列を引いている**ため、見出しが空だとアーカイブ側の部隊は永久に読めない。
フェーズ1の集計は部隊を使わないので今は表面化しないが、**フェーズ2の絞り込みで
「3ヶ月より前の予定だけ部隊が付かない」という分かりにくい不具合になる。**

`gas.js:3132` を次のように直す:

```js
  let archiveSheet = ss.getSheetByName(ARCHIVE_SHEET);
  if (!archiveSheet) { archiveSheet = ss.insertSheet(ARCHIVE_SHEET); archiveSheet.appendRow(HEADERS); }
  // ★2026-08-27 フェーズ1: 既にあるアーカイブも見出しを最新の列数に揃える。
  //   これが無いと列だけ21に伸びて見出しが空のままになり、
  //   sheetToRecords（見出しの文字で列を引く）がアーカイブの部隊を読めない。
  ensureHeaders_(archiveSheet);
```

**`add_member` も6列で書くようにする（`gas.js:800`）**

今は `memberSheet.appendRow([name, company, division, rate]);` と4つしか書いていない。
このままでも「有効」が空欄＝有効扱いになるので壊れはしないが、
掃除のあとに追加した人だけ空欄が並んで分かりにくいので揃える:

```js
      // ★2026-08-27 フェーズ1: 既定部隊は空、有効は○で作る
      memberSheet.appendRow([name, company, division, rate, '', '○']);
```

`gas.js:1292-1297` の members 出力を差し替える:

```js
    const members = mData.length > 1 ? mData.slice(1).map(r => ({
      name: String(r[0]||''),
      company: String(r[1]||''),
      division: String(r[2]||''),
      rate: Number(r[3]||0),
      // ★2026-08-27 フェーズ1
      butai: normalizeButai_(r[4]),
      // 「有効」列が空欄＝まだ何も入れていない → 有効とみなす（既存71件を巻き込まない）
      active: String(r[5]||'').trim() !== '×'
    })).filter(m => !filterByCompany || m.company === requestedCompany) : [];
```

- [ ] **Step 4: テストを実行して通ることを確認する**

```bash
cd cf && npx vitest run
```

想定: 全件 PASS（既存281＋新規14＝295）。

- [ ] **Step 5: コミット**

```bash
git add gas.js cf/test/gas-phase1.test.js
git commit -m "feat(gas): 日報データに21列目 部隊 を追加／職人マスタに既定部隊・有効"
```

---

### Task 3: GAS — 現場マスタに8段階のステータスを足す

既存の `完了`（真偽値）は集計が参照しているので**残す**。ステータスを正とし、`完了` を従属させる。
移行は**読むときに導出する**方式にして、既存184件の書き換えを不要にする。

**Files:**
- Modify: `gas.js:1411-1427`（`getOrCreateJobSiteSheet_` を11列→12列）
- Modify: `gas.js:1336-1342`（`doGet` の jobsites 出力）
- Modify: `gas.js:940-960` 付近（`update_site_status` アクション）
- Test: `cf/test/gas-phase1.test.js`（追記）

**Interfaces:**
- Produces: `SITE_STATUSES`（8つの配列）／`SITE_STATUS_DONE = ['完工','中止']`
- Produces: `normalizeSiteStatus_(raw, completedCell) -> string`（必ず8つのどれかを返す）
- Produces: doGet の jobsite に `status` が載る。Task 8 の画面がこれを読む

- [ ] **Step 1: 失敗するテストを書く**

`cf/test/gas-phase1.test.js` に追記:

```js
describe('案件ステータス（8段階）', () => {
  it('8つちょうど、順番も依頼どおり', () => {
    expect(ctx.SITE_STATUSES).toEqual([
      '見積中','受注','準備中','施工中','残工事','完工','延期','中止'
    ]);
  });

  it('完了扱いは完工と中止だけ', () => {
    expect(ctx.SITE_STATUS_DONE).toEqual(['完工','中止']);
  });

  it('保存済みの正しい値はそのまま返す', () => {
    ctx.SITE_STATUSES.forEach(s =>
      expect(ctx.normalizeSiteStatus_(s, false)).toBe(s));
  });

  it('★未設定（空欄）は 完了 列から導く＝既存184件を書き換えずに移行できる', () => {
    expect(ctx.normalizeSiteStatus_('', true)).toBe('完工');
    expect(ctx.normalizeSiteStatus_('', false)).toBe('施工中');
    expect(ctx.normalizeSiteStatus_(undefined, true)).toBe('完工');
  });

  it('完了列が文字列の TRUE でも完工と判定する', () => {
    expect(ctx.normalizeSiteStatus_('', 'TRUE')).toBe('完工');
    expect(ctx.normalizeSiteStatus_('', '完了')).toBe('完工');
  });

  it('知らない値は 完了 列から導き直す（勝手な値を通さない）', () => {
    expect(ctx.normalizeSiteStatus_('進行中', false)).toBe('施工中');
    expect(ctx.normalizeSiteStatus_('やめた', true)).toBe('完工');
  });

  it('前後の空白を落とす', () => {
    expect(ctx.normalizeSiteStatus_(' 残工事 ', false)).toBe('残工事');
  });

  it('isSiteStatusDone_ が 完了 列に書く値を決める', () => {
    expect(ctx.isSiteStatusDone_('完工')).toBe(true);
    expect(ctx.isSiteStatusDone_('中止')).toBe(true);
    ['見積中','受注','準備中','施工中','残工事','延期'].forEach(s =>
      expect(ctx.isSiteStatusDone_(s)).toBe(false));
  });
});
```

- [ ] **Step 2: テストを実行して失敗を確認する**

```bash
cd cf && npx vitest run test/gas-phase1.test.js
```

想定: `SITE_STATUSES is not defined` で FAIL。

- [ ] **Step 3: 最小の実装を書く**

`gas.js` の `BUTAI_VALUES` の定義の下に追加:

```js
// ★2026-08-27 フェーズ1: 案件ステータス（依頼10項目の⑧）
//   既存の「完了」列（真偽値）は集計エンジンが参照しているので消さない。
//   ステータスを正とし、完了列を従属させる（二重管理を避けるため）。
const SITE_STATUSES = ['見積中', '受注', '準備中', '施工中', '残工事', '完工', '延期', '中止'];
const SITE_STATUS_DONE = ['完工', '中止'];

function isSiteStatusDone_(status) {
  return SITE_STATUS_DONE.indexOf(String(status || '').trim()) >= 0;
}

// ★移行は「読むときに導出する」方式。既存184件を書き換えなくてよい。
//   ステータス列が空欄・壊れた値のときは、今までの「完了」列から導く。
function normalizeSiteStatus_(raw, completedCell) {
  const s = String(raw == null ? '' : raw).trim();
  if (SITE_STATUSES.indexOf(s) >= 0) return s;
  const c = String(completedCell == null ? '' : completedCell).trim().toUpperCase();
  const done = completedCell === true || c === 'TRUE' || c === '完了' || c === '1';
  return done ? '完工' : '施工中';
}
```

`gas.js:1411-1427` の `getOrCreateJobSiteSheet_` を12列に:

```js
function getOrCreateJobSiteSheet_(ss) {
  let sheet = ss.getSheetByName(JOBSITE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(JOBSITE_SHEET);
    sheet.appendRow(['元請名', '現場名', '工番', '事業部', '年度', '連番', '売上', '読み', '完了', '請求方式', '拠点', 'ステータス']);
  } else {
    ensureColumns_(sheet, 12);
    const headers = sheet.getRange(1, 1, 1, 12).getValues()[0];
    if (String(headers[6] || '').trim() !== '売上') sheet.getRange(1, 7).setValue('売上');
    if (String(headers[7] || '').trim() !== '読み') sheet.getRange(1, 8).setValue('読み');
    if (String(headers[8] || '').trim() !== '完了') sheet.getRange(1, 9).setValue('完了');
    if (String(headers[9] || '').trim() !== '請求方式') sheet.getRange(1, 10).setValue('請求方式');
    if (String(headers[10] || '').trim() !== '拠点') sheet.getRange(1, 11).setValue('拠点');
    // ★2026-08-27 フェーズ1: 案件ステータス（8段階）
    if (String(headers[11] || '').trim() !== 'ステータス') sheet.getRange(1, 12).setValue('ステータス');
  }
  return sheet;
}
```

`gas.js:1341` の `kyoten: String(r[10] || '').trim()` の後ろにカンマ区切りで追加:

```js
      kyoten: String(r[10] || '').trim(),
      // ★2026-08-27 フェーズ1: 未設定なら「完了」列から導く（既存184件は無書き換えで移行）
      status: normalizeSiteStatus_(r[11], r[8])
```

`gas.js:940-960` 付近の `update_site_status` アクションを探し、`completed` を書いている箇所の
**直後に**ステータス列も書くよう追記する。さらに新しいアクション `set_site_status` を
同じ `if (action === ...)` の並びに追加する:

```js
    if (action === 'set_site_status') {
      const genba = String(body.genba || '').trim();
      const loc = String(body.loc || '').trim();
      const status = String(body.status || '').trim();
      if (SITE_STATUSES.indexOf(status) < 0) return error('知らないステータスです: ' + status);
      const sheet = getOrCreateJobSiteSheet_(ss);
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0] || '').trim() === genba && String(data[i][1] || '').trim() === loc) {
          sheet.getRange(i + 1, 12).setValue(status);
          // ★完了列を従属させる（集計エンジンがこちらを見ているため必ず揃える）
          sheet.getRange(i + 1, 9).setValue(isSiteStatusDone_(status));
          logOperation_(ss, 'set_site_status', genba + '/' + loc, 'ステータス=' + status, updatedBy);
          return ok({ status: status, completed: isSiteStatusDone_(status) });
        }
      }
      return error('現場が見つかりません: ' + genba + '/' + loc);
    }
```

- [ ] **Step 4: テストを実行して通ることを確認する**

```bash
cd cf && npx vitest run
```

想定: 全件 PASS（＋新規8）。

- [ ] **Step 5: コミット**

```bash
git add gas.js cf/test/gas-phase1.test.js
git commit -m "feat(gas): 現場マスタに案件ステータス8段階を追加（完了列は従属させる）"
```

---

### Task 4: GAS — 変更履歴シート（変更前の値が残る）

依頼⑦「誰が登録・変更・削除したか」「**元の予定も確認できる**」。
今の `操作ログ` は `行数=2` としか残っていないので、変更前が復元できない。

`update` は「新しい行を追加 → 古い行を削除」（`gas.js:600-617`）なので、
**削除する直前に古い行の値が手元にある**。ここで拾って記録する。
編集するとIDが変わるため、**旧ID→新IDを必ず記録**する。

**Files:**
- Modify: `gas.js`（`OPLOG_SHEET` 定義の近くに `HISTORY_SHEET` を追加）
- Modify: `gas.js:1499-1512` 付近（`getOrCreateOpLogSheet_` の下に履歴用を追加）
- Modify: `gas.js:575-620`（`add` / `delete` / `update` の3アクション）
- Test: `cf/test/gas-phase1.test.js`（追記）

**Interfaces:**
- Produces: `HISTORY_SHEET = '変更履歴'`、8列 `['日時','操作','旧ID','新ID','項目','変更前','変更後','実行者']`
- Produces: `diffDailyRows_(headers, oldRows, newRows) -> [{oldId, newId, field, before, after}]`
- Produces: `getOrCreateHistorySheet_(ss)` / `logHistory_(ss, entries, action, user)`

- [ ] **Step 1: 失敗するテストを書く**

`cf/test/gas-phase1.test.js` に追記:

```js
describe('変更履歴 diffDailyRows_', () => {
  // gas.js の HEADERS 21列に合わせた行を作る補助
  const H = () => ctx.HEADERS;
  const mkRow = (over) => {
    const base = {
      '登録日時': '2026/08/27 10:00', '作業日': '2026-08-28', '元請名': 'きんでん西',
      '現場名': 'A現場', '氏名': '元', '役割': '代表', '出勤': '08:00', '退勤': '17:00',
      '人工': 1, 'メモ': '', '夜勤': '', '会社': 'グローライズ', 'ID': 'X1',
      '更新者': '向', '色': '', '事業部': 'INF', '工番': 'INF-26-001',
      '作業区分': '現場作業', '車両': '', '拠点': '本社', '部隊': '1部隊'
    };
    Object.assign(base, over || {});
    return H().map(h => base[h]);
  };

  it('変わった項目だけを返す', () => {
    const oldR = mkRow({ ID: 'X1' });
    const newR = mkRow({ ID: 'X2', '現場名': 'B現場' });
    const d = ctx.diffDailyRows_(H(), [oldR], [newR]);
    const fields = d.map(x => x.field);
    expect(fields).toContain('現場名');
    expect(fields).not.toContain('元請名');
  });

  it('変更前と変更後の両方が残る（元の予定が確認できる）', () => {
    const d = ctx.diffDailyRows_(H(), [mkRow({ '人工': 1 })], [mkRow({ ID: 'X2', '人工': 0.5 })]);
    const k = d.find(x => x.field === '人工');
    expect(String(k.before)).toBe('1');
    expect(String(k.after)).toBe('0.5');
  });

  it('★旧IDと新IDが繋がる（編集するとIDが変わるため）', () => {
    const d = ctx.diffDailyRows_(H(), [mkRow({ ID: 'OLD' })], [mkRow({ ID: 'NEW', 'メモ': 'あ' })]);
    expect(d[0].oldId).toBe('OLD');
    expect(d[0].newId).toBe('NEW');
  });

  it('登録日時は毎回変わるので履歴に出さない', () => {
    const oldR = mkRow({ '登録日時': '2026/08/27 10:00' });
    const newR = mkRow({ ID: 'X2', '登録日時': '2026/08/27 11:00' });
    expect(ctx.diffDailyRows_(H(), [oldR], [newR]).map(x => x.field)).not.toContain('登録日時');
  });

  it('IDそのものは項目としては出さない（旧ID/新IDの欄で見えるため）', () => {
    const d = ctx.diffDailyRows_(H(), [mkRow({ ID: 'A' })], [mkRow({ ID: 'B', 'メモ': 'x' })]);
    expect(d.map(x => x.field)).not.toContain('ID');
  });

  it('何も変わっていなければ空を返す', () => {
    expect(ctx.diffDailyRows_(H(), [mkRow()], [mkRow()])).toEqual([]);
  });

  it('人が増えた（追加された）行は 追加 として出る', () => {
    const d = ctx.diffDailyRows_(H(), [mkRow({ '氏名': '元' })],
      [mkRow({ ID: 'X2', '氏名': '元' }), mkRow({ ID: 'X3', '氏名': '中島' })]);
    const add = d.find(x => x.field === '(追加)');
    expect(add).toBeTruthy();
    expect(add.after).toContain('中島');
  });

  it('人が減った（外された）行は 削除 として出る', () => {
    const d = ctx.diffDailyRows_(H(),
      [mkRow({ '氏名': '元' }), mkRow({ ID: 'X9', '氏名': '中島' })],
      [mkRow({ ID: 'X2', '氏名': '元' })]);
    const del = d.find(x => x.field === '(削除)');
    expect(del).toBeTruthy();
    expect(del.before).toContain('中島');
  });

  it('★同じ人が同じ日に2件ある場合も、2件目の変更を取りこぼさない', () => {
    // 本番に250件ある形（現場＋事務所 など）。連番を鍵に混ぜていないと握りつぶされる。
    const oldRows = [
      mkRow({ ID: 'A1', '現場名': 'A現場', '作業区分': '現場作業' }),
      mkRow({ ID: 'A2', '現場名': '事務所', '作業区分': '事務所' })
    ];
    const newRows = [
      mkRow({ ID: 'B1', '現場名': 'A現場', '作業区分': '現場作業' }),
      mkRow({ ID: 'B2', '現場名': '事務所', '作業区分': '事務所', '人工': 0.5 })
    ];
    const d = ctx.diffDailyRows_(H(), oldRows, newRows);
    const k = d.find(x => x.field === '人工');
    expect(k).toBeTruthy();
    expect(k.oldId).toBe('A2');
    expect(k.newId).toBe('B2');
    // 2件目を「削除された」と誤記録していないこと
    expect(d.find(x => x.field === '(削除)')).toBeUndefined();
  });

  it('★同じ人が同じ日に2件→1件に減ったら、減った1件だけを削除として記録する', () => {
    const oldRows = [
      mkRow({ ID: 'A1', '現場名': 'A現場' }),
      mkRow({ ID: 'A2', '現場名': '事務所' })
    ];
    const newRows = [mkRow({ ID: 'B1', '現場名': 'A現場' })];
    const d = ctx.diffDailyRows_(H(), oldRows, newRows);
    const del = d.filter(x => x.field === '(削除)');
    expect(del.length).toBe(1);
    expect(del[0].oldId).toBe('A2');
  });

  it('部隊の変更も拾う', () => {
    const d = ctx.diffDailyRows_(H(), [mkRow({ '部隊': '1部隊' })],
      [mkRow({ ID: 'X2', '部隊': '3部隊' })]);
    const k = d.find(x => x.field === '部隊');
    expect(k.before).toBe('1部隊');
    expect(k.after).toBe('3部隊');
  });

  it('日付や時刻の型が違っても文字列として比べる（誤検知しない）', () => {
    const oldR = mkRow({ '人工': 1 });
    const newR = mkRow({ ID: 'X2', '人工': '1' });
    expect(ctx.diffDailyRows_(H(), [oldR], [newR])).toEqual([]);
  });
});

describe('変更履歴シートの形', () => {
  it('8列で、依頼どおりの並び', () => {
    expect(ctx.HISTORY_HEADERS).toEqual(
      ['日時','操作','旧ID','新ID','項目','変更前','変更後','実行者']);
  });
  it('シート名は 変更履歴', () => {
    expect(ctx.HISTORY_SHEET).toBe('変更履歴');
  });
});
```

- [ ] **Step 2: テストを実行して失敗を確認する**

```bash
cd cf && npx vitest run test/gas-phase1.test.js
```

想定: `diffDailyRows_ is not a function` で FAIL。

- [ ] **Step 3: 最小の実装を書く**

`gas.js:14`（`const OPLOG_SHEET = '操作ログ';`）の下に追加:

```js
const HISTORY_SHEET = '変更履歴';
```

`gas.js:1499` の `getOrCreateOpLogSheet_` の下に追加:

```js
// ★2026-08-27 フェーズ1: 変更履歴（依頼⑦）
//   操作ログ（誰が何をしたか）とは別物。こちらは「何が何に変わったか」を残す。
//   両方残す：操作ログはアーカイブ・マージ等の管理操作も記録しており用途が違う。
const HISTORY_HEADERS = ['日時', '操作', '旧ID', '新ID', '項目', '変更前', '変更後', '実行者'];

function getOrCreateHistorySheet_(ss) {
  let sheet = ss.getSheetByName(HISTORY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(HISTORY_SHEET);
    sheet.appendRow(HISTORY_HEADERS);
  } else {
    ensureColumns_(sheet, HISTORY_HEADERS.length);
  }
  return sheet;
}

// 履歴に出さない列。登録日時は毎回変わるので出すと差分が埋もれる。
// IDは旧ID/新IDの欄で見えるので項目としては重複。
const HISTORY_SKIP_FIELDS = ['登録日時', 'ID'];

function rowSummary_(headers, arr) {
  const g = (h) => {
    const i = headers.indexOf(h);
    return i >= 0 ? String(arr[i] == null ? '' : arr[i]).trim() : '';
  };
  return [g('作業日'), g('氏名'), g('元請名') + '/' + g('現場名'), g('作業区分')]
    .filter(Boolean).join(' ');
}

/**
 * 編集前後の行を突き合わせ、変わった項目だけを返す。
 * 突き合わせの鍵は「作業日＋氏名＋その中での何番目か」。
 * IDは編集のたびに変わるので鍵に使えない。
 *
 * ★「何番目か」を鍵に混ぜる理由（実データで確認・2026-08-27）:
 *   同じ人が同じ日に複数行を持つ組み合わせが本番に250件ある
 *   （現場＋事務所、昼＋夜勤 など）。単純に「作業日＋氏名」を鍵にすると
 *   2件目以降が握りつぶされ、2件目を直した履歴が残らないうえ、
 *   件数が変わったときに「削除された」と誤って記録してしまう。
 *
 * 戻り値: [{oldId, newId, field, before, after}]
 */
function diffDailyRows_(headers, oldRows, newRows) {
  const idIdx = headers.indexOf('ID');
  const dIdx = headers.indexOf('作業日');
  const nIdx = headers.indexOf('氏名');
  const cell = (arr, i) => String(arr && arr[i] != null ? arr[i] : '').trim();

  // 同じ「作業日＋氏名」が複数あっても取りこぼさないよう連番を振る
  const indexRows = (rows) => {
    const map = {}, seen = {};
    (rows || []).forEach(r => {
      const base = cell(r, dIdx) + '|' + cell(r, nIdx);
      seen[base] = (seen[base] || 0) + 1;
      map[base + '#' + seen[base]] = r;
    });
    return map;
  };
  const oldMap = indexRows(oldRows);
  const newMap = indexRows(newRows);
  // ※ oldRows はシートを下から上へ走査して積むため画面の順と逆になる。
  //    呼び出す前に必ず reverse() して渡すこと（★Codexレビュー[P2]#5）。

  const out = [];
  Object.keys(oldMap).forEach(k => {
    const o = oldMap[k];
    const n = newMap[k];
    if (!n) {
      out.push({ oldId: cell(o, idIdx), newId: '', field: '(削除)',
                 before: rowSummary_(headers, o), after: '' });
      return;
    }
    headers.forEach((h, i) => {
      if (HISTORY_SKIP_FIELDS.indexOf(h) >= 0) return;
      const a = cell(o, i), b = cell(n, i);
      if (a !== b) {
        out.push({ oldId: cell(o, idIdx), newId: cell(n, idIdx), field: h, before: a, after: b });
      }
    });
  });
  Object.keys(newMap).forEach(k => {
    if (oldMap[k]) return;
    const n = newMap[k];
    out.push({ oldId: '', newId: cell(n, idIdx), field: '(追加)',
               before: '', after: rowSummary_(headers, n) });
  });

  // ★Codexレビュー[P2]#5: 業務項目に差が無くてもIDは編集のたびに必ず変わる。
  //   差分ゼロで何も残さないと、そこで旧ID→新IDの鎖が切れて過去へ遡れなくなる。
  //   1件も出力が無いときだけ、対応関係だけを記録する行を残す。
  if (!out.length) {
    Object.keys(oldMap).forEach(k => {
      const o = oldMap[k], n = newMap[k];
      if (!n) return;
      const oid = cell(o, idIdx), nid = cell(n, idIdx);
      if (oid && nid && oid !== nid) {
        out.push({ oldId: oid, newId: nid, field: '(ID引継ぎ)', before: oid, after: nid });
      }
    });
  }
  return out;
}

// 履歴をまとめて1回で書く（appendRowを繰り返すとGASが遅くなるため）
//
// ★Codexレビュー[P1]#2: ここで例外を握りつぶしてはいけない。
//   握りつぶすと「履歴が書けなかったのに旧予定は消えた」＝
//   元の予定を永久に復元できない状態が、成功応答のまま起きる。
//   呼び出し側は、この関数が成功したことを確認してから原本を消すこと。
//   SpreadsheetApp.flush() まで通して初めて「書けた」と言える。
function logHistory_(ss, action, entries, user) {
  if (!entries || !entries.length) return 0;
  const sheet = getOrCreateHistorySheet_(ss);
  const now = new Date().toLocaleString('ja-JP');
  const values = entries.map(e => [
    now, action, e.oldId || '', e.newId || '', e.field || '',
    e.before == null ? '' : e.before, e.after == null ? '' : e.after, user || ''
  ]);
  sheet.getRange(sheet.getLastRow() + 1, 1, values.length, HISTORY_HEADERS.length)
       .setValues(values);
  SpreadsheetApp.flush();     // 実際に書き終わるまで待つ
  return values.length;
}

// ★Codexレビュー[P1]#3: 削除は「全項目」を残さないと復元できない。
//   要約（作業日・氏名・現場・区分）だけでは、役割・出退勤・人工・メモ・夜勤・
//   会社・事業部・工番・車両・拠点・部隊が失われる。
//   バックアップ後に登録された予定を消したら、二度と戻せない。
function rowFullJson_(headers, arr) {
  const o = {};
  headers.forEach((h, i) => {
    const v = arr && arr[i];
    o[h] = (v == null) ? '' : String(v);
  });
  return JSON.stringify(o);
}
```

`gas.js` の `add` アクション（`logOperation_(ss, 'add', ...)` の行の**直前**）に追加:

```js
      // addだけは握りつぶす。履歴が書けなくても原本の予定は正しく保存されており
      // 失うものが無い。ここでエラーにすると履歴シートの不調で新規登録が全部止まる。
      try {
        logHistory_(ss, 'add', values.map(v => ({
          oldId: '', newId: String(v[idCol] || ''), field: '(新規)',
          before: '', after: rowSummary_(HEADERS, v)
        })), updatedBy);
      } catch (e) {}
```

`gas.js:586-598` の `delete` アクションを差し替える。

**★Codexレビュー[P1]#2#3**: 削除も「履歴を書いてから消す」。
さらに `変更前` には**21列すべてをJSONで**残す。要約だけでは復元できない。

```js
    if (action === 'delete') {
      const ids = body.ids || [];
      if (ids.length === 0) return ok({deleted: 0});
      const data = sheet.getDataRange().getValues();
      const rowsToDelete = [];
      const deletedRows = [];                   // ★2026-08-27 消す前に中身を控える
      for (let i = data.length - 1; i >= 1; i--) {
        const rowId = String(data[i][idCol] || '').trim();
        if (rowId && ids.includes(rowId)) {
          rowsToDelete.push(i + 1);
          deletedRows.push(data[i]);
        }
      }

      // ★1) 先に履歴を書く。失敗したら1行も消さずにエラーで返す。
      try {
        logHistory_(ss, 'delete', deletedRows.slice().reverse().map(r => ({
          oldId: String(r[idCol] || ''), newId: '', field: '(削除)',
          // 全項目をJSONで残す＝ここから元の予定をそのまま作り直せる
          before: rowFullJson_(HEADERS, r), after: ''
        })), updatedBy);
      } catch (e) {
        return error('変更履歴を記録できなかったため、削除を中止しました（予定はそのまま残っています）: ' + e);
      }

      // ★2) 履歴が残ってから消す
      rowsToDelete.forEach(rowNum => sheet.deleteRow(rowNum));
      logOperation_(ss, 'delete', 'IDs=' + ids.length + '件', '削除行=' + rowsToDelete.length, updatedBy);
      return ok({deleted: rowsToDelete.length, requested: ids.length});
    }
```

⚠️ **`add` は握りつぶしたままでよい**（何も壊れないため）。
新規登録で履歴が書けなくても、原本の予定は正しく保存されており失うものが無い。
逆にここでエラーにすると、履歴シートの不調だけで**新規登録が全部止まる**。
`add` だけは `try { logHistory_(...) } catch (e) {}` で囲む。

`gas.js:601-617` の `update` アクションを差し替える:

**★Codexレビュー[P1]#2: 順番が命。**
元の計画は「新行を足す → 旧行を消す → 履歴を書く」だった。
履歴の書き込みが失敗すると **旧予定は既に消えているのに成功応答**になり、
依頼⑦の「元の予定も確認できる」が保証できない。

正しい順番は **「値を作る → 履歴を書く（flushまで確認）→ 新行を足す → 旧行を消す」**。
IDは画面が作って送ってくるので、行を足す前に新IDが分かる＝先に履歴を書ける。
履歴で例外が出れば **原本は1行も動いていない**状態で止まる。

```js
    if (action === 'update') {
      const rows = requireDailyRows_(body);
      const ids = body.ids || [];
      const rowsToDelete = [];
      const oldRows = [];                       // ★2026-08-27 変更前の値を控える
      if (ids.length > 0) {
        const data = sheet.getDataRange().getValues();
        for (let i = data.length - 1; i >= 1; i--) {
          const rowId = String(data[i][idCol] || '').trim();
          if (rowId && ids.includes(rowId)) {
            rowsToDelete.push(i + 1);
            oldRows.push(data[i]);              // 消す前にここで拾う
          }
        }
      }
      const values = buildDailyValues_(ss, rows, updatedBy);

      // ★1) 先に履歴を書く。ここで失敗したら原本は1行も動かさずにエラーで返す。
      //    oldRowsは下から積んでいるので reverse して画面の順に戻す（[P2]#5）。
      try {
        logHistory_(ss, 'update', diffDailyRows_(HEADERS, oldRows.slice().reverse(), values), updatedBy);
      } catch (e) {
        return error('変更履歴を記録できなかったため、保存を中止しました（元の予定はそのまま残っています）: ' + e);
      }

      // ★2) 履歴が残ってから原本を触る
      appendDailyValues_(sheet, values);
      rowsToDelete.forEach(rowNum => sheet.deleteRow(rowNum));

      logOperation_(ss, 'update', rows[0].genba + '/' + (rows[0].loc || ''), '行数=' + rows.length + ', 旧ID=' + ids.length, updatedBy);
      return ok({updated: rows.length});
    }
```

- [ ] **Step 4: テストを実行して通ることを確認する**

```bash
cd cf && npx vitest run
```

想定: 全件 PASS（＋新規12）。

- [ ] **Step 5: コミット**

```bash
git add gas.js cf/test/gas-phase1.test.js
git commit -m "feat(gas): 変更履歴シート（変更前の値と旧ID→新IDが残る）"
```

---

### Task 5: GAS — データ掃除（★必ず dry-run を先に出し、推測で書き換えない）

設計書 §3.6。**壊れたデータを直す。行数は変えない。**

**★Codexレビュー[P1]#4 で当初案を全面的に作り直した。** 当初案には2つの重大な誤りがあった:

1. **重複を「先勝ちで捨てる」実装だった** — 2行目以降の 事業部・**単価**・既定部隊を
   比較せず破棄していた。**単価は給料の元数字**であり、2行目に正しい値が入っていたら失われる。
   「統合」と書いておきながら実際は削除だった。
2. **`/応援/` という正規表現で「人でない枠」を推測していた** — これは**間違い**。
   予定が0件の14人を実データで確認したところ、`川端` `井上` `作本` `児玉` `杉本仁（兄）`
   `いくや` など**実在の人が多数含まれる**。予定が無いことと人でないことは別の話であり、
   推測で無効化すると**実在の職人が空き人員リストから消える**。

**新しい方針: 掃除は「機械が確実に判断できること」だけやる。人の判断が要るものは一覧を出すだけ。**

| 対象 | 機械がやる | 人が決める |
|---|---|---|
| 会社名の文字化け（1件） | ✅ 直す（既知の会社名と一意に決まるため確実） | — |
| 職人マスタの氏名重複（9件） | ✅ 非空の値だけを寄せて統合。**値が食い違ったら中止して一覧を出す** | 食い違った場合のみ |
| 予定に出るがマスタに無い（1件） | ✅ 追加（`有効=○`） | — |
| 「人でない枠」の無効化（14件候補） | ❌ **やらない**。候補一覧を出すだけ | ✅ Task 6B の画面で1件ずつ |
| 既存2,664行への部隊の流し込み | ❌ **やらない**（Task 5B 参照） | ✅ 割当表を決めてから別途 |

**Files:**
- Modify: `gas.js`（末尾に追加）
- Test: `cf/test/gas-phase1.test.js`（追記）

**Interfaces:**
- Produces: `cleanupMastersPhase1(apply)` — 引数なし/`false` で **dry-run**（何も書かない）、
  `true` で実行。Apps Scriptのエディタから手で実行する
- Produces: `fixMojibakeCompany_(v) -> string`
- Produces: `mergeMemberRows_(rows) -> {merged, conflicts}`

- [ ] **Step 1: 失敗するテストを書く**

```js
describe('データ掃除', () => {
  const MOJI = '�';   // 文字化けを表す記号

  it('文字化けした会社名を直す', () => {
    expect(ctx.fixMojibakeCompany_('グロ' + MOJI + 'ライズ')).toBe('グローライズ');
    expect(ctx.fixMojibakeCompany_('グロ?ライズ')).toBe('グローライズ');
  });

  it('正しい会社名はそのまま返す', () => {
    ['グローライズ','和信カインド','GRミツマ','GRHD','ラーテル'].forEach(c =>
      expect(ctx.fixMojibakeCompany_(c)).toBe(c));
  });

  it('関係ない文字列は触らない', () => {
    expect(ctx.fixMojibakeCompany_('よその会社')).toBe('よその会社');
    expect(ctx.fixMojibakeCompany_('')).toBe('');
  });

  it('化け方が曖昧でどの会社か決まらないときは触らない', () => {
    const v = MOJI.repeat(5);
    expect(ctx.fixMojibakeCompany_(v)).toBe(v);
  });

  it('★重複行は非空の値を寄せて統合する（先勝ちで捨てない）', () => {
    const r = ctx.mergeMemberRows_([
      ['元','グローライズ','',0,'',''],
      ['元','グローライズ','INF',25000,'2部隊','']
    ]);
    expect(r.conflicts.length).toBe(0);
    expect(r.merged.length).toBe(1);
    expect(r.merged[0][2]).toBe('INF');
    expect(r.merged[0][3]).toBe(25000);      // ★単価を失わない
    expect(r.merged[0][4]).toBe('2部隊');
  });

  it('★値が食い違ったら統合せず conflicts に出す（勝手に決めない）', () => {
    const r = ctx.mergeMemberRows_([
      ['元','グローライズ','INF',25000,'',''],
      ['元','グローライズ','ICT',30000,'','']
    ]);
    expect(r.conflicts.length).toBeGreaterThan(0);
    expect(r.conflicts[0].name).toBe('元');
  });

  it('会社が違えば別人として扱う（同姓の別会社を潰さない）', () => {
    const r = ctx.mergeMemberRows_([
      ['元','グローライズ','INF',25000,'',''],
      ['元','和信カインド','',0,'','']
    ]);
    expect(r.merged.length).toBe(2);
    expect(r.conflicts.length).toBe(0);
  });

  it('★「人でない枠」を機械で判定する関数は存在しない（推測しない設計）', () => {
    expect(ctx.looksLikeNonPerson_).toBeUndefined();
  });
});
```

- [ ] **Step 2: テストを実行して失敗を確認する**

```bash
cd cf && npx vitest run test/gas-phase1.test.js
```

想定: `fixMojibakeCompany_ is not a function` で FAIL。

- [ ] **Step 3: 最小の実装を書く**

`gas.js` の末尾に追加:

```js
// ★2026-08-27 フェーズ1: データ掃除
//   本番実データで見つかった汚れ（設計書 §1.3）を直す。行数は変えない。
//   ★方針: 機械が確実に判断できることだけ直す。人の判断が要るものは一覧を出すだけ。

const KNOWN_COMPANIES = ['グローライズ', '和信カインド', 'GRミツマ', 'GRHD', 'ラーテル'];
const MOJIBAKE_RE = /[�?]/;

// 文字化けした会社名を、既知の会社名のどれかに寄せる。
// 判定は「化けていない文字だけで一意に決まるか」。決まらなければ触らない。
function fixMojibakeCompany_(v) {
  const s = String(v == null ? '' : v).trim();
  if (!s) return s;
  if (KNOWN_COMPANIES.indexOf(s) >= 0) return s;
  if (!MOJIBAKE_RE.test(s)) return s;              // 化けていないなら触らない
  const pattern = new RegExp('^' + s.split('').map(function (ch) {
    return MOJIBAKE_RE.test(ch) ? '.' : ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }).join('') + '$');
  const hits = KNOWN_COMPANIES.filter(function (c) {
    return c.length === s.length && pattern.test(c);
  });
  return hits.length === 1 ? hits[0] : s;          // 1つに決まるときだけ直す
}

/**
 * 職人マスタの重複行を統合する。
 * ★Codexレビュー[P1]#4: 先勝ちで捨てず、非空の値だけを寄せる。
 *   同じ項目に違う値が入っていたら統合せず conflicts に出して人の判断に回す
 *   （特に単価は給料の元数字なので機械が選んではいけない）。
 * 鍵は (会社, 氏名)。会社が違えば別人。
 * 戻り値: {merged: [[6列], ...], conflicts: [{name, company, field, values}]}
 */
function mergeMemberRows_(rows) {
  const MERGE_FIELDS = [2, 3, 4, 5];               // 事業部 / 単価 / 既定部隊 / 有効
  const FIELD_NAMES = { 2: '事業部', 3: '単価', 4: '既定部隊', 5: '有効' };
  const order = [];
  const byKey = {};
  const conflicts = [];
  const isEmpty = function (v, i) {
    if (v == null || String(v).trim() === '') return true;
    return i === 3 && Number(v) === 0;             // 単価0は「未設定」とみなす
  };

  (rows || []).forEach(function (r) {
    const name = String(r[0] == null ? '' : r[0]).trim();
    if (!name) return;
    const company = String(r[1] == null ? '' : r[1]).trim();
    const key = company + '|' + name;
    if (!byKey[key]) {
      byKey[key] = [name, company, r[2], r[3], r[4], r[5]];
      order.push(key);
      return;
    }
    const cur = byKey[key];
    MERGE_FIELDS.forEach(function (i) {
      const a = cur[i], b = r[i];
      if (isEmpty(b, i)) return;                   // 足す値が空なら何もしない
      if (isEmpty(a, i)) { cur[i] = b; return; }   // 今が空なら埋める
      if (String(a).trim() !== String(b).trim()) { // 両方に値があって食い違う
        conflicts.push({ name: name, company: company, field: FIELD_NAMES[i],
                         values: [String(a), String(b)] });
      }
    });
  });
  return { merged: order.map(function (k) { return byKey[k]; }), conflicts: conflicts };
}

/**
 * Apps Scriptのエディタから手で実行する。
 *   cleanupMastersPhase1()      … dry-run。何も書かずに結果だけログに出す
 *   cleanupMastersPhase1(true)  … 実行
 *
 * (1) 職人マスタの重複を統合（食い違いがあれば中止）
 * (2) 予定に出るのにマスタに無い人を (会社,氏名) で判定して追加
 * (3) 日報データの文字化けした会社名を直す
 * (4) 「予定が0件の人」の一覧を出す（★無効化はしない。人がTask 6Bの画面で決める）
 * 行数は1行も増減しない。
 */
function cleanupMastersPhase1(apply) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('他の処理が動いています。少し待ってからもう一度実行してください。');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const report = { dryRun: !apply, 重複統合: 0, 追加: 0, 会社名修正: 0,
                     食い違い: [], 予定0件の候補: [] };

    // (1) 重複の統合
    const mSheet = getOrCreateMemberSheet_(ss);          // ここで6列に拡張される
    const mData = mSheet.getDataRange().getValues();
    const body = mData.slice(1).map(function (r) { return r.slice(0, 6); });
    const beforeCount = body.filter(function (r) { return String(r[0] || '').trim(); }).length;
    const m = mergeMemberRows_(body);
    report.重複統合 = beforeCount - m.merged.length;
    report.食い違い = m.conflicts;
    if (m.conflicts.length) {
      // ★勝手に決めない。人が直してからもう一度実行してもらう。
      Logger.log('中止: 値が食い違う重複があります\n' + JSON.stringify(m.conflicts, null, 2));
      return report;
    }

    // (2) 予定に出るのにマスタに無い人（★会社も見る）
    const nSheet = ss.getSheetByName(SHEET_NAME);
    const nData = nSheet.getDataRange().getValues();
    const nameIdx = HEADERS.indexOf('氏名');
    const compIdx = HEADERS.indexOf('会社');
    const keep = m.merged.slice();
    const known = {};
    keep.forEach(function (r) { known[String(r[1]).trim() + '|' + String(r[0]).trim()] = true; });
    const seenInNippo = {};
    for (let i = 1; i < nData.length; i++) {
      const nm = String(nData[i][nameIdx] || '').trim();
      if (!nm) continue;
      const co = fixMojibakeCompany_(nData[i][compIdx]);
      seenInNippo[nm] = true;
      const key = co + '|' + nm;
      if (known[key]) continue;
      known[key] = true;
      keep.push([nm, co, '', 0, '', '○']);            // ★既定は必ず有効
      report.追加++;
    }

    // (4) 「予定が0件」の候補を出すだけ（無効化はしない）
    keep.forEach(function (r) {
      if (!seenInNippo[String(r[0]).trim()]) {
        report.予定0件の候補.push(String(r[0]).trim() + '（' + String(r[1]).trim() + '）');
      }
    });

    // (3) 日報データの会社名の文字化け
    let changed = 0;
    const col = [];
    for (let i = 1; i < nData.length; i++) {
      const raw = String(nData[i][compIdx] == null ? '' : nData[i][compIdx]);
      const fixed = fixMojibakeCompany_(raw);
      if (fixed !== raw.trim()) changed++;
      col.push([fixed]);
    }
    report.会社名修正 = changed;

    if (!apply) {
      Logger.log('【dry-run】書き込みはしていません\n' + JSON.stringify(report, null, 2));
      return report;
    }

    // 書き込み（★上書き→余りを消す順。逆にすると途中で落ちたとき空になる）
    const oldLastRow = mSheet.getLastRow();
    if (keep.length) mSheet.getRange(2, 1, keep.length, 6).setValues(keep);
    const extraRows = oldLastRow - 1 - keep.length;
    if (extraRows > 0) mSheet.getRange(2 + keep.length, 1, extraRows, 6).clearContent();
    if (changed > 0 && col.length) {
      nSheet.getRange(2, compIdx + 1, col.length, 1).setValues(col);
    }
    SpreadsheetApp.flush();

    logOperation_(ss, 'cleanup_masters_phase1', '職人マスタ/日報データ',
      JSON.stringify({ 重複統合: report.重複統合, 追加: report.追加, 会社名修正: report.会社名修正 }), 'system');
    Logger.log(JSON.stringify(report, null, 2));
    return report;
  } finally {
    lock.releaseLock();
  }
}
```

`EXPORT_SNIPPET` の `fixMojibakeCompany_, looksLikeNonPerson_` を
`fixMojibakeCompany_, mergeMemberRows_` に差し替える（`looksLikeNonPerson_` は作らない）。

- [ ] **Step 4: テストを実行して通ることを確認する**

```bash
cd cf && npx vitest run
```

- [ ] **Step 5: コミット**

```bash
git add gas.js cf/test/gas-phase1.test.js
git commit -m "feat(gas): データ掃除（dry-run必須・重複は値を寄せて統合・推測で無効化しない）"
```

---

### Task 5B: 部隊の割り当ては「作らない」— ★利用者が決めるまで空のままにする

**★Codexレビュー[P2]#6 で判明した、設計書の穴。**

設計書 §5 の手順3に「既定部隊の流し込み」と書いたが、**流し込む中身が存在しない。**
誰が1部隊で誰が2部隊なのかは、**社長・利用者しか知らない情報**であり、
コードからも実データからも導けない。

**推測して埋めてはいけない理由**:
- 間違った部隊が2,664行に書き込まれると、正しい割当が決まったあと
  「どれが推測でどれが本物か」を区別できなくなる
- 拠点のときは「GRミツマ＝関東支店」という**運用実態と一致する明確な根拠**があったが、
  部隊にはそれが無い

**したがってフェーズ1では:**

| | やること |
|---|---|
| 部隊の仕組み | ✅ 作る（21列目・入力欄・表示・既定部隊の自動補完） |
| 既定部隊の中身 | ❌ **空のまま出す。** Task 6B の画面で利用者が1人ずつ設定する |
| 既存2,664行の部隊 | ❌ **空のまま。** 割当表が決まってから別途バックフィルする |

- [ ] **Step 1: 設計書の §5 手順3 から「既定部隊の流し込み」を削る**

`docs/superpowers/specs/2026-08-27-zensha-jinin-haichi-design.md` の
反映の順番の表を、実態に合わせて直す（掃除だけにする）。

- [ ] **Step 2: リリース確認の文言を直す**

Task 10 Step 11 の「責任者を選ぶと部隊が自動で入る」は、
**先に Task 6B の画面で誰か1人に既定部隊を設定してから**でないと確認できない。
確認手順にその1行を足す。

- [ ] **Step 3: 利用者に割当表を尋ねる（フェーズ1の最後・リリース後）**

引き継ぎ書に「**部隊の割当表が未確定**。誰が1〜4部隊かを決めてもらう必要がある」と明記する。
将来バックフィルするときの注意も残す:
- 鍵は氏名だけでなく **(会社, 氏名)**
- 対象は現場系の予定だけ（事務所・休みに部隊は不要）
- **日報データとアーカイブの両方**に流す
- 必ず dry-run を先に出す

---

### Task 6: 画面 — 部隊の入力欄と表示（index.html / admin.html）

拠点で確立した方式をそのまま横に伸ばす。`s-` は新規登録、`e-` は編集モーダル。

**Files:**
- Modify: `index.html`（拠点の欄 `s-kyoten-row` / `e-kyoten-row` の直後に部隊の欄）
- Modify: `index.html:1848` 付近（`kyotenTag` の下に `butaiTag`）
- Modify: `index.html:2264-2266`（送信する member に `butai` を載せる）
- Modify: `admin.html`（同じ変更）

**Interfaces:**
- Consumes: Task 2 の doGet member `{name, company, division, butai, active}`
- Consumes: Task 2 の 21列目 `部隊`
- Produces: `readButai(p) -> string` / `refreshButaiField(p, opts)` / `butaiTag(v) -> html`

- [ ] **Step 1: 現在の拠点の実装を読む（真似る対象を確認する）**

```bash
sed -n '355,365p' index.html; sed -n '1840,1860p' index.html; sed -n '2040,2070p' index.html
```

- [ ] **Step 2: 部隊の欄をHTMLに足す**

`index.html` の `s-kyoten-row` のブロックの**直後**に:

```html
  <div id="s-butai-row" style="margin-bottom:12px">
    <label>部隊</label>
    <select id="s-butai" onchange="onButaiChanged('s')" style="margin-bottom:0">
      <option value="">部隊なし</option>
      <option value="1部隊">1部隊</option>
      <option value="2部隊">2部隊</option>
      <option value="3部隊">3部隊</option>
      <option value="4部隊">4部隊</option>
    </select>
  </div>
```

`e-kyoten-row` の直後に、同じものを `e-` 接頭辞で足す。`admin.html` にも同じ2箇所。

- [ ] **Step 3: JSを足す**

`index.html` の `kyotenTouched` の定義の近くに:

```js
// ★2026-08-27 フェーズ1: 部隊。拠点と同じ「手で変えた値は上書きしない」方式。
const BUTAI_VALUES=['1部隊','2部隊','3部隊','4部隊'];
const butaiTouched={s:false,e:false};
function onButaiChanged(p){butaiTouched[p]=true;}
function normalizeButai(v){const s=String(v==null?'':v).trim();return BUTAI_VALUES.includes(s)?s:'';}
function memberDefaultButai(name){
  const m=(allMembers||[]).find(x=>String(x.name||'').trim()===String(name||'').trim());
  return m?normalizeButai(m.butai):'';
}
// 代表者を選んだら、その人の既定部隊を初期値として入れる（手で変えた後は上書きしない）
function refreshButaiField(p,opts){
  const force=opts&&opts.force;
  const el=document.getElementById(p+'-butai');
  if(!el)return;
  if(butaiTouched[p]&&!force)return;
  const leader=document.getElementById(p+'-leader');
  el.value=memberDefaultButai(leader?leader.value:'');
}
function readButai(p){
  const el=document.getElementById(p+'-butai');
  return el?normalizeButai(el.value):'';
}
function butaiTag(v){
  const b=normalizeButai(v);
  if(!b)return '';
  return `<span class="butai-tag">${esc(b)}</span>`;
}
```

CSS（拠点タグの定義の近くに）:

```css
.butai-tag{display:inline-block;font-size:10px;padding:1px 5px;border-radius:3px;background:#EDE7F6;color:#5E35B1;border:1px solid #B39DDB;margin-left:3px}
```

- [ ] **Step 3B: ★部隊を画面のデータ構造に通す（これが無いと編集で部隊が消える）**

**★Codexレビュー[P1]#1。実コードで確認済み（2026-08-27）。**
これを飛ばすと **「部隊と無関係な編集をしただけで、保存済みの部隊が消える」**
という最悪の壊れ方をする。拠点で起きたバグ（[P1]#2）とまったく同じ形。

原因は2つ:

1. **`parseRows`（`index.html:2496` / `admin.html:2255`）は列を1つずつ手で写している。**
   最後が `kyoten:String(r['拠点']||'')` で終わっており、**部隊は捨てられる**
2. **`groupNippos`（`index.html:2897` / `admin.html:3007`）の members は
   `{name, role}` の2つだけ。** 計画が読もうとしていた `g.members[0].butai` は**必ず undefined**

しかも保存時は `readButai('e')` が空文字を送るため、`resolveButai_` の仕様
（画面が項目を送ってきたら空欄も尊重する）により**既定部隊でも補完されず、消える**。

**両ファイルの4箇所すべてを直す:**

`parseRows` の `kyoten:` の行の後ろに追加（`index.html:2511` / `admin.html` の同じ場所）:

```js
    kyoten:String(r['拠点']||''),     // ★2026-08-26 拠点（本社/関東支店/両方）
    butai:String(r['部隊']||'')       // ★2026-08-27 部隊（1〜4部隊）
```

`groupNippos` の `groups[key]={...}` に `butai` を足し、members にも持たせる:

```js
    if(!groups[key])groups[key]={date:n.date, /* …既存のまま… */ ,
      vehicle:n.vehicle||'', butai:n.butai||'',      // ★2026-08-27 部隊
      isGhost:!!n.isGhost,isPending:!!n.isPending,originalId:n.originalId||''};
    groups[key].members.push({name:n.name,role:n.role,butai:n.butai||''});
```

⚠️ **`groups[key]` の作り方に注意**: グループの鍵に部隊は入れない。
同じ予定（同じID）の全行は同じ部隊であるべきなので、グループ単位で1つ持てばよい。
鍵に足すと、部隊だけ違う行がグループとして割れてカレンダーの表示が崩れる。

- [ ] **Step 3C: ★一括編集の入口も塞ぐ（Codexが見つけたもう1つの穴）**

編集の入口は単一編集だけではない。**一括編集は別の関数**
（`index.html:3728` / `admin.html:3687`）で、当初の計画は単一編集しか見ていなかった。

```bash
grep -n "bulkEditMode\|function openBulkEdit\|function applyBulkEdit" index.html admin.html | head
```

一括編集を開くときも `butaiTouched.e=false` にし、**選択した予定の部隊が全部同じときだけ
その値を初期表示する**。バラバラなら空にして「変更しない」を意味させる:

```js
  // ★2026-08-27: 一括編集では、選んだ予定の部隊が全部同じときだけ初期値に入れる。
  //   直前に開いた単一編集の部隊が残っていると、無関係な複数予定へ誤って適用される。
  butaiTouched.e=false;
  const bs=[...new Set(targetGroups.map(g=>String(g.butai||'')))];
  const eb=document.getElementById('e-butai');
  if(eb)eb.value=(bs.length===1?bs[0]:'');
```

- [ ] **Step 4: 送信と読み戻しを繋ぐ**

`index.html:2264-2266`（`const members=[{name:leader,role:'代表'}];` の周辺）で、
組み立てた各 member に `butai` を載せる:

```js
  const butai=readButai('s');
  const members=[{name:leader,role:'代表',butai:butai}];
  selectedMembers.forEach(i=>{if(shokunin[i]!==leader)members.push({name:shokunin[i],role:'同行',butai:butai});});
```

編集モーダルでも同様に `readButai('e')` を載せる。
編集モーダルを開く関数（`openEditModal` 相当・`kyotenTouched.e=false` を書いている `index.html:3463` 付近）に:

```js
  // ★Step 3B でグループに butai を持たせたので、そこから読む
  butaiTouched.e=false;
  const eb=document.getElementById('e-butai');
  if(eb)eb.value=normalizeButai(g.butai);
```

新規フォームのクリア処理（`index.html:2332` の `kyotenTouched.s=false;` の隣）に:

```js
    butaiTouched.s=false;refreshButaiField('s',{force:true});
```

代表者の `<select>` の `onchange`（`index.html:398` の `checkLocationJobNo()`）に `refreshButaiField('s')` を足す。

- [ ] **Step 5: ★起動時キャッシュにも既定部隊を残す（見落とすと初回だけ動く不具合になる）**

`index.html:2403`（`saveSnapshot`）と `admin.html:2163` は、端末に残す職人データを
**`name` / `company` / `division` の3つに絞り込んでいる**（単価を意図的に捨てるため）。
Worker側の `sanitizeForStorage` とまったく同じ罠。ここに書き足さないと、
**2回目以降の起動（キャッシュから描くとき）だけ既定部隊が空になる**。
再現しにくく、原因も分かりにくい種類の不具合になる。

両ファイルの該当行を置き換える:

```js
      // ★2026-08-27 フェーズ1: 既定部隊と有効も端末に残す（無いとキャッシュ起動時だけ部隊が入らない）。
      //   単価(rate)は引き続き残さない（給料情報のため意図的に捨てている）。
      members:(json.members||[]).map(m=>({name:m.name,company:m.company,division:m.division,butai:m.butai,active:m.active!==false})),
```

- [ ] **Step 6: カレンダーに部隊の印を出す**

`index.html:3046` 付近（`memberTags` を組み立てている箇所）と `index.html:3140-3142`（詳細表示）で、
拠点タグを出している場所の隣に `butaiTag(...)` を足す。

- [ ] **Step 7: admin.html に同じ変更を入れる**

```bash
grep -n "s-kyoten-row\|e-kyoten-row\|kyotenTouched\|function kyotenTag\|role:'代表'" admin.html
```

同じ位置に同じものを入れる。

- [ ] **Step 8: 構文チェック**

```bash
node -e "const s=require('fs').readFileSync('index.html','utf8');const m=[...s.matchAll(/<script(?![^>]*src=)[^>]*>([\s\S]*?)<\/script>/g)];m.forEach((x,i)=>{try{new Function(x[1])}catch(e){console.log('index.html block',i,e.message)}});console.log('index.html blocks:',m.length)"
```

`admin.html` にも同じチェックを実行する。エラーが出なければ OK。

- [ ] **Step 9: コミット**

```bash
git add index.html admin.html
git commit -m "feat(画面): 予定に部隊（1〜4部隊）を追加。代表者から既定部隊が自動で入る"
```

---

### Task 6B: 職人管理モーダルで「既定部隊」と「有効」を編集できるようにする

**★これが無いと Task 6 の「自動で入る」が一生発動しない。**
既定部隊はスプレッドシートを直接触るしか設定手段が無くなり、
利用者（非エンジニア）が使えない。現場マスタの拠点で同じ状態を作ってしまった反省。

`admin.html:810` に既存の**職人管理モーダル**があり、事業部と単価を編集できる。
`update_member_division`（`gas.js:805`）とまったく同じ形で2つ足すだけ。

**Files:**
- Modify: `gas.js`（`update_member_division` の隣にアクションを2つ追加）
- Modify: `admin.html:4789` 以降（職人管理の行に列を2つ追加）
- Test: `cf/test/gas-phase1.test.js`（追記）

**Interfaces:**
- Consumes: Task 2 の `normalizeButai_` / 職人マスタ6列
- Produces: GASアクション `update_member_butai`（`{name, company, butai}`）
- Produces: GASアクション `update_member_active`（`{name, company, active}`）

- [ ] **Step 1: 失敗するテストを書く**

`EXPORT_SNIPPET` に `normalizeMemberActive_` を足したうえで、`cf/test/gas-phase1.test.js` に追記:

```js
describe('職人の有効/無効', () => {
  it('×だけが無効。それ以外は全部有効', () => {
    expect(ctx.normalizeMemberActive_('×')).toBe(false);
    expect(ctx.normalizeMemberActive_('x')).toBe(false);
    expect(ctx.normalizeMemberActive_('✕')).toBe(false);
    ['○', 'o', '', '　', undefined, null, true].forEach(v =>
      expect(ctx.normalizeMemberActive_(v)).toBe(true));
  });

  it('★空欄は有効（既存71件を巻き込まないための既定）', () => {
    expect(ctx.normalizeMemberActive_('')).toBe(true);
  });
});
```

- [ ] **Step 2: テストを実行して失敗を確認する**

```bash
cd cf && npx vitest run test/gas-phase1.test.js
```

- [ ] **Step 3: gas.js に実装する**

`normalizeButai_` の下に追加:

```js
// 「×」だけを無効とみなす。空欄・未記入は有効（既存の職人を巻き込まないため）。
function normalizeMemberActive_(v) {
  const s = String(v == null ? '' : v).trim();
  return !(s === '×' || s === 'x' || s === 'X' || s === '✕');
}
```

Task 2 で書いた doGet の `active:` を、この関数を使う形に揃える:

```js
      active: normalizeMemberActive_(r[5])
```

`update_member_division` アクションの直後に2つ追加:

```js
    if (action === 'update_member_butai') {
      const memberSheet = getOrCreateMemberSheet_(ss);
      const name = String(body.name || '').trim();
      const company = String(body.company || '').trim();
      const butai = normalizeButai_(body.butai);   // 知らない値は空になる
      const data = memberSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === name && String(data[i][1]).trim() === company) {
          memberSheet.getRange(i + 1, 5).setValue(butai);
          logOperation_(ss, 'update_member_butai', name + '/' + company, '既定部隊=' + (butai || '(なし)'), updatedBy);
          return ok({updated: name, butai: butai});
        }
      }
      return ok({updated: null});
    }

    if (action === 'update_member_active') {
      const memberSheet = getOrCreateMemberSheet_(ss);
      const name = String(body.name || '').trim();
      const company = String(body.company || '').trim();
      const active = body.active !== false && String(body.active) !== 'false';
      const data = memberSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === name && String(data[i][1]).trim() === company) {
          memberSheet.getRange(i + 1, 6).setValue(active ? '○' : '×');
          logOperation_(ss, 'update_member_active', name + '/' + company, active ? '有効' : '無効', updatedBy);
          return ok({updated: name, active: active});
        }
      }
      return ok({updated: null});
    }
```

- [ ] **Step 4: admin.html の職人管理に列を2つ足す**

`admin.html:4789` からの `// ===== 職人管理 =====` を読み、
職人1人の行を描いている箇所に、事業部のプルダウンと同じ形で足す:

```js
const BUTAI_OPTIONS=['','1部隊','2部隊','3部隊','4部隊'];
function memberButaiSelect(name,cur){
  return `<select onchange="setMemberButai('${esc(name)}',this.value)">`+
    BUTAI_OPTIONS.map(b=>`<option value="${b}"${b===(cur||'')?' selected':''}>${b||'（なし）'}</option>`).join('')+
    `</select>`;
}
async function setMemberButai(name,butai){
  try{
    const res=await fetch(GAS_URL,{method:'POST',body:JSON.stringify({action:'update_member_butai',name,company:currentCompany,butai}),headers:{'Content-Type':'text/plain'}});
    const json=await res.json();
    if(json.status!=='ok')throw new Error(json.message||'保存できませんでした');
    // ★Codexレビュー[P2]#8: 実物のキャッシュ名は settingsMembers（admin.html:4794）。
    //   memberCache という変数は存在しない。捨てるのではなく該当要素を更新する。
    const sm=(settingsMembers||[]).find(x=>x.name===name&&x.company===currentCompany);
    if(sm)sm.butai=butai;
    const am=(allMembers||[]).find(x=>x.name===name&&x.company===currentCompany);
    if(am)am.butai=butai;
    showAlert(`${name} の既定部隊を「${butai||'なし'}」にしました`,'ok');
  }catch(e){showAlert('保存できませんでした','err');}
}
async function setMemberActive(name,active){
  try{
    const res=await fetch(GAS_URL,{method:'POST',body:JSON.stringify({action:'update_member_active',name,company:currentCompany,active}),headers:{'Content-Type':'text/plain'}});
    const json=await res.json();
    if(json.status!=='ok')throw new Error(json.message||'保存できませんでした');
    const sm=(settingsMembers||[]).find(x=>x.name===name&&x.company===currentCompany);
    if(sm)sm.active=active;
    const am=(allMembers||[]).find(x=>x.name===name&&x.company===currentCompany);
    if(am)am.active=active;
    showAlert(`${name} を${active?'有効':'無効'}にしました`,'ok');
  }catch(e){showAlert('保存できませんでした','err');}
}
```

有効の切替はチェックボックスにする:

```js
`<label><input type="checkbox" ${active?'checked':''} onchange="setMemberActive('${esc(name)}',this.checked)"> 有効</label>`
```

- [ ] **Step 4B: ★loadSettingsMembers に新しい2列を読ませる（Codexレビュー[P2]#8）**

`loadSettingsMembers`（`admin.html:4814-4821`）は
**氏名・単価・会社・事業部の4つしか読んでいない**。ここに足さないと、
画面は既定部隊・有効の**現在値を表示できず、常に空から上書き**してしまう。

```js
      const header=data[0]||[];
      const nameIdx=header.indexOf('氏名'),rateIdx=header.indexOf('単価');
      if(nameIdx<0||rateIdx<0){
        throw new Error('職人マスタのヘッダに「氏名」または「単価」列が見つかりません: '+JSON.stringify(header));
      }
      const companyIdx=header.indexOf('会社'),divIdx=header.indexOf('事業部');
      // ★2026-08-27 フェーズ1: 既定部隊・有効。まだ列が無い場合は -1 になるので
      //   その時は既定値（空／有効）で扱う＝GASを出す前でも画面が壊れない。
      const butaiIdx=header.indexOf('既定部隊'),activeIdx=header.indexOf('有効');
```

`settingsMembers=data.slice(1).map(r=>{...})` が返すオブジェクトに2つ足す:

```js
          butai: butaiIdx>=0?String(r[butaiIdx]||'').trim():'',
          active: activeIdx>=0?String(r[activeIdx]||'').trim()!=='×':true,
```

⚠️ **`throw` するのは氏名・単価が無いときだけ**にする。既定部隊・有効を必須にすると、
**Worker/GASを出す前に画面を開いた瞬間に職人管理が死ぬ**（反映順の事故）。

- [ ] **Step 5: 構文チェックとテスト**

Task 6 Step 8 と同じコマンド＋ `cd cf && npx vitest run`。

- [ ] **Step 6: コミット**

```bash
git add gas.js admin.html cf/test/gas-phase1.test.js
git commit -m "feat(admin): 職人管理で既定部隊と有効/無効を編集できるようにする"
```

---

### Task 7: 画面 — 役割の呼称を「責任者／班員」に変える（表示だけ）

**保存値 `代表` / `同行` は1文字も変えない。** 変えるのは利用者の目に見える文言だけ。

**Files:**
- Modify: `index.html`（表示ラベル9箇所前後）
- Modify: `admin.html`（同じ）

**Interfaces:**
- Consumes: なし
- Produces: なし（表示のみ。`=== '代表'` の比較は1つも触らない）

- [ ] **Step 1: 変える対象と変えない対象を仕分ける**

```bash
grep -n "代表\|同行" index.html | grep -v "==='代表'\|==='同行'\|role:'代表'\|role:'同行'\|role==='代表'\|role==='同行'"
```

出てきた行が**表示ラベル**＝変える対象。
`role:'代表'` `m.role==='代表'` のような**コード上の比較と代入は変えない**。

- [ ] **Step 2: 表示ラベルだけを置き換える**

対象（`index.html` / `admin.html` の両方）:

| 今 | 変更後 |
|---|---|
| `<label>代表者<span...` | `<label>責任者<span...` |
| `<option value="">代表者を選択してください</option>` | `<option value="">責任者を選択してください</option>` |
| `同行メンバー（タップで選択）` | `班員（タップで選択）` |
| `<b>代表者</b>を選ぶ（その日の現場責任者）` | `<b>責任者</b>を選ぶ（その日の現場責任者）` |
| `<b>同行者</b>（代表以外のメンバー）を選ぶ` | `<b>班員</b>（責任者以外のメンバー）を選ぶ` |
| `（代表）`（詳細表示 `index.html:3142`） | `（責任者）` |
| `デフォルトは<b>代表者の事業部</b>` | `デフォルトは<b>責任者の事業部</b>` |
| `初日だけ他事業部の人が代表する場合` | `初日だけ他事業部の人が責任者になる場合` |
| `空きは「○」、使用中は<b>代表者名と行先</b>` | `空きは「○」、使用中は<b>責任者名と行先</b>` |
| `操作者・日付・代表者・同行メンバーは残ります` | `操作者・日付・責任者・班員は残ります` |
| `時間・元請・代表者・メンバー・車両等` | `時間・元請・責任者・メンバー・車両等` |
| `代表者は休み/倉庫モードでも選択可能`（コメント） | `責任者は休み/倉庫モードでも選択可能` |
| `if(!leader){showAlert('代表者を選択してください','err')` | `...showAlert('責任者を選択してください','err')` |

- [ ] **Step 3: 保存値を1つも変えていないことを確認する**

```bash
grep -o "role:'代表'\|role:'同行'\|role==='代表'\|role==='同行'\|role === '代表'\|=== '代表'\|=== '同行'\|==='代表'\|==='同行'" gas.js index.html admin.html | wc -l
```

想定: **39**（2026-08-27 実測の基準値。内訳: gas.js 2 / index.html 19 / admin.html 18）。
**1つでも減っていたら保存値を壊している。** 増えるのも想定外なので調べ直すこと。

- [ ] **Step 4: 構文チェック**

Task 6 Step 8 と同じコマンドを両ファイルに実行する。

- [ ] **Step 5: 自動テストを流して壊していないことを確認する**

```bash
cd cf && npx vitest run
```

- [ ] **Step 6: コミット**

```bash
git add index.html admin.html
git commit -m "feat(画面): 役割の呼称を責任者／班員へ（表示のみ。保存値の代表/同行は不変）"
```

---

### Task 8: 画面 — 現場ごとの案件ステータス（8段階）

**Files:**
- Modify: `admin.html`（現場管理 `screen-genba` の現場一覧）
- Modify: `index.html`（現場管理 `screen-genba`。表示のみ・編集はadminに寄せる）

**Interfaces:**
- Consumes: Task 3 の jobsite `{genba, loc, jobNo, completed, billingMethod, kyoten, status}`
- Consumes: Task 3 のGASアクション `set_site_status`（`{genba, loc, status}` を送る）

- [ ] **Step 1: 現場一覧の描画箇所を見つける**

```bash
grep -n "billingMethod\|completed" admin.html | head -20
```

- [ ] **Step 2: ステータスのプルダウンを足す（admin.html）**

現場1件を描画している箇所に:

```html
<select class="site-status" data-genba="${esc(j.genba)}" data-loc="${esc(j.loc)}"
        onchange="setSiteStatus(this)">
  ${['見積中','受注','準備中','施工中','残工事','完工','延期','中止']
    .map(s=>`<option value="${s}"${(j.status||'施工中')===s?' selected':''}>${s}</option>`).join('')}
</select>
```

JS:

```js
const SITE_STATUSES=['見積中','受注','準備中','施工中','残工事','完工','延期','中止'];
async function setSiteStatus(el){
  const genba=el.dataset.genba,loc=el.dataset.loc,status=el.value;
  const prev=el.dataset.prev||'';
  el.disabled=true;
  try{
    const r=await postGas({action:'set_site_status',genba,loc,status});
    if(r&&r.status==='ok'){el.dataset.prev=status;showAlert(`${loc||genba} を「${status}」にしました`,'ok');loadData();}
    else{el.value=prev;showAlert('変更できませんでした','err');}
  }catch(e){el.value=prev;showAlert('変更できませんでした','err');}
  finally{el.disabled=false;}
}
```

`postGas` は既存の送信関数に合わせる（`grep -n "function postGas\|function callGas" admin.html` で確認）。

- [ ] **Step 3: ★index.html の「完了にする／進行中に戻す」ボタンを撤去する**

**★Codexレビュー[P2]#7。実コードで確認済み（`index.html:3214-3224`）。**

職人用の予定詳細に「完了にする／進行中に戻す」ボタンが残っており、
`update_site_status`（`completed` の真偽値だけを送る）を呼んでいる。
これを放置すると:

> 管理者が現場を **「延期」** にする
>   → 職人が予定詳細で「完了にする」を押す
>     → `完了=TRUE` になり、次に読んだとき `normalizeSiteStatus_` が
>        **「延期」を「完工」に書き換える**

つまり **管理者の設定が職人の操作で勝手に消える**。
計画に書いた「indexは表示のみ」とも矛盾していた。

やること:
1. `index.html` の完了/進行中ボタンとその呼び出し関数を**削除**する
   （現場一覧・予定詳細の両方。`grep -n "完了にする\|進行中に戻す" index.html` で全部出す）
2. 代わりに**ステータスの文字を表示するだけ**にする
3. `gas.js` の旧 `update_site_status` は、**ステータス列に値が入っていたら上書きを拒否**する

```js
      // ★2026-08-27 [P2]#7: 8段階ステータスが設定済みなら、旧APIで上書きさせない。
      //   （延期・中止が「完工」「施工中」に化けるのを防ぐ）
      const curStatus = String(data[i][11] || '').trim();
      if (SITE_STATUSES.indexOf(curStatus) >= 0) {
        return error('この現場はステータス「' + curStatus + '」が設定されています。変更は管理者画面から行ってください。');
      }
```

4. 管理者は `set_site_status` に一本化する（Step 2 で作ったもの）

⚠️ **ステータス変更も履歴に残す**: `set_site_status` の中で、
変更前のステータスを `logHistory_` に `field:'案件ステータス'` として記録する。
現場マスタは日報データと別なので `旧ID/新ID` は空でよく、`項目` に現場名を含める。

- [ ] **Step 4: 構文チェック**

Task 6 Step 8 と同じコマンド。

- [ ] **Step 5: コミット**

```bash
git add index.html admin.html
git commit -m "feat(画面): 現場ごとに案件ステータス8段階を設定できるようにする"
```

---

### Task 9: 画面 — admin.html に「履歴」タブを足す

**Files:**
- Modify: `gas.js`（`doPost` に `get_history` アクション）
- Modify: `admin.html`（タブ1つ＋画面1つ）
- Test: `cf/test/gas-phase1.test.js`（追記）

**Interfaces:**
- Consumes: Task 4 の `変更履歴` シート（8列）
- Produces: GASアクション `get_history` → `{status:'ok', rows:[[日時,操作,旧ID,新ID,項目,変更前,変更後,実行者], ...]}`（新しい順・最大500件）

- [ ] **Step 1: 失敗するテストを書く**

```js
describe('履歴の取り出し', () => {
  it('新しい順に並べ替える', () => {
    const rows = [
      ['2026/08/25 10:00','update','A','B','メモ','','あ','向'],
      ['2026/08/27 09:00','update','C','D','メモ','','い','元'],
      ['2026/08/26 12:00','add','','E','(新規)','','う','中島']
    ];
    const out = ctx.sortHistoryRows_(rows);
    expect(out[0][0]).toBe('2026/08/27 09:00');
    expect(out[2][0]).toBe('2026/08/25 10:00');
  });

  it('件数の上限で切る', () => {
    const rows = Array.from({length: 700}, (_, i) =>
      ['2026/08/' + String((i % 28) + 1).padStart(2,'0') + ' 10:00','update','','','メモ','','x','向']);
    expect(ctx.sortHistoryRows_(rows, 500).length).toBe(500);
  });

  it('空でも落ちない', () => {
    expect(ctx.sortHistoryRows_([])).toEqual([]);
    expect(ctx.sortHistoryRows_(null)).toEqual([]);
  });
});
```

- [ ] **Step 2: テストを実行して失敗を確認する**

```bash
cd cf && npx vitest run test/gas-phase1.test.js
```

- [ ] **Step 3: 実装（gas.js）**

`logHistory_` の下に:

```js
const HISTORY_MAX_ROWS = 500;

// 日時は 'YYYY/M/D H:mm:ss' の文字列で保存されている。数値に直して新しい順に並べる。
function historyTimeValue_(v) {
  const s = String(v == null ? '' : v).trim();
  const m = /^(\d{4})\/(\d{1,2})\/(\d{1,2})[ 　]+(\d{1,2}):(\d{2})(?::(\d{2}))?/.exec(s);
  if (!m) return 0;
  return new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0)).getTime();
}

function sortHistoryRows_(rows, limit) {
  const arr = (rows || []).slice();
  arr.sort((a, b) => historyTimeValue_(b[0]) - historyTimeValue_(a[0]));
  return arr.slice(0, limit || HISTORY_MAX_ROWS);
}
```

`doPost` のアクションの並びに:

```js
    if (action === 'get_history') {
      const sheet = getOrCreateHistorySheet_(ss);
      const data = sheet.getDataRange().getValues();
      const body2 = data.length > 1 ? data.slice(1) : [];
      return ok({ rows: sortHistoryRows_(body2, Number(body.limit) || HISTORY_MAX_ROWS) });
    }
```

- [ ] **Step 4: 実装（admin.html）**

タブバーに1つ足す:

```html
<button class="tab" onclick="switchTab('history')"><span class="tab-icon">🕘</span>履歴</button>
```

画面:

```html
<div id="screen-history" class="screen">
<div class="page-title">変更履歴</div>
<div class="card" style="overflow-x:auto">
  <div id="history-body" style="min-width:640px">読み込み中…</div>
</div>
<button class="btn btn-secondary" onclick="loadHistory()">更新</button>
</div>
```

JS:

**★Codexレビュー[P1]#3 の後半**: 平坦な明細表だけでは
設計書の「**予定を選ぶと変更前が見える**」を満たさない。
**1回の編集ごとにまとめて表示し、削除は全項目を開けるようにする。**

```js
// 履歴を「1回の操作」単位にまとめる。
// 同じ日時・同じ操作・同じ実行者の明細は1つの編集としてくくる。
function groupHistory(rows){
  const groups=[];const idx={};
  rows.forEach(x=>{
    const key=x[0]+'|'+x[1]+'|'+x[7];
    if(idx[key]===undefined){idx[key]=groups.length;groups.push({at:x[0],action:x[1],by:x[7],items:[]});}
    groups[idx[key]].items.push({oldId:x[2],newId:x[3],field:x[4],before:x[5],after:x[6]});
  });
  return groups;
}

const ACTION_LABEL={add:'登録',update:'変更',delete:'削除'};

// 削除の「変更前」は21列すべてのJSON。人が読める表に開く。
function renderDeletedRecord(json){
  let o;try{o=JSON.parse(json);}catch(e){return `<div>${esc(String(json))}</div>`;}
  const skip=['登録日時','色'];
  return '<table class="tbl" style="font-size:12px">'+
    Object.keys(o).filter(k=>!skip.includes(k)&&String(o[k]).trim()!=='')
      .map(k=>`<tr><th style="text-align:left;white-space:nowrap">${esc(k)}</th><td>${esc(String(o[k]))}</td></tr>`).join('')+
    '</table>';
}

async function loadHistory(){
  const el=document.getElementById('history-body');
  el.textContent='読み込み中…';
  try{
    const r=await postGas({action:'get_history',limit:500});
    if(!r||r.status!=='ok'){el.textContent='読み込めませんでした';return;}
    if(!r.rows.length){el.textContent='まだ履歴はありません';return;}
    el.innerHTML=groupHistory(r.rows).map((g,gi)=>{
      const isDel=g.action==='delete';
      const body=isDel
        ? g.items.map(it=>
            `<details><summary>削除された予定を見る（ID: ${esc(it.oldId)}）</summary>${renderDeletedRecord(it.before)}</details>`
          ).join('')
        : '<table class="tbl" style="font-size:12px"><thead><tr><th>項目</th><th>変更前</th><th>変更後</th></tr></thead><tbody>'+
          g.items.map(it=>`<tr><td>${esc(it.field)}</td>`+
            `<td style="color:#c0392b">${esc(it.before)}</td>`+
            `<td style="color:#27ae60">${esc(it.after)}</td></tr>`).join('')+
          '</tbody></table>';
      // 旧ID→新ID が分かるように出す（編集するとIDが変わるため）
      const ids=[...new Set(g.items.filter(it=>it.oldId&&it.newId&&it.oldId!==it.newId)
        .map(it=>`${it.oldId} → ${it.newId}`))];
      return `<div class="card" style="margin-bottom:10px">
        <div style="font-weight:600;margin-bottom:6px">${esc(g.at)}　${esc(ACTION_LABEL[g.action]||g.action)}　<span style="color:#666;font-weight:400">${esc(g.by)}</span></div>
        ${ids.length?`<div style="font-size:11px;color:#999;margin-bottom:6px">ID: ${esc(ids.join(' / '))}</div>`:''}
        ${body}
      </div>`;
    }).join('');
  }catch(e){el.textContent='読み込めませんでした';}
}
```

`switchTab` の中で `if(name==='history')loadHistory();` を足す。
`placeKyotenBar()` が `.screen.active` を探すので、**新しい画面にも `page-title` があること**を確認する
（2026-08-26 の拠点バーの作り直しで、見出しの直後に置く実装になっているため）。

- [ ] **Step 5: 構文チェック＋テスト**

```bash
cd cf && npx vitest run
```

- [ ] **Step 6: コミット**

```bash
git add gas.js admin.html cf/test/gas-phase1.test.js
git commit -m "feat(admin): 変更履歴タブ（誰がいつ何を何に変えたかを見る）"
```

---

### Task 10: 本番リリースと実機確認

**★順番を守ること。設計書 §5。画面をGASより先に出すと復元不能。**

- [ ] **Step 1: バックアップ**

Googleスプレッドシートを「ファイル → コピーを作成」で
`予定管理_backup_20260827` として複製する（利用者のChromeで実施）。

- [ ] **Step 2: 全テストを流す**

```bash
cd cf && npx vitest run
```

想定: 全件 PASS。1件でも落ちていたら次へ進まない。

- [ ] **Step 3: Codexレビューを通す（2回目・実装後の差分）**

設計と計画のレビューは**着手前に実施済み**（2026-08-27。結果は
`_local/review/codex_out3_20260827.txt`）。ここでは**書いたコードの差分**をかける。

```bash
git diff main --stat
```

`gas.js` / `cf/src/sync.js` / `index.html` / `admin.html` の差分をレビューにかける。
**指摘が出たら直してから次へ進む**（2026-08-26 は出した後にレビューして5件の欠陥が出た）。

⚠️ Codexに投げるときは **「graphify を使わない・python も pip も実行しない」を必ず先頭に書く**。
書かないと知識グラフの構築に寄り道して時間とAPI費用を浪費する（2026-08-27 に実際に発生）。
サンドボックスは `-s danger-full-access` を付けないとこのPCではファイルを読めない。

- [ ] **Step 4: Worker を出す（無影響）**

```bash
cd cf && npx wrangler deploy
```

```bash
curl -s "https://yotei-api.miscjigyoubu.workers.dev/api/health"
```

想定: `status:"ok"`、`rows` が2,664前後のまま。**この時点でGASはまだ20列なので何も変わらない。**

- [ ] **Step 5: GAS を出す**

Apps Script エディタ（`yotei.glorise@gmail.com` でログイン）で `gas.js` を貼り付け、
**「デプロイを管理」→ 既存のデプロイを編集 → 新バージョン**。
⛔ **「新しいデプロイ」を押さないこと**（URLが変わって全画面が死ぬ）。

- [ ] **Step 6: 21列になったことを確認する**

```bash
curl -s "https://script.google.com/macros/s/AKfycbxp2eUcpIjCj0ZWyAPPD9m3egJrKdWmXRK2AVnFrmBm4iO1QHCk-FZEH5LFFv7OloqcjQ/exec?compact=1" | node -e "let s='';process.stdin.on('data',d=>s+=d).on('end',()=>{const j=JSON.parse(s);console.log('列数',j.headers.length,'/ 20〜21列目',j.headers.slice(19));console.log('行数',j.rows.length)})"
```

想定: 列数21、`['拠点','部隊']`、行数2,664前後。

- [ ] **Step 7: データ掃除を実行する**

**★必ず2回に分ける。**

1. Apps Script エディタで `cleanupMastersPhase1` を**引数なしで実行**（＝dry-run。何も書かない）
2. 実行ログを読む。`{"dryRun":true,"重複統合":9,"追加":1,"会社名修正":1,"食い違い":[],"予定0件の候補":[...]}`
   - **`食い違い` が空でなければここで止める。** 単価などが食い違う重複がある＝人が決める必要がある
   - `予定0件の候補` は**そのまま利用者に見せる**（無効化はしない。誰が実在の人かは機械に分からない）
3. 数字が想定どおりなら、エディタで一時的に `cleanupMastersPhase1(true)` を呼ぶ関数を作って実行

- [ ] **Step 8: 行数が変わっていないことを確認する**

```bash
curl -s "https://script.google.com/macros/s/AKfycbxp2eUcpIjCj0ZWyAPPD9m3egJrKdWmXRK2AVnFrmBm4iO1QHCk-FZEH5LFFv7OloqcjQ/exec?compact=1" | node -e "let s='';process.stdin.on('data',d=>s+=d).on('end',()=>{const j=JSON.parse(s);console.log('行数',j.rows.length,'（2,664前後なら正常）');console.log('職人',j.members.length,'人');const bad=j.rows.filter(r=>/[�?]/.test(String(r[11]||'')));console.log('会社名が化けた行',bad.length,'件（0なら正常）')})"
```

- [ ] **Step 9: D1へ取り込む**

```bash
curl -s -X POST "https://yotei-api.miscjigyoubu.workers.dev/api/sync"
```

```bash
curl -s "https://yotei-api.miscjigyoubu.workers.dev/api/health"
```

想定: `lastSync.ok` が 1、`snapshotAt` が今の時刻に進んでいる。
⛔ **ここが失敗したまま画面を出さないこと**（設計書 §5 の注意）。

- [ ] **Step 10: 画面を出す**

```bash
git push
```

GitHub Pages のビルド完了を待ってから（1〜2分）次へ。

- [ ] **Step 11: Chromeで実機確認（全部やる。1タブで満足しない）**

利用者のChromeで確認する:

1. `index.html` を**再読み込み**（キャッシュが残るため）
2. 会社を **グローライズ / 和信カインド / GRミツマ / GRHD / ラーテル** の順に切り替える
3. 拠点を **全拠点 / 本社 / 関東支店** で切り替える（**昨日の機能を壊していないこと**）
4. タブを **カレンダー / 空き確認 / 現場管理 / 車確認 / 事務** の全部押す
5. `admin.html` の**職人管理**を開き、誰か1人に**既定部隊「2部隊」を設定**する
   （★これをやらないと次が確認できない。フェーズ1は既定部隊が空で出るため）
6. `index.html` の新規登録フォームを開く → **その人を責任者に選ぶと部隊が自動で入る**
7. 部隊を手で「3部隊」に変える → **別の責任者に変えても3部隊のまま**（上書きされない）
8. 1件登録する → カレンダーに**部隊の印**が出る
9. ★**部隊と無関係な項目だけ**（メモ等）を編集して保存 → **部隊が消えていない**ことを確認（[P1]#1）
10. その予定を編集して部隊を変える → 保存
11. `admin.html` の **履歴タブ** に「部隊 1部隊 → 3部隊」が出る
12. `admin.html` の**現場管理**でステータスを「延期」に変える → 再読み込みで残っている
13. ★`index.html` の予定詳細に **「完了にする」ボタンが無い**ことを確認（[P2]#7 の撤去確認）
14. ★予定を1件**削除** → 履歴タブで「削除された予定を見る」を開き**全項目が読める**ことを確認（[P1]#3）
15. 画面幅 **390px**（スマホ）で 1〜4 をもう一度
16. `president.html` を開く → **社長予定が今まで通り出る**（触っていない証明）
17. 登録したテスト予定を**削除して後片付け**する

- [ ] **Step 12: 引き継ぎ書を更新する**

`引き継ぎ.md` の §3 に `### 3.10 部隊・案件ステータス・変更履歴（フェーズ1）` を足す。
含めること: 何を作ったか / 21列目の位置 / **保存値の代表・同行は変えていないこと** /
デプロイの順番 / `cleanupMastersPhase1` は1回きりの関数であること / フェーズ2以降に残したこと。

- [ ] **Step 13: コミット＆プッシュ**

```bash
git add 引き継ぎ.md && git commit -m "docs: フェーズ1（部隊・案件ステータス・変更履歴）の記録" && git push
```

---

## Self-Review

**1. Spec coverage（設計書 §3・§4 の各項目に担当タスクがあるか）**

| 設計書 | 担当 |
|---|---|
| §3.1 日報データ21列目 部隊 | Task 2 |
| §3.2 職人マスタ 既定部隊・有効 | Task 2 |
| §3.3 現場マスタ ステータス8段階 | Task 3 |
| §3.4 変更履歴シート | Task 4 |
| §3.5 役割の呼称（表示だけ） | Task 7 |
| §3.6 データ掃除 | Task 5 |
| §4 画面（カレンダー・登録編集・現場管理・履歴） | Task 6, 8, 9 |
| §5 反映の順番 | Task 1（先出し）＋ Task 10 |
| §7 完了条件 | Task 10 Step 2/8/9/11 |

**2. Placeholder scan** — 「あとで」「適切に」「TBD」なし。全ステップに実コードあり。

**3. Type consistency**
- 21列目の名前は Task 1 の `OPTIONAL_HEADERS[1]` と Task 2 の `HEADERS[20]` の両方で `'部隊'` — 一致
- member の項目名は Task 2 のGAS出力・Task 1 の `sanitizeForStorage`・Task 6 の `memberDefaultButai` すべてで `butai` / `active` — 一致
- jobsite の項目名は Task 3 のGAS出力と Task 8 の画面で `status` — 一致
- `normalizeButai_`（GAS側・アンダースコア付き）と `normalizeButai`（画面側・なし）は**別物**。GASの命名規約（private関数は末尾 `_`）に従っており意図的
