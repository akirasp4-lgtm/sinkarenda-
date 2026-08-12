# 社長予定カレンダー 通信分離 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 社長予定のWeb通信を日報・集計の長時間ロックから分離し、朝6時・夜21時のLINE通知契約を変えずに読み込み・保存失敗を解消する。

**Architecture:** `doPost(e)` で認証とaction判定を先に行い、`pres_*` だけを専用ハンドラーへ早期分岐する。`pres_list` はロックなしの読み取り、3つの書き込みは同一Googleユーザー内で `getUserLock()` により直列化し、日報系は既存の `getScriptLock()` を維持する。異なるGoogleユーザー間の競合でも行ずれを起こさないよう、削除は12列のクリアとして実装し、一覧ではIDが空の行を除外する。

**Tech Stack:** Google Apps Script JavaScript、Google Sheets、Node.js built-in `node:test` / `node:vm`、clasp。

## Global Constraints

- `社長予定`シート名と12列の名前・順序を変更しない。
- `pres_list` / `pres_add` / `pres_update` / `pres_delete` の入出力を変更しない。
- GASウェブアプリURL、デプロイID、PIN、合言葉を変更しない。
- `president.html` とLINEボット本番VMは変更しない。
- 朝6時・夜21時のLINE送信は実行せず、同じ `pres_list` 契約で読めることだけ確認する。
- 既存の未保存変更はステージ・コミットしない。

---

### Task 1: GAS回帰テストを先に作る

**Files:**
- Create: `tools/president/test_lock_isolation.mjs`
- Read: `gas.js`

**Interfaces:**
- Consumes: `doPost(e)` とGASグローバルサービス。
- Produces: `node --test tools/president/test_lock_isolation.mjs` で実行できる回帰テスト。

- [ ] **Step 1: Node VM上にGAS実行環境を作る**

`gas.js` 本体を `vm.runInContext` で実行し、外部サービスだけをインメモリ実装に差し替える。Fake Sheetは実際の12列と複数件の社長予定を持ち、`appendRow` / `setValues` / `clearContent` を実データへ反映する。

```js
const context = vm.createContext({
  LockService: {
    getScriptLock: () => scriptLock,
    getUserLock: () => userLock,
  },
  SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet },
  ContentService,
  PropertiesService,
  Session,
  Utilities,
});
vm.runInContext(gasSource, context);
```

- [ ] **Step 2: 壊れる変更を名前にしたテストを書く**

```js
test('pres_list bypasses a busy daily-report script lock', () => {
  const app = loadGas({ scriptLockAvailable: false });
  const body = post(app, { action: 'pres_list', pin: '1203' });
  assert.equal(body.status, 'ok');
  assert.equal(app.metrics.scriptTry, 0);
});

test('pres_list skips daily-sheet initialization and preserves the 12-column contract', () => {
  const app = loadGas();
  const body = post(app, { action: 'pres_list', pin: '1203' });
  assert.deepEqual(Object.keys(body.rows[0]), PRES_HEADERS);
  assert.equal(app.metrics.dailyDataReads, 0);
});

for (const action of ['pres_add', 'pres_update', 'pres_delete']) {
  test(`${action} uses the president write lock instead of the daily-report lock`, () => {
    const app = loadGas({ scriptLockAvailable: false });
    assert.equal(post(app, payloadFor(action)).status, 'ok');
    assert.equal(app.metrics.scriptTry, 0);
    assert.equal(app.metrics.userTry, 1);
  });
}

test('daily-report actions remain protected by the script lock', () => {
  const app = loadGas({ scriptLockAvailable: false });
  assert.equal(post(app, { action: 'add', rows: [] }).status, 'error');
  assert.equal(app.metrics.scriptTry, 1);
});
```

- [ ] **Step 3: REDを確認する**

Run: `node --test tools/president/test_lock_isolation.mjs`

Expected: `pres_list` と3つの `pres_*` 書き込みテストが、現在の共通script lock利用または日報シート読取りを理由にFAILする。日報actionの既存ロックテストはPASSする。

- [ ] **Step 4: テストだけをコミットする**

```powershell
git add -- tools/president/test_lock_isolation.mjs
git commit -m "test: 社長予定のロック分離回帰テストを追加"
```

---

### Task 2: 社長予定の専用ハンドラーを実装する

**Files:**
- Modify: `gas.js:128-846`
- Test: `tools/president/test_lock_isolation.mjs`

**Interfaces:**
- Consumes: `body`, `action`, `updatedBy`、既存 `ok` / `error` / `authError_`。
- Produces: `isPresidentAction_(action): boolean` と `handlePresidentAction_(body, action, updatedBy): TextOutput`。

- [ ] **Step 1: `doPost` の共通前処理をロック前へ移す**

```js
function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return error(err.toString());
  }
  if (!calAuthOk_(body.k)) return authError_();
  const action = body.action || 'add';
  const updatedBy = String(body.updatedBy || '');
  if (isPresidentAction_(action)) {
    return handlePresidentAction_(body, action, updatedBy);
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return error('現在他の人が更新中です。数秒待ってから再度お試しください。');
  }
  // 以降は既存の日報処理
}
```

- [ ] **Step 2: 読み取り専用 `pres_list` を実装する**

```js
function isPresidentAction_(action) {
  return ['pres_list', 'pres_add', 'pres_update', 'pres_delete'].includes(action);
}

function handlePresidentAction_(body, action, updatedBy) {
  if (String(body.pin || '') !== PRES_PIN) return error('認証に失敗しました');
  if (action === 'pres_list') {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(PRES_SHEET);
    if (!sheet) return ok({rows: []});
    return ok({rows: serializePresidentRows_(sheet)});
  }
  // 書き込み処理へ続く
}
```

`serializePresidentRows_` は `PRES_HEADERS` をレスポンスキーに使い、日付・時刻の既存整形を維持する。

- [ ] **Step 3: 3つの書き込みをユーザーロックで保護する**

```js
const lock = LockService.getUserLock();
if (!lock.tryLock(10000)) {
  return error('現在他の人が更新中です。数秒待ってから再度お試しください。');
}
try {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const presSheet = getOrCreatePresSheet_(ss);
  // 既存 pres_add / pres_update / pres_delete 本体をそのまま移す
} catch (err) {
  return error(err.toString());
} finally {
  lock.releaseLock();
}
```

- [ ] **Step 4: 旧 `doPost` 内の `pres_*` ブロックを削除する**

同じ処理が二重に残らないよう、旧751〜846行相当を削除する。他のaction分岐は変更しない。

- [ ] **Step 5: 削除で行番号を動かさない**

`pres_delete` は対象行を物理削除せず、12列を `clearContent()` する。`pres_list` はIDが空の行を除外する。削除と別予定の更新が異なるGoogleユーザーから同時に走っても、更新対象の行番号が別予定へずれないことを回帰テストで確認する。

- [ ] **Step 6: GREENを確認する**

Run: `node --test tools/president/test_lock_isolation.mjs`

Expected: 全テストPASS。

- [ ] **Step 7: 実装をコミットする**

```powershell
git add -- gas.js
git commit -m "fix: 社長予定を日報の長時間ロックから分離"
```

---

### Task 3: ローカル全検証を行う

**Files:**
- Verify: `gas.js`
- Verify: `tools/president/test_lock_isolation.mjs`
- Verify: `tools/holidays/test_holidays.mjs`

**Interfaces:**
- Consumes: Task 1・2の完成物。
- Produces: デプロイ可否の検証記録。

- [ ] **Step 1: JavaScript構文を検査する**

Run: `node --check gas.js`

Expected: exit 0、出力なし。

- [ ] **Step 2: 社長予定回帰テストを再実行する**

Run: `node --test tools/president/test_lock_isolation.mjs`

Expected: 全テストPASS。

- [ ] **Step 3: 既存祝日回帰テストを実行する**

Run: `node tools/holidays/test_holidays.mjs`

Expected: `ALL PASS`。

- [ ] **Step 4: GASデプロイ前チェックを実行する**

Run: `python C:/Users/akira/.agents/skills/gas-deploy/scripts/predeploy_check.py .`

Expected: syntax ERRORなし。既知の公開PIN・トークン警告は、今回未変更のため差分と照合して判断する。

- [ ] **Step 5: 差分を確認する**

Run: `git diff HEAD~2 -- gas.js tools/president/test_lock_isolation.mjs`

Expected: 社長予定の分岐・テスト以外の変更なし。

---

### Task 4: 既存GASウェブアプリへ反映する

**Files:**
- Source: `gas.js`
- Deploy mirror: `C:/Users/akira/gr_pptx_build/yotei-gas/コード.js`

**Interfaces:**
- Consumes: 検証済み `gas.js`。
- Produces: 同じデプロイID・URLの新バージョン。

- [ ] **Step 1: claspの所有アカウントと対象scriptIdを確認する**

Run: `clasp --version`

Run: `clasp login --status`

Run: `Get-Content -Raw C:/Users/akira/gr_pptx_build/yotei-gas/.clasp.json`

Expected: 対象scriptIdが予定管理本番 `1BXSKkYbrU4nhuFVi_YsujzP19zMpBHmG_xxMhysUp-365yg0BaeSMV5t`。所有アカウントでない場合は本番変更を止め、ブラウザ経路へ切り替える。

- [ ] **Step 2: デプロイ用ファイルを更新し、差分を検証する**

`gas.js` を `コード.js` に同期後、SHA-256が一致することを確認する。同期先の `.clasp.json` / `appsscript.json` は変更しない。

- [ ] **Step 3: 既存デプロイを更新する**

Run: `clasp push -f`

Run: `clasp deploy --deploymentId AKfycbxp2eUcpIjCj0ZWyAPPD9m3egJrKdWmXRK2AVnFrmBm4iO1QHCk-FZEH5LFFv7OloqcjQ --description "社長予定のロック分離 2026-08-12"`

Expected: URL不変で新バージョン作成成功。

---

### Task 5: 本番契約と速度を確認する

**Files:**
- Verify only: 本番GAS・本番LINEボットVM。

**Interfaces:**
- Consumes: 更新済み本番GAS。
- Produces: 既存LINE通知契約維持とCRUD成功の証拠。

- [ ] **Step 1: `pres_list` の契約を確認する**

本番VMの `gas_client.get_pres_events(force=True)` を呼び、`status=ok`、既存12キー、取得件数、応答時間を確認する。LINEの `push` は呼ばない。

- [ ] **Step 2: 通知対象外の日付でCRUDを確認する**

2099-12-31のタイトル `Codexロック分離確認` を `pres_add` し、`pres_list` で同じID・12列を確認後、必ず `pres_delete` する。削除後に同じIDが0件であることを確認する。

- [ ] **Step 3: 社長画面の更新速度を計測する**

本番 `president.html` を開き、更新ボタンからローディング終了までを計測する。既存予定は編集しない。

- [ ] **Step 4: 最終状態を確認する**

Run: `git status --short --branch`

Expected: 今回のコミット済みファイル以外は、作業開始前からの利用者変更だけが残る。
