# 現場管理タブ強化 + 倉庫請求集計 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 現場管理タブに「👤人で探す」「📌予定だけ」モードを追加し、倉庫作業を「元請別請求集計_フィルタ用」シートに人別×日別で出力できるようにする。

**Architecture:** 現場管理タブ上部にモード切り替えボタンを追加。`gmMode` 変数と `getActiveGmGroups()` 関数でデータ源を差し替え、既存の `renderGmList` / 編集・削除 UI を全モード共通で流用する。GAS 側は `generateBillingFilterSheet_` の倉庫除外フィルタを緩和し、倉庫を `loc='倉庫作業'` に正規化してフィルタ用シートに出す（通常の請求集計シートは変更なし）。

**Tech Stack:** Vanilla JS + HTML + CSS（index.html / admin.html） + Google Apps Script（gas.js）

**Spec:** `docs/superpowers/specs/2026-05-29-genba-search-warehouse-billing-design.md`

---

## ファイル構成

| ファイル | 変更内容 |
|---|---|
| `gas.js` | `generateBillingFilterSheet_` L1846-1849 — 倉庫除外解除 + 倉庫を正規化 + 空 genba ブロック追加 |
| `index.html` | ① HTML: screen-genba にモードバー + フィルタカード追加 ② JS: `gmMode`, `getActiveGmGroups`, `setGmMode`, `renderPersonFilter`, `getPersonGroups`, `renderPinFilter`, `setPinFilterMode`, `getPinGroups`, `refreshGmTab` 追加 ③ 既存関数の `getGmGroups()` 参照を `getActiveGmGroups()` に統一 ④ `updateGmSelects()` 呼び出し 3 箇所を `refreshGmTab()` に変更 |
| `admin.html` | index.html と同じ変更（呼び出し箇所は L1351, L1375, L1776） |

---

### Task 1: gas.js — 倉庫をフィルタ用シートに追加

**Files:**
- Modify: `gas.js` L1846-1849

- [ ] **Step 1: 対象コードを確認**

`gas.js` L1843-1851 を読み、以下のコードがあることを確認：

```js
// 倉庫は元請に請求しない作業のため除外
const workRecords = records.filter(r => r.yakin !== '休み' && r.yakin !== '予定' && r.yakin !== '倉庫');
const months = [...new Set(workRecords.map(r => r.month).filter(Boolean))].sort().reverse();
const genbas = [...new Set(workRecords.map(r => r.genba).filter(Boolean))].sort();
```

- [ ] **Step 2: 倉庫を正規化してフィルタ用シートに含める**

上記 4 行を以下に置き換える：

```js
// 休み/予定は除外。倉庫は genba='', loc='倉庫作業' に正規化してフィルタ用シートに含める
// （通常の請求集計シート generateBillingSummary_ は変更なし — 倉庫除外のまま）
const workRecords = records
  .filter(r => r.yakin !== '休み' && r.yakin !== '予定')
  .map(r => r.yakin === '倉庫'
    ? Object.assign({}, r, { genba: '', loc: '倉庫作業' })
    : r);
const months = [...new Set(workRecords.map(r => r.month).filter(Boolean))].sort().reverse();
const genbas = [...new Set(workRecords.map(r => r.genba).filter(Boolean))].sort();
// 倉庫作業が存在する場合は空 genba ブロックを末尾に追加（倉庫は元請名が空）
if (workRecords.some(r => r.loc === '倉庫作業' && !r.genba)) genbas.push('');
```

- [ ] **Step 3: Commit**

```bash
cd "C:/Users/akira/OneDrive/Desktop/Claude/予定管理アプリ作成"
git add gas.js
git commit -m "fix(gas): 倉庫を元請別請求集計_フィルタ用シートに出力（loc=倉庫作業で正規化）"
```

- [ ] **Step 4: GAS デプロイ**

`gas-deploy` スキルを使って gas.js を Apps Script にデプロイし、Web App を再デプロイする。

- [ ] **Step 5: 手動確認**

1. アプリの「事務」タブ → 「集計を更新」をクリック
2. スプレッドシート「元請別請求集計_フィルタ用」を開く
3. 「現場名」列に「倉庫作業」行が人別＋合計行で出ていることを確認
4. 「元請別請求集計」（通常シート）には倉庫行が出ていないことを確認
5. フィルタで「現場名 = 倉庫作業」に絞れることを確認

---

### Task 2: index.html — JS基盤（gmMode / getActiveGmGroups / 既存関数の共通化）

**Files:**
- Modify: `index.html` L2574 付近（`// ========== 現場別管理 ==========` セクション）

- [ ] **Step 1: gmMode 変数 + コア関数を追加**

`// ========== 現場別管理 ==========` コメントの直後、`function updateGmSelects(){` の**前**に以下を挿入：

```js
// 現場管理タブ: 探し方モード ('genba' | 'person' | 'pin')
let gmMode = 'genba';
let pinFilterMode = 'overdue'; // 'overdue' | 'month'

function refreshGmTab() {
  if (gmMode === 'genba') updateGmSelects();
  else if (gmMode === 'person') renderPersonFilter();
  else if (gmMode === 'pin') renderPinFilter();
}

function setGmMode(mode) {
  gmMode = mode;
  ['genba','person','pin'].forEach(m => {
    const btn = document.getElementById('gm-mode-' + m);
    if (btn) btn.classList.toggle('active', m === mode);
  });
  const cards = {
    genba: document.getElementById('gm-filter-genba'),
    person: document.getElementById('gm-filter-person'),
    pin: document.getElementById('gm-filter-pin'),
  };
  Object.entries(cards).forEach(([k, el]) => {
    if (el) el.style.display = k === mode ? '' : 'none';
  });
  document.getElementById('gm-list-wrap').style.display = 'none';
  refreshGmTab();
}

function getActiveGmGroups() {
  if (gmMode === 'person') return getPersonGroups();
  if (gmMode === 'pin') return getPinGroups();
  return getGmGroups();
}
```

- [ ] **Step 2: renderGmList の getGmGroups 参照を置き換え**

`renderGmList` 関数（L2623 付近）内の 2 箇所を変更する：

変更前（先頭行）:
```js
function renderGmList(){
  const groups=getGmGroups();
```
変更後:
```js
function renderGmList(){
  const groups=getActiveGmGroups();
```

変更前（編集ボタン内の onclick）:
```js
<button class="gm-edit-btn" onclick="openEditModal(${idx},getGmGroups())">編集</button>
```
変更後:
```js
<button class="gm-edit-btn" onclick="openEditModal(${idx},getActiveGmGroups())">編集</button>
```

- [ ] **Step 3: 期限切れピン行に赤ボーダーを追加**

`renderGmList` 内の `groups.map((g,idx)=>{` ループの `return` 直前に 1 行追加：

```js
// 変更前:
    return`<div class="gm-day-row">
```
```js
// 変更後:
    const _overdueStyle=(gmMode==='pin'&&pinFilterMode==='overdue')?' style="border-left:3px solid #ff3b30;padding-left:12px"':'';
    return`<div class="gm-day-row"${_overdueStyle}>
```

- [ ] **Step 4: openBulkEditModal / deleteGmChecked の参照を統一**

`openBulkEditModal` 関数（L2666 付近）:

変更前:
```js
function openBulkEditModal(){
  const groups=getGmGroups();
```
変更後:
```js
function openBulkEditModal(){
  const groups=getActiveGmGroups();
```

`deleteGmChecked` 関数（L2702 付近）:

変更前:
```js
async function deleteGmChecked(){
  const groups=getGmGroups();
```
変更後:
```js
async function deleteGmChecked(){
  const groups=getActiveGmGroups();
```

- [ ] **Step 5: updateGmSelects 呼び出し 3 箇所を refreshGmTab に変更**

以下の 3 行をそれぞれ `refreshGmTab()` に変更する（`updateGmSelects()` の**定義**は変更しない）：

- L1475 付近: `updateGmSelects();` → `refreshGmTab();`（全体リフレッシュ関数の中）
- L1500 付近: `if(t==='genba')updateGmSelects();` → `if(t==='genba')refreshGmTab();`（switchTab 内）
- L1964 付近: `updateGmSelects();` → `refreshGmTab();`（loadData コールバック内）

- [ ] **Step 6: Commit**

```bash
git add index.html
git commit -m "refactor(index): gmMode + getActiveGmGroups で現場管理タブのデータ源を共通化"
```

---

### Task 3: index.html — HTML（モード切り替えバー + フィルタカード）

**Files:**
- Modify: `index.html` L449-476（`screen-genba` セクション）

- [ ] **Step 1: CSS を追加**

既存 `<style>` ブロックの末尾（`</style>` の直前）に追加：

```css
.gm-mode-bar{display:flex;gap:8px;margin-bottom:8px;}
.gm-mode-bar .btn{flex:1;font-size:12px;padding:6px 4px;}
.btn-secondary.active{background:#1D9E75 !important;color:#fff !important;border-color:#1D9E75 !important;}
```

- [ ] **Step 2: screen-genba の HTML を書き換え**

`screen-genba` 内の先頭部分（`<div class="page-title">` から最初の `</div>` カード終わりまで）を以下に置き換える：

変更前:
```html
<div id="screen-genba" class="screen">
<div class="page-title">現場別管理</div>
<div class="card">
  <label>元請名 / 月</label>
  <div style="display:flex;gap:8px;margin-bottom:8px">
    <select id="gm-genba" onchange="onGmGenbaChange()" style="flex:2;margin-bottom:0"><option value="">元請を選択</option></select>
    <select id="gm-month" onchange="onGmGenbaChange()" style="flex:1;margin-bottom:0"><option value="">全期間</option></select>
  </div>
  <label>現場名</label>
  <select id="gm-location" onchange="renderGmList()" style="margin-bottom:0"><option value="">現場を選択</option></select>
</div>
```

変更後:
```html
<div id="screen-genba" class="screen">
<div class="page-title">現場別管理</div>
<div class="gm-mode-bar">
  <button id="gm-mode-genba" class="btn btn-secondary active" onclick="setGmMode('genba')">🏗️ 現場で探す</button>
  <button id="gm-mode-person" class="btn btn-secondary" onclick="setGmMode('person')">👤 人で探す</button>
  <button id="gm-mode-pin" class="btn btn-secondary" onclick="setGmMode('pin')">📌 予定だけ</button>
</div>
<!-- 現場フィルタ（既存） -->
<div id="gm-filter-genba" class="card">
  <label>元請名 / 月</label>
  <div style="display:flex;gap:8px;margin-bottom:8px">
    <select id="gm-genba" onchange="onGmGenbaChange()" style="flex:2;margin-bottom:0"><option value="">元請を選択</option></select>
    <select id="gm-month" onchange="onGmGenbaChange()" style="flex:1;margin-bottom:0"><option value="">全期間</option></select>
  </div>
  <label>現場名</label>
  <select id="gm-location" onchange="renderGmList()" style="margin-bottom:0"><option value="">現場を選択</option></select>
</div>
<!-- 人フィルタ（新規） -->
<div id="gm-filter-person" class="card" style="display:none">
  <label>人を選ぶ</label>
  <select id="gm-person" onchange="renderGmList()" style="margin-bottom:8px"><option value="">人を選択</option></select>
  <label>月（任意）</label>
  <select id="gm-person-month" onchange="renderGmList()"><option value="">全期間</option></select>
</div>
<!-- 予定ピンフィルタ（新規） -->
<div id="gm-filter-pin" class="card" style="display:none">
  <div style="display:flex;gap:8px;margin-bottom:8px">
    <button id="gm-pin-overdue-btn" class="btn btn-secondary active" style="flex:1;font-size:12px" onclick="setPinFilterMode('overdue')">📌 期限切れのみ</button>
    <button id="gm-pin-month-btn" class="btn btn-secondary" style="flex:1;font-size:12px" onclick="setPinFilterMode('month')">月で絞る</button>
  </div>
  <select id="gm-pin-month" onchange="renderGmList()" style="display:none"><option value="">月を選択</option></select>
</div>
```

- [ ] **Step 3: Commit**

```bash
git add index.html
git commit -m "feat(index): 現場管理タブにモード切り替えUI追加（現場/人/📌予定）"
```

---

### Task 4: index.html — getPersonGroups / renderPersonFilter 実装

**Files:**
- Modify: `index.html`（現場別管理セクション末尾 `// ========== 空き確認 ==========` の直前）

- [ ] **Step 1: renderPersonFilter と getPersonGroups を追加**

`// ========== 空き確認 ==========` コメントの**直前**に以下を挿入：

```js
function renderPersonFilter() {
  const pSel = document.getElementById('gm-person');
  if (pSel) {
    const shokunin = getShokunin();
    const curP = pSel.value;
    pSel.innerHTML = '<option value="">人を選択</option>' +
      shokunin.map(s => `<option value="${esc(s)}"${s===curP?' selected':''}>${esc(s)}</option>`).join('');
  }
  const mSel = document.getElementById('gm-person-month');
  if (mSel) {
    const nippos = filteredNippos().filter(n => !n.isGhost);
    const months = [...new Set(nippos.map(n => (n.date||'').slice(0,7)).filter(Boolean))].sort().reverse();
    const curM = mSel.value;
    mSel.innerHTML = '<option value="">全期間</option>' + months.map(m => {
      const p = m.split('-');
      return `<option value="${m}"${m===curM?' selected':''}>${p[0]}年${Number(p[1])}月</option>`;
    }).join('');
  }
  renderGmList();
}

function getPersonGroups() {
  const person = (document.getElementById('gm-person')||{}).value || '';
  const month  = (document.getElementById('gm-person-month')||{}).value || '';
  if (!person) return [];
  let nippos = filteredNippos().filter(n => !n.isGhost);
  if (month) nippos = nippos.filter(n => (n.date||'').startsWith(month));
  return groupNippos(nippos)
    .filter(g => g.members.some(m => m.name === person))
    .sort((a, b) => a.date.localeCompare(b.date));
}
```

- [ ] **Step 2: Commit**

```bash
git add index.html
git commit -m "feat(index): 人で探すモード実装（getPersonGroups / renderPersonFilter）"
```

- [ ] **Step 3: 手動確認**

1. 現場管理タブ →「👤 人で探す」ボタンをタップ → 人を選ぶセレクタと月セレクタが出る
2. 職人マスタの名前が選べることを確認
3. 名前を選ぶ → その人が代表 or メンバーで入った予定が日付順に一覧表示される
4. 月で絞ると当月分のみになる
5. 1件タップ → 詳細/編集/削除が従来通り動く

---

### Task 5: index.html — getPinGroups / renderPinFilter / setPinFilterMode 実装

**Files:**
- Modify: `index.html`（Task 4 で追加した `getPersonGroups` の直後）

- [ ] **Step 1: renderPinFilter / setPinFilterMode / getPinGroups を追加**

`getPersonGroups` 関数の直後に以下を挿入：

```js
function renderPinFilter() {
  const mSel = document.getElementById('gm-pin-month');
  if (mSel) {
    const nippos = filteredNippos().filter(n => !n.isGhost && n.yotei);
    const months = [...new Set(nippos.map(n => (n.date||'').slice(0,7)).filter(Boolean))].sort().reverse();
    const curM = mSel.value;
    mSel.innerHTML = '<option value="">月を選択</option>' + months.map(m => {
      const p = m.split('-');
      return `<option value="${m}"${m===curM?' selected':''}>${p[0]}年${Number(p[1])}月</option>`;
    }).join('');
  }
  renderGmList();
}

function setPinFilterMode(mode) {
  pinFilterMode = mode;
  ['overdue','month'].forEach(m => {
    const btn = document.getElementById('gm-pin-' + m + '-btn');
    if (btn) btn.classList.toggle('active', m === mode);
  });
  const mSel = document.getElementById('gm-pin-month');
  if (mSel) mSel.style.display = mode === 'month' ? '' : 'none';
  renderGmList();
}

function getPinGroups() {
  const today = new Date().toISOString().slice(0, 10);
  let nippos = filteredNippos().filter(n => !n.isGhost && n.yotei);
  if (pinFilterMode === 'overdue') {
    nippos = nippos.filter(n => (n.date||'') < today);
  } else {
    const month = (document.getElementById('gm-pin-month')||{}).value || '';
    if (month) nippos = nippos.filter(n => (n.date||'').startsWith(month));
  }
  return groupNippos(nippos).sort((a, b) => a.date.localeCompare(b.date));
}
```

- [ ] **Step 2: Commit**

```bash
git add index.html
git commit -m "feat(index): 📌予定ピン一覧モード実装（getPinGroups / setPinFilterMode）"
```

- [ ] **Step 3: 手動確認**

1. 「📌 予定だけ」ボタンをタップ
2. 初期表示（期限切れのみ）: 今日より前の📌予定が赤い左ボーダー付きで並ぶ
3. 「月で絞る」ボタンをタップ → 月セレクタが現れ、選んだ月の📌が全部出る
4. 1件タップ → 編集画面で「予定（📌）」→「現場作業」に変更して保存 → 一覧から消える（確定化成功）
5. 削除ボタンでも削除できることを確認

---

### Task 6: admin.html — 同じ変更を適用

**Files:**
- Modify: `admin.html`

admin.html は index.html と同じ構造を持つ。Task 2〜5 で行った変更をすべて適用する。

- [ ] **Step 1: 変更箇所を確認**

admin.html の対応行番号を確認：
- `updateGmSelects` 呼び出し: L1351, L1375, L1776（3箇所を `refreshGmTab()` に変更）
- `screen-genba` セクション: `gm-genba`, `gm-month`, `gm-location` のある HTML ブロック
- `// ========== 現場別管理 ==========` セクション: `function updateGmSelects` の前（L2497付近）
- `renderGmList` 内の `getGmGroups()` 参照 2 箇所 → `getActiveGmGroups()`
- `openBulkEditModal` / `deleteGmChecked` 内の `getGmGroups()` → `getActiveGmGroups()`

- [ ] **Step 2: CSS を追加**

admin.html の `</style>` 直前に（index.html と同じ CSS を追加）：

```css
.gm-mode-bar{display:flex;gap:8px;margin-bottom:8px;}
.gm-mode-bar .btn{flex:1;font-size:12px;padding:6px 4px;}
.btn-secondary.active{background:#1D9E75 !important;color:#fff !important;border-color:#1D9E75 !important;}
```

- [ ] **Step 3: Task 2 の JS を admin.html に適用**

`// ========== 現場別管理 ==========` 直後、`function updateGmSelects(){` の前に以下を挿入（index.html の Task 2 Step 1 と同一コード）：

```js
let gmMode = 'genba';
let pinFilterMode = 'overdue';

function refreshGmTab() {
  if (gmMode === 'genba') updateGmSelects();
  else if (gmMode === 'person') renderPersonFilter();
  else if (gmMode === 'pin') renderPinFilter();
}

function setGmMode(mode) {
  gmMode = mode;
  ['genba','person','pin'].forEach(m => {
    const btn = document.getElementById('gm-mode-' + m);
    if (btn) btn.classList.toggle('active', m === mode);
  });
  const cards = {
    genba: document.getElementById('gm-filter-genba'),
    person: document.getElementById('gm-filter-person'),
    pin: document.getElementById('gm-filter-pin'),
  };
  Object.entries(cards).forEach(([k, el]) => {
    if (el) el.style.display = k === mode ? '' : 'none';
  });
  document.getElementById('gm-list-wrap').style.display = 'none';
  refreshGmTab();
}

function getActiveGmGroups() {
  if (gmMode === 'person') return getPersonGroups();
  if (gmMode === 'pin') return getPinGroups();
  return getGmGroups();
}
```

- [ ] **Step 4: renderGmList / openBulkEditModal / deleteGmChecked の参照を統一**

admin.html の `renderGmList`, `openBulkEditModal`, `deleteGmChecked` 内の `getGmGroups()` を `getActiveGmGroups()` に変更（Task 2 Step 2〜4 と同じ）。

- [ ] **Step 5: Task 2 Step 3 の赤ボーダーを admin.html にも追加**

admin.html の `renderGmList` 内の `return` 直前に:

```js
const _overdueStyle=(gmMode==='pin'&&pinFilterMode==='overdue')?' style="border-left:3px solid #ff3b30;padding-left:12px"':'';
return`<div class="gm-day-row"${_overdueStyle}>
```

- [ ] **Step 6: updateGmSelects 呼び出し 3 箇所を refreshGmTab に変更**

- L1351 付近: `updateGmSelects();` → `refreshGmTab();`
- L1375 付近: `if(t==='genba')updateGmSelects();` → `if(t==='genba')refreshGmTab();`
- L1776 付近: `updateGmSelects();` → `refreshGmTab();`

- [ ] **Step 7: Task 3 の HTML を admin.html に適用**

admin.html の `screen-genba` 先頭部分を Task 3 Step 2 と同じ HTML に書き換える（page-title → gm-mode-bar → gm-filter-genba / gm-filter-person / gm-filter-pin）。

- [ ] **Step 8: Task 4 / 5 の JS を admin.html に適用**

`// ========== 空き確認 ==========`（または類似のセクション区切り）の直前に、`renderPersonFilter`, `getPersonGroups`, `renderPinFilter`, `setPinFilterMode`, `getPinGroups` を追加（index.html の Task 4 / 5 と同一コード）。

- [ ] **Step 9: Commit & Push**

```bash
git add admin.html
git commit -m "feat(admin): 現場管理タブに人で探す/📌予定ピン一覧モードを追加（index.htmlと同期）"
git push origin main
```

- [ ] **Step 10: 最終動作確認**

1. ブラウザ Ctrl+Shift+R でキャッシュクリア
2. **②人で探す**: 現場管理タブ →「👤 人で探す」→ 名前選択 → 予定一覧が出る
3. **③予定ピン**: 「📌 予定だけ」→ 期限切れピンが赤ボーダー付きで出る → 1件開いて確定化できる
4. **③月切り替え**: 「月で絞る」→ 月選択 → その月の📌が出る
5. **①現場で探す**: 従来通り動くことを確認（回帰なし）
6. **④倉庫集計**: 事務タブ「集計を更新」→ フィルタ用シートに「倉庫作業」行が出る（Task 1 で確認済みであれば省略可）
7. admin.html でも同様の動作を確認

---

## 自己レビュー

**Spec カバレッジ確認:**
- ② 個人名検索（代表/メンバー両方） → Task 4 ✅
- ③ 期限切れピン一覧 + 月で絞る + 確定化動線 → Task 5 ✅
- ④ 倉庫をフィルタ用シートへ（通常シートは変更なし） → Task 1 ✅
- admin.html への適用 → Task 6 ✅
- データ再読み込み時のモード別更新（refreshGmTab） → Task 2 Step 5 ✅

**型整合性:**
- `getActiveGmGroups()` は `getGmGroups()` と同じ `groups` 配列を返す。`renderGmList`, `openBulkEditModal`, `deleteGmChecked`, `openEditModal` はすべて `groups[idx]` 形式で使うので互換性あり ✅
- `getPinGroups()` / `getPersonGroups()` は `groupNippos()` の戻り値をソートして返すだけなので同じ構造 ✅
- `pinFilterMode` は `let` で宣言済み（Task 2 Step 1）、`setPinFilterMode` と `getPinGroups` の両方から参照 ✅
