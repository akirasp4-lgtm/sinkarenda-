# 請求計算（常用・日当）自動化 第1段階 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax.

**Goal:** admin.html に「💰 請求計算」タブを新設し、出面データ×単価（元請×現場×職人で記憶）＋経費で請求額を自動算出、数式入りExcelで出力できるようにする。

**Architecture:** 単価は新シート「請求単価マスタ」に元請×現場×職人キーで記憶し次回自動表示。応援/請負は現場マスタの新「請求方式」列＋画面で行ごと切替。Excel出力は GAS で「請求計算」シートを数式入り（金額=出面×単価、合計=SUM）で生成し、既存の `exportSheetAsXlsxBase64_` でxlsx化（Google Sheetsの数式はxlsxに保持される）。

**Tech Stack:** Google Apps Script + Google Sheets / Vanilla JS + HTML（admin.html）

**Spec:** `docs/superpowers/specs/2026-05-29-invoice-calculation-design.md`

---

## ファイル構成

| ファイル | 変更内容 |
|---|---|
| `gas.js` | ①`BILLING_RATE_SHEET`定数 ②`getOrCreateBillingRateSheet_` ③`get_billing_rates`/`save_billing_rate` ④現場マスタ請求方式列マイグレーション＋`doGet`に`billingMethod` ⑤`update_site_billing_method` ⑥`generate_billing_calc_xlsx`（数式シート生成→xlsx） |
| `admin.html` | ①タブ「💰請求計算」追加 ②`screen-billing` HTML ③請求計算JS（出面集計・単価表示/保存・金額/経費/方式・合計・Excel出力） |
| `index.html` | 変更なし |

検証は手動（GAS実行 + ブラウザ操作）。自動テストフレームワークは無いプロジェクトなので、各タスクに手動確認手順を記す。

---

### Task 1: gas.js — 請求単価マスタ（シート＋読み書き）

**Files:** Modify `gas.js`（定数は行12付近、関数はマスタ系関数の近く、アクションは doPost 内）

- [ ] **Step 1: シート定数を追加**

行12 `const BILLING_FILTER_SHEET = '元請別請求集計_フィルタ用';` の直後に追加：

```javascript
const BILLING_RATE_SHEET = '請求単価マスタ';
```

- [ ] **Step 2: シート取得/作成関数を追加**

`getOrCreateJobSiteSheet_` 関数（行1023付近）の直後に追加：

```javascript
function getOrCreateBillingRateSheet_(ss) {
  let sheet = ss.getSheetByName(BILLING_RATE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(BILLING_RATE_SHEET);
    sheet.appendRow(['元請名', '現場名', '職人名', '単価', '更新日時']);
  }
  return sheet;
}
```

- [ ] **Step 3: get_billing_rates / save_billing_rate アクションを追加**

doPost 内、`update_site_revenue` アクション（行494付近）の直後に追加：

```javascript
if (action === 'get_billing_rates') {
  const sheet = getOrCreateBillingRateSheet_(ss);
  const data = sheet.getDataRange().getValues();
  const rates = data.length > 1 ? data.slice(1).map(r => ({
    genba: String(r[0] || ''),
    loc: String(r[1] || ''),
    name: String(r[2] || ''),
    rate: Number(r[3] || 0)
  })).filter(x => x.genba) : [];
  return ok({rates: rates});
}

if (action === 'save_billing_rate') {
  const sheet = getOrCreateBillingRateSheet_(ss);
  const genba = String(body.genba || '').trim();
  const loc = String(body.loc || '').trim();
  const name = String(body.name || '').trim();
  const rate = Number(body.rate || 0);
  if (!genba || !name) return error('元請名・職人名は必須です');
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === genba && String(data[i][1]).trim() === loc && String(data[i][2]).trim() === name) {
      sheet.getRange(i + 1, 4).setValue(rate);
      sheet.getRange(i + 1, 5).setValue(now);
      return ok({updated: true});
    }
  }
  sheet.appendRow([genba, loc, name, rate, now]);
  return ok({added: true});
}
```

- [ ] **Step 4: Commit**

```bash
cd "C:/Users/akira/OneDrive/Desktop/Claude/予定管理アプリ作成"
git add gas.js
git commit -m "feat(gas): 請求単価マスタ（元請×現場×職人で単価記憶）の読み書きアクション追加"
```

---

### Task 2: gas.js — 現場マスタ「請求方式」列

**Files:** Modify `gas.js`（`getOrCreateJobSiteSheet_` 行1023付近、`doGet` の jobsites 行950付近、doPost）

- [ ] **Step 1: 現場マスタを10列に拡張（マイグレーション）**

`getOrCreateJobSiteSheet_`（行1023付近）を以下に置き換える：

```javascript
function getOrCreateJobSiteSheet_(ss) {
  let sheet = ss.getSheetByName(JOBSITE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(JOBSITE_SHEET);
    sheet.appendRow(['元請名', '現場名', '工番', '事業部', '年度', '連番', '売上', '読み', '完了', '請求方式']);
  } else {
    ensureColumns_(sheet, 10);
    const headers = sheet.getRange(1, 1, 1, 10).getValues()[0];
    if (String(headers[6] || '').trim() !== '売上') sheet.getRange(1, 7).setValue('売上');
    if (String(headers[7] || '').trim() !== '読み') sheet.getRange(1, 8).setValue('読み');
    if (String(headers[8] || '').trim() !== '完了') sheet.getRange(1, 9).setValue('完了');
    if (String(headers[9] || '').trim() !== '請求方式') sheet.getRange(1, 10).setValue('請求方式');
  }
  return sheet;
}
```

- [ ] **Step 2: doGet の jobsites に billingMethod を追加**

doGet の jobsites 組み立て（行950付近）を以下に置き換える（空欄は「応援」をデフォルトに）：

```javascript
const jobsites = jData.length > 1 ? jData.slice(1).map(r => ({
  genba: String(r[0] || ''),
  loc: String(r[1] || ''),
  jobNo: String(r[2] || ''),
  completed: String(r[8] || '').trim() !== '',
  billingMethod: String(r[9] || '').trim() || '応援'
})).filter(j => j.genba) : [];
```

> 注意：請求方式は10列目 = 0始まりインデックス `r[9]`（完了は `r[8]`）。

- [ ] **Step 3: update_site_billing_method アクションを追加**

doPost 内、Task 1 で追加した `save_billing_rate` の直後に追加：

```javascript
if (action === 'update_site_billing_method') {
  const jobSiteSheet = getOrCreateJobSiteSheet_(ss);
  const genba = String(body.genba || '').trim();
  const loc = String(body.loc || '').trim();
  const method = String(body.method || '応援').trim();
  const data = jobSiteSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === genba && String(data[i][1]).trim() === loc) {
      jobSiteSheet.getRange(i + 1, 10).setValue(method);
      logOperation_(ss, 'update_site_billing_method', genba + '/' + loc, '方式=' + method, updatedBy);
      return ok({updated: true});
    }
  }
  return error('現場マスタに該当現場が見つかりません');
}
```

- [ ] **Step 4: Commit**

```bash
git add gas.js
git commit -m "feat(gas): 現場マスタに請求方式列追加＋doGetにbillingMethod＋更新アクション"
```

---

### Task 3: gas.js — 請求計算シート生成＋xlsx出力

**Files:** Modify `gas.js`（`BILLING_CALC_SHEET`定数、doPost に `generate_billing_calc_xlsx`）

数式入りシートを作って既存 `exportSheetAsXlsxBase64_` でxlsx化する。Google Sheetsの数式（`=D2*E2`, `=SUM(...)`）はxlsxに関数として保持される。

- [ ] **Step 1: 定数を追加**

Task 1 で足した `const BILLING_RATE_SHEET` の直後に追加：

```javascript
const BILLING_CALC_SHEET = '請求計算';
```

- [ ] **Step 2: generate_billing_calc_xlsx アクションを追加**

doPost 内、`update_site_billing_method` の直後に追加。`body.lines` は admin から送る明細配列（`{loc, name, manDays, rate, method, expense}`）、`body.genba`・`body.month` はヘッダ用：

```javascript
if (action === 'generate_billing_calc_xlsx') {
  const genba = String(body.genba || '');
  const month = String(body.month || '');
  const lines = Array.isArray(body.lines) ? body.lines : [];
  let sheet = ss.getSheetByName(BILLING_CALC_SHEET);
  if (sheet) { sheet.clear(); } else { sheet = ss.insertSheet(BILLING_CALC_SHEET); }
  sheet.appendRow([genba + '　' + month + '　請求計算']);
  sheet.appendRow(['現場名', '職人名', '出面数', '単価', '金額', '経費', '方式']);
  const dataStart = 3; // 1=タイトル 2=ヘッダ 3=先頭データ
  lines.forEach(ln => {
    const r = sheet.getLastRow() + 1;
    const isOuen = String(ln.method || '応援') === '応援';
    // 金額 = 出面数 × 単価（応援のみ。請負は空欄）
    const amountFormula = isOuen ? '=C' + r + '*D' + r : '';
    sheet.appendRow([
      String(ln.loc || ''), String(ln.name || ''),
      Number(ln.manDays || 0), Number(ln.rate || 0),
      amountFormula, isOuen ? Number(ln.expense || 0) : 0,
      String(ln.method || '応援')
    ]);
  });
  const dataEnd = sheet.getLastRow();
  if (dataEnd >= dataStart) {
    const totalRow = dataEnd + 1;
    // 合計行：金額合計＋経費合計＝総合計
    sheet.getRange(totalRow, 1).setValue('合計');
    sheet.getRange(totalRow, 5).setFormula('=SUM(E' + dataStart + ':E' + dataEnd + ')');
    sheet.getRange(totalRow, 6).setFormula('=SUM(F' + dataStart + ':F' + dataEnd + ')');
    sheet.getRange(totalRow, 7).setValue('総合計');
    sheet.getRange(totalRow + 1, 4).setValue('請求合計');
    sheet.getRange(totalRow + 1, 5).setFormula('=E' + totalRow + '+F' + totalRow);
  }
  SpreadsheetApp.flush();
  const b64 = exportSheetAsXlsxBase64_(ss, sheet);
  return ok({filename: '請求計算_' + genba + '_' + month + '.xlsx', base64: b64});
}
```

- [ ] **Step 3: exportSheetAsXlsxBase64_ の戻り値を確認**

`exportSheetAsXlsxBase64_`（行1423付近）を読み、戻り値が base64 文字列であることを確認。違えば呼び出しを合わせる。`export_sheet_xlsx` アクションが返す JSON 形（`{base64, filename}` 等）に合わせて admin 側の受け取りを実装する（Task 6）。

- [ ] **Step 4: Commit**

```bash
git add gas.js
git commit -m "feat(gas): 請求計算シートを数式入りで生成しxlsx出力するアクション追加"
```

---

### Task 4: admin.html — 請求計算タブ（HTML）

**Files:** Modify `admin.html`（タブバー 行609付近、screen群、switchTab の tabs配列 行1391付近）

- [ ] **Step 1: タブボタンを追加**

タブバーの事務タブ `<button ... switchTab('jimu')>` の直後に追加：

```html
<button class="tab" onclick="switchTab('billing')"><span class="tab-icon">💰</span>請求計算</button>
```

- [ ] **Step 2: switchTab の tabs 配列に追加**

`switchTab` 内の `const tabs=['list','avail','genba','vehicle','jimu'];` を以下に（DOMのボタン順序と一致させる。billing は jimu の直後に置いたので配列も jimu の後）：

```javascript
const tabs=['list','avail','genba','vehicle','jimu','billing'];
```

そして `switchTab` の末尾付近（他タブの初期化呼び出しに倣って）に追加：

```javascript
if(t==='billing')initBillingTab();
```

- [ ] **Step 3: screen-billing の HTML を追加**

事務画面 `<div id="screen-jimu" class="screen">...</div>` の閉じ `</div>` の直後に追加：

```html
<div id="screen-billing" class="screen">
<div class="page-title">💰 請求計算（常用・日当）</div>
<div class="card">
  <label>元請名 / 月</label>
  <div style="display:flex;gap:8px;margin-bottom:0">
    <select id="bill-genba" onchange="renderBillingTable()" style="flex:2;margin-bottom:0"><option value="">元請を選択</option></select>
    <select id="bill-month" onchange="renderBillingTable()" style="flex:1;margin-bottom:0"></select>
  </div>
</div>
<div id="bill-table-wrap" style="display:none">
  <div class="card" style="padding:12px 16px;overflow-x:auto">
    <table class="tbl" id="bill-table">
      <thead><tr><th>現場</th><th>職人</th><th>出面</th><th>方式</th><th>単価</th><th>金額</th><th>経費</th></tr></thead>
      <tbody id="bill-tbody"></tbody>
    </table>
    <div id="bill-total" style="text-align:right;font-weight:600;margin-top:10px;font-size:15px"></div>
  </div>
  <button class="btn" id="bill-export-btn" onclick="exportBillingXlsx()" style="background:#1D9E75;margin-top:8px">Excelで出力（関数入り）</button>
</div>
<button class="btn btn-secondary" style="margin-top:8px" onclick="loadData()">更新</button>
</div>
```

> 経費は「現場ごとに1欄」。同じ現場の最初の行にだけ経費入力欄を出し、2行目以降は空にする（Task 5 の描画で制御）。

- [ ] **Step 4: Commit**

```bash
git add admin.html
git commit -m "feat(admin): 請求計算タブのHTML（タブ・元請月セレクタ・明細テーブル枠）追加"
```

---

### Task 5: admin.html — 請求計算ロジック（JS）

**Files:** Modify `admin.html`（現場別管理セクションの近く、または空き確認の手前にまとめて追加）

- [ ] **Step 1: グローバル状態と初期化を追加**

`// ===== 空き確認 =====` の直前に追加：

```javascript
// ===== 請求計算 =====
let billingRates = [];   // [{genba,loc,name,rate}]
let billingDirty = {};   // 画面で編集した単価 key->rate（保存用）

async function initBillingTab(){
  // 単価マスタを取得
  try{
    const res=await fetch(getGasUrl(),{method:'POST',body:JSON.stringify({action:'get_billing_rates'}),headers:{'Content-Type':'text/plain'}});
    const json=await res.json();
    if(json.status==='ok')billingRates=json.rates||[];
  }catch(e){console.error('単価マスタ取得失敗',e);}
  populateBillingSelectors();
}

function populateBillingSelectors(){
  const nippos=filteredNippos().filter(n=>!n.isGhost);
  const genbas=[...new Set(nippos.map(n=>n.genba).filter(Boolean))].sort();
  const gSel=document.getElementById('bill-genba');
  const curG=gSel.value;
  gSel.innerHTML='<option value="">元請を選択</option>'+genbas.map(g=>`<option value="${esc(g)}"${g===curG?' selected':''}>${esc(g)}</option>`).join('');
  const mSel=document.getElementById('bill-month');
  const months=[...new Set(nippos.map(n=>(n.date||'').slice(0,7)).filter(Boolean))].sort().reverse();
  const curM=mSel.value;
  mSel.innerHTML=months.map(m=>{const p=m.split('-');return `<option value="${m}"${m===curM?' selected':''}>${p[0]}年${Number(p[1])}月</option>`;}).join('');
  renderBillingTable();
}
```

- [ ] **Step 2: 出面集計＋テーブル描画を追加**

Step 1 の直後に追加。出面数は「元請×現場×職人」で、休み・予定・倉庫・ghost を除いた実働日数：

```javascript
function billingRate(genba,loc,name){
  const key=genba+'|||'+loc+'|||'+name;
  if(key in billingDirty)return billingDirty[key];
  const hit=billingRates.find(r=>r.genba===genba&&r.loc===loc&&r.name===name);
  return hit?hit.rate:'';
}
function siteBillingMethod(genba,loc){
  const j=(allJobsites||[]).find(x=>x.genba===genba&&x.loc===loc);
  return j&&j.billingMethod?j.billingMethod:'応援';
}

function getBillingLines(){
  const genba=document.getElementById('bill-genba').value;
  const month=document.getElementById('bill-month').value;
  if(!genba||!month)return[];
  let nippos=filteredNippos().filter(n=>!n.isGhost&&n.genba===genba&&(n.date||'').startsWith(month)
    &&n.yakin!=='休み'&&n.yakin!=='予定'&&n.yakin!=='倉庫'&&!n.yasumi&&!n.yotei&&!n.souko);
  // 現場×職人 → 出面した日付の集合（重複日は1出面）
  const map={};
  nippos.forEach(n=>{
    const loc=n.loc||'';
    const key=loc+'|||'+n.name;
    if(!map[key])map[key]={loc:loc,name:n.name,dates:new Set()};
    map[key].dates.add(n.date);
  });
  return Object.values(map).map(x=>({
    genba:genba,loc:x.loc,name:x.name,manDays:x.dates.size,
    method:siteBillingMethod(genba,x.loc)
  })).sort((a,b)=>a.loc.localeCompare(b.loc)||a.name.localeCompare(b.name));
}

function renderBillingTable(){
  const genba=document.getElementById('bill-genba').value;
  const lines=getBillingLines();
  const wrap=document.getElementById('bill-table-wrap');
  if(lines.length===0){wrap.style.display='none';return;}
  wrap.style.display='block';
  let prevLoc=null;
  document.getElementById('bill-tbody').innerHTML=lines.map((ln,idx)=>{
    const rate=billingRate(genba,ln.loc,ln.name);
    const isOuen=ln.method==='応援';
    const amount=(isOuen&&rate!==''&&rate!=null)?(ln.manDays*Number(rate)):'';
    const firstOfSite=ln.loc!==prevLoc; prevLoc=ln.loc;
    const expenseCell=(isOuen&&firstOfSite)
      ? `<input type="number" class="bill-exp" data-loc="${esc(ln.loc)}" style="width:80px" oninput="updateBillingTotal()">`
      : '';
    const rateInput=`<input type="number" class="bill-rate" data-genba="${esc(genba)}" data-loc="${esc(ln.loc)}" data-name="${esc(ln.name)}" value="${rate===''?'':rate}" style="width:90px" ${isOuen?'':'disabled'} oninput="onBillingRateInput(this)">`;
    const methodSel=`<select class="bill-method" data-genba="${esc(genba)}" data-loc="${esc(ln.loc)}" onchange="onBillingMethodChange(this)">
      <option value="応援"${isOuen?' selected':''}>応援</option>
      <option value="請負"${!isOuen?' selected':''}>請負</option></select>`;
    return `<tr data-idx="${idx}">
      <td>${esc(ln.loc)}</td><td>${esc(ln.name)}</td><td style="text-align:right">${ln.manDays}</td>
      <td>${methodSel}</td><td>${rateInput}</td>
      <td class="bill-amount" style="text-align:right">${amount===''?'-':amount.toLocaleString()}</td>
      <td>${expenseCell}</td></tr>`;
  }).join('');
  updateBillingTotal();
}
```

- [ ] **Step 2 メモ:** `allJobsites` は doGet 応答で更新される既存グローバル。`esc()` も既存ユーティリティ。

- [ ] **Step 3: 単価入力・方式変更・合計のハンドラを追加**

Step 2 の直後に追加：

```javascript
function onBillingRateInput(el){
  const key=el.dataset.genba+'|||'+el.dataset.loc+'|||'+el.dataset.name;
  billingDirty[key]=el.value===''?'':Number(el.value);
  renderBillingTable(); // 金額を即再計算（入力欄の値は billingDirty に保持されるので保たれる）
}

async function onBillingMethodChange(el){
  const genba=el.dataset.genba, loc=el.dataset.loc, method=el.value;
  // 同じ現場の全行に反映（現場マスタ更新）
  try{
    await fetch(getGasUrl(),{method:'POST',body:JSON.stringify({action:'update_site_billing_method',genba,loc,method,updatedBy:getUsername()}),headers:{'Content-Type':'text/plain'}});
    const j=(allJobsites||[]).find(x=>x.genba===genba&&x.loc===loc);
    if(j)j.billingMethod=method; else (allJobsites=allJobsites||[]).push({genba,loc,billingMethod:method});
  }catch(e){console.error(e);}
  renderBillingTable();
}

function updateBillingTotal(){
  let amountSum=0, expSum=0;
  document.querySelectorAll('#bill-tbody tr').forEach(tr=>{
    const a=tr.querySelector('.bill-amount').textContent.replace(/[^0-9]/g,'');
    if(a)amountSum+=Number(a);
  });
  document.querySelectorAll('.bill-exp').forEach(inp=>{if(inp.value)expSum+=Number(inp.value);});
  document.getElementById('bill-total').textContent=
    `金額合計 ${amountSum.toLocaleString()}円 ＋ 経費 ${expSum.toLocaleString()}円 ＝ 請求合計 ${(amountSum+expSum).toLocaleString()}円`;
}
```

- [ ] **Step 4: Commit**

```bash
git add admin.html
git commit -m "feat(admin): 請求計算ロジック（出面集計・単価表示/編集・方式・経費・合計）"
```

---

### Task 6: admin.html — Excel出力＋単価保存

**Files:** Modify `admin.html`（Task 5 の続き）

- [ ] **Step 1: Excel出力関数を追加**

Task 5 のハンドラ群の直後に追加。出力前に編集済み単価をマスタ保存→明細をGASに送ってxlsxを受け取りダウンロード：

```javascript
async function exportBillingXlsx(){
  const genba=document.getElementById('bill-genba').value;
  const month=document.getElementById('bill-month').value;
  if(!genba||!month){alert('元請と月を選んでください');return;}
  const btn=document.getElementById('bill-export-btn');
  btn.disabled=true;btn.textContent='出力中...';
  try{
    // 1) 編集された単価をマスタに保存
    for(const key of Object.keys(billingDirty)){
      const [g,l,n]=key.split('|||');
      const rate=billingDirty[key];
      if(rate==='')continue;
      await fetch(getGasUrl(),{method:'POST',body:JSON.stringify({action:'save_billing_rate',genba:g,loc:l,name:n,rate:Number(rate),updatedBy:getUsername()}),headers:{'Content-Type':'text/plain'}});
    }
    // 2) 明細を組み立て（経費は現場の先頭行に紐づく）
    const lines=getBillingLines();
    const expMap={};
    document.querySelectorAll('.bill-exp').forEach(inp=>{expMap[inp.dataset.loc]=Number(inp.value||0);});
    const seen={};
    const payload=lines.map(ln=>{
      const rate=billingRate(genba,ln.loc,ln.name);
      const exp=(!seen[ln.loc])?(expMap[ln.loc]||0):0; seen[ln.loc]=true;
      return {loc:ln.loc,name:ln.name,manDays:ln.manDays,rate:(rate===''?0:Number(rate)),method:ln.method,expense:exp};
    });
    // 3) xlsx生成リクエスト
    const res=await fetch(getGasUrl(),{method:'POST',body:JSON.stringify({action:'generate_billing_calc_xlsx',genba,month,lines:payload}),headers:{'Content-Type':'text/plain'}});
    const json=await res.json();
    if(json.status!=='ok'){alert('出力エラー：'+json.message);return;}
    // 4) base64 を Blob 化してダウンロード
    const bin=atob(json.base64);
    const bytes=new Uint8Array(bin.length);
    for(let i=0;i<bin.length;i++)bytes[i]=bin.charCodeAt(i);
    const blob=new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    const a=document.createElement('a');
    a.href=URL.createObjectURL(blob);a.download=json.filename||'請求計算.xlsx';
    document.body.appendChild(a);a.click();a.remove();
    // 保存済みなので dirty をマスタ側へ反映
    for(const key of Object.keys(billingDirty)){const[g,l,n]=key.split('|||');const hit=billingRates.find(r=>r.genba===g&&r.loc===l&&r.name===n);if(hit)hit.rate=Number(billingDirty[key]);else billingRates.push({genba:g,loc:l,name:n,rate:Number(billingDirty[key])});}
    billingDirty={};
  }catch(e){alert('出力エラー：'+e.message);}
  finally{btn.disabled=false;btn.textContent='Excelで出力（関数入り）';}
}
```

- [ ] **Step 2: 既存 export_sheet_xlsx の戻り値形と整合確認**

`gas.js` の `export_sheet_xlsx`（行392付近）を読み、返す JSON のキー名（`base64` / `filename` 等）を確認。`generate_billing_calc_xlsx`（Task 3）の戻り値と admin の受け取り（`json.base64` / `json.filename`）を一致させる。異なれば合わせる。

- [ ] **Step 3: Commit & Push**

```bash
git add admin.html
git commit -m "feat(admin): 請求計算のExcel出力（単価保存→数式xlsx生成→ダウンロード）"
git push origin main
```

---

### Task 7: デプロイ＆検証

- [ ] **Step 1: pre-deploy チェック**

```bash
cd "C:/Users/akira/.claude/skills/gas-deploy" && python scripts/predeploy_check.py "C:/Users/akira/OneDrive/Desktop/Claude/予定管理アプリ作成"
```
Expected: PASS（既存の VEHICLE_RES_TOKEN 等のWARNは許容）。ERROR が出たら修正。

- [ ] **Step 2: GAS デプロイ**

`gas-deploy` スキルのブラウザ手順で「予定管理」プロジェクト（scriptId `1BXSKkYbrU4nhuFVi_YsujzP19zMpBHmG_xxMhysUp-365yg0BaeSMV5t`）にコード.gs を貼り替え→保存→「デプロイを管理」→既存デプロイの新バージョン（v36）発行。デプロイID/URL不変を確認。

- [ ] **Step 3: 手動検証**

1. admin を開き Ctrl+Shift+R → 「💰請求計算」タブが出る
2. 元請＋月を選ぶ → 現場×職人×出面数が並ぶ
3. 単価を入力 → 金額（出面×単価）が即計算、合計が更新
4. 方式を「請負」に変更 → その現場の行は単価/金額/経費が対象外
5. 経費（現場の先頭行）を入力 → 請求合計に反映
6. 「Excelで出力」→ xlsx ダウンロード → Excelで開き、出面数や単価を変えると金額・合計が自動再計算（関数が入っている）
7. 別の月に切り替え→戻る、または再読込 → 入れた単価が自動表示される（マスタ保存確認）
8. スプレッドシートに「請求単価マスタ」「請求計算」シートができ、現場マスタに「請求方式」列が増えている

- [ ] **Step 4: 引き継ぎ更新**

`引き継ぎ.md` 冒頭に請求計算 第1段階の実装＋GAS v36 デプロイを追記し commit & push。

---

## 自己レビュー

**Spec カバレッジ:**
- 単価を元請×現場×職人で記憶・次回自動 → Task 1（マスタ）＋ Task 5（`billingRate`表示）＋ Task 6（保存）✅
- 応援/請負を現場マスタ方式列＋画面切替 → Task 2 ＋ Task 5（`siteBillingMethod`/`onBillingMethodChange`）✅
- 出面×単価で金額、請負は対象外 → Task 5（`renderBillingTable` の `isOuen`）✅
- 経費は応援のみ・現場ごと1欄 → Task 5（`firstOfSite` 制御）✅
- 関数入りExcel出力 → Task 3（`=C*D`, `=SUM`）＋ Task 6 ✅
- 第1段階のみ（帳票・請負中身は除外）✅

**型整合性:**
- 明細オブジェクトのキー `{genba,loc,name,manDays,rate,method,expense}` を Task 5/6（admin送信）と Task 3（GAS受信 `ln.loc/ln.name/ln.manDays/ln.rate/ln.method/ln.expense`）で一致 ✅
- 単価キー `genba|||loc|||name` を `billingRate`/`onBillingRateInput`/保存ループで統一 ✅
- GAS戻り値 `{base64, filename}` は Task 3 で定義、Task 6 で受信。Task 6 Step2 で既存 export と整合確認 ✅

**未確定（実装時に確認）:**
- `exportSheetAsXlsxBase64_` の戻り値が生base64か `data:` 付きかを Task 3 Step3 で確認し、admin の `atob` 前処理を合わせる。
- タブが6個になりスマホ幅で窮屈な場合、`.tab` のフォント/パディング調整（既存CSSに従う）。
