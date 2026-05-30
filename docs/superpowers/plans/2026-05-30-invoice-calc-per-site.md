# 請求計算タブ「現場ごと」化（第1段階・改）Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans (このプロジェクトは自動テスト無し。各タスクは編集→手動確認→コミット)。Steps use checkbox (`- [ ]`) syntax.

**Goal:** admin の「💰請求計算」タブを「現場×職人」から「現場ごと」に作り直す。応援＝延べ出面×現場単価＋経費、請負＝今回請求額の手入力。単価は元請×現場で記憶。

**Architecture:** 単価マスタのキーから職人を外し（元請×現場）、出面を現場単位で合算。請負現場は出面・単価計算をやめ金額手入力に。Excel出力は現場ごとの行で、応援は `=出面×単価`、請負は手入力値。既存 `exportSheetAsXlsxBase64_` で xlsx 化。

**Tech Stack:** Google Apps Script + Google Sheets / Vanilla JS + HTML（admin.html）

**Spec:** `docs/superpowers/specs/2026-05-30-invoice-calc-per-site-design.md`

---

## ファイル構成

| ファイル | 変更内容 |
|---|---|
| `gas.js` | ①`getOrCreateBillingRateSheet_` 4列化＋旧5列移行 ②`get_billing_rates`/`save_billing_rate` を元請×現場キーに ③`generate_billing_calc_xlsx` を現場ごと payload に |
| `admin.html` | ①テーブルヘッダを6列に（職人削除） ②グローバル `billingExp`/`billingAmt` 追加 ③`billingRate`/`getBillingLines` を現場集約に ④`renderBillingTable`＋ハンドラ群を方式分岐に ⑤`exportBillingXlsx` を現場ごと payload に |
| `index.html` | 変更なし |

検証は手動（GAS実行＋ブラウザ操作）。

---

### Task 1: gas.js — 単価マスタを「元請×現場」キーに

**Files:** Modify `gas.js`（`getOrCreateBillingRateSheet_` 行1124付近、`get_billing_rates` 行512、`save_billing_rate` 行524）

- [ ] **Step 1: `getOrCreateBillingRateSheet_` を4列化＋旧5列移行**

行1124〜1131 を以下に置き換える：

```javascript
function getOrCreateBillingRateSheet_(ss) {
  let sheet = ss.getSheetByName(BILLING_RATE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(BILLING_RATE_SHEET);
    sheet.appendRow(['元請名', '現場名', '単価', '更新日時']);
    return sheet;
  }
  // 旧5列（元請/現場/職人/単価/更新日時）からの移行：職人列があれば元請×現場へ畳む
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  if (String(headers[2] || '').trim() === '職人名') {
    const data = sheet.getDataRange().getValues();
    const map = {};
    for (let i = 1; i < data.length; i++) {
      const g = String(data[i][0] || '').trim();
      if (!g) continue;
      const l = String(data[i][1] || '').trim();
      map[g + '|||' + l] = { genba: g, loc: l, rate: Number(data[i][3] || 0), ts: String(data[i][4] || '') };
    }
    sheet.clear();
    sheet.appendRow(['元請名', '現場名', '単価', '更新日時']);
    Object.keys(map).forEach(k => { const x = map[k]; sheet.appendRow([x.genba, x.loc, x.rate, x.ts]); });
  }
  return sheet;
}
```

- [ ] **Step 2: `get_billing_rates` から職人を外す**

行512〜522 を以下に置き換える（単価は3列目＝`r[2]`）：

```javascript
    if (action === 'get_billing_rates') {
      const sheet = getOrCreateBillingRateSheet_(ss);
      const data = sheet.getDataRange().getValues();
      const rates = data.length > 1 ? data.slice(1).map(r => ({
        genba: String(r[0] || ''),
        loc: String(r[1] || ''),
        rate: Number(r[2] || 0)
      })).filter(x => x.genba) : [];
      return ok({rates: rates});
    }
```

- [ ] **Step 3: `save_billing_rate` を元請×現場キーに**

行524〜542 を以下に置き換える（単価＝3列目、更新日時＝4列目）：

```javascript
    if (action === 'save_billing_rate') {
      const sheet = getOrCreateBillingRateSheet_(ss);
      const genba = String(body.genba || '').trim();
      const loc = String(body.loc || '').trim();
      const rate = Number(body.rate || 0);
      if (!genba) return error('元請名は必須です');
      const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === genba && String(data[i][1]).trim() === loc) {
          sheet.getRange(i + 1, 3).setValue(rate);
          sheet.getRange(i + 1, 4).setValue(now);
          return ok({updated: true});
        }
      }
      sheet.appendRow([genba, loc, rate, now]);
      return ok({added: true});
    }
```

- [ ] **Step 4: 構文チェック**

Run: `python "C:/Users/akira/.claude/skills/gas-deploy/scripts/predeploy_check.py" "C:/Users/akira/OneDrive/Desktop/Claude/予定管理アプリ作成"`
Expected: PASS（既知WARN 2件＝VEHICLE_RES_TOKEN / manifest無しは許容）

- [ ] **Step 5: Commit**

```bash
cd "C:/Users/akira/OneDrive/Desktop/Claude/予定管理アプリ作成"
git add gas.js
git commit -m "feat(gas): 請求単価マスタを元請×現場キーに（職人を外す＋旧5列移行）"
```

---

### Task 2: gas.js — `generate_billing_calc_xlsx` を現場ごとに

**Files:** Modify `gas.js`（`generate_billing_calc_xlsx` 行560付近）

応援＝`=出面×単価`、請負＝手入力金額（数値）。列を「現場名/出面数/単価/金額/経費/方式」の6列に。

- [ ] **Step 1: アクション本体を置き換え**

`if (action === 'generate_billing_calc_xlsx') { ... }` ブロック（行560〜592）を以下に置き換える：

```javascript
    if (action === 'generate_billing_calc_xlsx') {
      const genba = String(body.genba || '');
      const month = String(body.month || '');
      const lines = Array.isArray(body.lines) ? body.lines : [];
      let sheet = ss.getSheetByName(BILLING_CALC_SHEET);
      if (sheet) { sheet.clear(); } else { sheet = ss.insertSheet(BILLING_CALC_SHEET); }
      sheet.appendRow([genba + '　' + month + '　請求計算']);
      sheet.appendRow(['現場名', '出面数', '単価', '金額', '経費', '方式']);
      const dataStart = 3; // 1=タイトル 2=ヘッダ 3=先頭データ
      lines.forEach(ln => {
        const r = sheet.getLastRow() + 1;
        const isOuen = String(ln.method || '応援') === '応援';
        if (isOuen) {
          // 金額 = 出面(B) × 単価(C)
          sheet.appendRow([
            String(ln.loc || ''), Number(ln.manDays || 0), Number(ln.rate || 0),
            '=B' + r + '*C' + r, Number(ln.expense || 0), '応援'
          ]);
        } else {
          // 請負：今回請求額を金額(D)に直接。出面/単価/経費は空
          sheet.appendRow([
            String(ln.loc || ''), '', '', Number(ln.amount || 0), 0, '請負'
          ]);
        }
      });
      const dataEnd = sheet.getLastRow();
      if (dataEnd >= dataStart) {
        const totalRow = dataEnd + 1;
        sheet.getRange(totalRow, 1).setValue('合計');
        sheet.getRange(totalRow, 4).setFormula('=SUM(D' + dataStart + ':D' + dataEnd + ')');
        sheet.getRange(totalRow, 5).setFormula('=SUM(E' + dataStart + ':E' + dataEnd + ')');
        sheet.getRange(totalRow + 1, 3).setValue('請求合計');
        sheet.getRange(totalRow + 1, 4).setFormula('=D' + totalRow + '+E' + totalRow);
      }
      SpreadsheetApp.flush();
      const result = exportSheetAsXlsxBase64_(ss, sheet);
      return ok({filename: '請求計算_' + genba + '_' + month + '.xlsx', base64: result.base64});
    }
```

- [ ] **Step 2: 構文チェック**

Run: `python "C:/Users/akira/.claude/skills/gas-deploy/scripts/predeploy_check.py" "C:/Users/akira/OneDrive/Desktop/Claude/予定管理アプリ作成"`
Expected: PASS

- [ ] **Step 3: Commit**

```bash
git add gas.js
git commit -m "feat(gas): 請求計算xlsxを現場ごとに（応援=出面×単価/請負=金額手入力）"
```

---

### Task 3: admin.html — ヘッダ6列化＋グローバル＋現場集約

**Files:** Modify `admin.html`（screen-billing の thead、`// ===== 請求計算 =====` のグローバル、`billingRate`、`getBillingLines`）

- [ ] **Step 1: テーブルヘッダを6列に（職人を削除）**

screen-billing 内の thead 行を置き換える：

old:
```html
      <thead><tr><th>現場</th><th>職人</th><th>出面</th><th>方式</th><th>単価</th><th>金額</th><th>経費</th></tr></thead>
```
new:
```html
      <thead><tr><th>現場</th><th>方式</th><th>出面</th><th>単価</th><th>金額</th><th>経費</th></tr></thead>
```

- [ ] **Step 2: グローバルに経費・請負金額マップを追加**

`let billingDirty = {};   // 画面で編集した単価 key->rate（保存用）` の直後に追加：

```javascript
let billingExp = {};     // 経費 loc->円（応援）
let billingAmt = {};     // 請負の今回請求額 loc->円
```

- [ ] **Step 3: `billingRate` から職人キーを外す**

`function billingRate(genba,loc,name){ ... }` を以下に置き換える：

```javascript
function billingRate(genba,loc){
  const key=genba+'|||'+loc;
  if(key in billingDirty)return billingDirty[key];
  const hit=billingRates.find(r=>r.genba===genba&&r.loc===loc);
  return hit?hit.rate:'';
}
```

- [ ] **Step 4: `getBillingLines` を現場集約に**

`function getBillingLines(){ ... }` を以下に置き換える（延べ出面＝(職人×日付)のユニーク数を現場ごとに合算）：

```javascript
function getBillingLines(){
  const genba=document.getElementById('bill-genba').value;
  const month=document.getElementById('bill-month').value;
  if(!genba||!month)return[];
  let nippos=filteredNippos().filter(n=>!n.isGhost&&n.genba===genba&&(n.date||'').startsWith(month)
    &&n.yakin!=='休み'&&n.yakin!=='予定'&&n.yakin!=='倉庫'&&!n.yasumi&&!n.yotei&&!n.souko);
  // 現場(loc)ごとに集約。延べ出面＝同一(職人,日付)を1としたユニーク数
  const map={};
  nippos.forEach(n=>{
    const loc=n.loc||'';
    if(!map[loc])map[loc]={loc:loc,days:new Set()};
    map[loc].days.add(n.name+'|||'+n.date);
  });
  return Object.keys(map).map(loc=>({
    genba:genba,loc:loc,manDays:map[loc].days.size,
    method:siteBillingMethod(genba,loc)
  })).sort((a,b)=>a.loc.localeCompare(b.loc));
}
```

- [ ] **Step 5: Commit**

```bash
git add admin.html
git commit -m "feat(admin): 請求計算を現場集約に（ヘッダ6列・単価キー元請×現場・延べ出面）"
```

---

### Task 4: admin.html — 描画と入力ハンドラを方式分岐に

**Files:** Modify `admin.html`（`renderBillingTable`、`onBillingRateInput`、`updateBillingTotal`、ハンドラ追加）

- [ ] **Step 1: `renderBillingTable` を方式分岐に置き換え**

`function renderBillingTable(){ ... }` を以下に置き換える：

```javascript
function renderBillingTable(){
  const genba=document.getElementById('bill-genba').value;
  const lines=getBillingLines();
  const wrap=document.getElementById('bill-table-wrap');
  if(lines.length===0){wrap.style.display='none';return;}
  wrap.style.display='block';
  document.getElementById('bill-tbody').innerHTML=lines.map(ln=>{
    const isOuen=ln.method==='応援';
    const methodSel=`<select class="bill-method" data-genba="${esc(genba)}" data-loc="${esc(ln.loc)}" onchange="onBillingMethodChange(this)">
      <option value="応援"${isOuen?' selected':''}>応援</option>
      <option value="請負"${!isOuen?' selected':''}>請負</option></select>`;
    if(isOuen){
      const rate=billingRate(genba,ln.loc);
      const amount=(rate!==''&&rate!=null)?(ln.manDays*Number(rate)):'';
      const rateInput=`<input type="number" class="bill-rate" data-genba="${esc(genba)}" data-loc="${esc(ln.loc)}" value="${rate===''?'':rate}" style="width:90px" oninput="onBillingRateInput(this)">`;
      const expVal=(ln.loc in billingExp)?billingExp[ln.loc]:'';
      const expInput=`<input type="number" class="bill-exp" data-loc="${esc(ln.loc)}" value="${expVal===''?'':expVal}" style="width:80px" oninput="onBillingExpInput(this)">`;
      return `<tr data-loc="${esc(ln.loc)}">
        <td>${esc(ln.loc)}</td><td>${methodSel}</td>
        <td style="text-align:right">${ln.manDays}</td>
        <td>${rateInput}</td>
        <td class="bill-amount" style="text-align:right">${amount===''?'-':amount.toLocaleString()}</td>
        <td>${expInput}</td></tr>`;
    }else{
      const amtVal=(ln.loc in billingAmt)?billingAmt[ln.loc]:'';
      const amtInput=`<input type="number" class="bill-amt" data-loc="${esc(ln.loc)}" value="${amtVal===''?'':amtVal}" style="width:120px" placeholder="今回請求額" oninput="onBillingAmtInput(this)">`;
      return `<tr data-loc="${esc(ln.loc)}">
        <td>${esc(ln.loc)}</td><td>${methodSel}</td>
        <td style="text-align:right;color:#bbb">—</td>
        <td style="color:#bbb">—</td>
        <td class="bill-amount" style="text-align:right">${amtInput}</td>
        <td style="color:#bbb">—</td></tr>`;
    }
  }).join('');
  updateBillingTotal();
}
```

- [ ] **Step 2: `onBillingRateInput` を元請×現場キーに（行ごと金額更新は維持）**

`function onBillingRateInput(el){ ... }` を以下に置き換える（出面は3列目＝`tr.children[2]`）：

```javascript
function onBillingRateInput(el){
  const key=el.dataset.genba+'|||'+el.dataset.loc;
  billingDirty[key]=el.value===''?'':Number(el.value);
  const tr=el.closest('tr');
  if(tr){
    const manDays=Number(tr.children[2].textContent)||0;
    const amountCell=tr.querySelector('.bill-amount');
    const amount=(el.value!==''&&el.value!=null)?(manDays*Number(el.value)):'';
    if(amountCell)amountCell.textContent=(amount===''?'-':amount.toLocaleString());
  }
  updateBillingTotal();
}
```

- [ ] **Step 3: 経費・請負金額のハンドラを追加**

Step 2 の `onBillingRateInput` の直後に追加：

```javascript
function onBillingExpInput(el){
  billingExp[el.dataset.loc]=el.value===''?'':Number(el.value);
  updateBillingTotal();
}
function onBillingAmtInput(el){
  billingAmt[el.dataset.loc]=el.value===''?'':Number(el.value);
  updateBillingTotal();
}
```

- [ ] **Step 4: `updateBillingTotal` を応援(計算)＋請負(手入力)対応に**

`function updateBillingTotal(){ ... }` を以下に置き換える：

```javascript
function updateBillingTotal(){
  let amountSum=0, expSum=0;
  document.querySelectorAll('#bill-tbody tr').forEach(tr=>{
    const cell=tr.querySelector('.bill-amount');
    if(!cell)return;
    const amtInput=cell.querySelector('.bill-amt'); // 請負＝手入力
    if(amtInput){ if(amtInput.value)amountSum+=Number(amtInput.value); }
    else { const a=cell.textContent.replace(/[^0-9]/g,''); if(a)amountSum+=Number(a); } // 応援＝計算結果テキスト
  });
  document.querySelectorAll('.bill-exp').forEach(inp=>{if(inp.value)expSum+=Number(inp.value);});
  document.getElementById('bill-total').textContent=
    `金額合計 ${amountSum.toLocaleString()}円 ＋ 経費 ${expSum.toLocaleString()}円 ＝ 請求合計 ${(amountSum+expSum).toLocaleString()}円`;
}
```

- [ ] **Step 5: インラインJS構文チェック**

Run:
```bash
node -e 'const fs=require("fs");const h=fs.readFileSync("admin.html","utf8");const re=/<script\b(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/gi;let m,i=0,bad=0;while((m=re.exec(h))){i++;try{new Function(m[1]);}catch(e){bad++;console.log("ERR#"+i+": "+e.message);}}console.log("blocks:"+i+" errors:"+bad);'
```
Expected: `blocks:1 errors:0`

- [ ] **Step 6: Commit**

```bash
git add admin.html
git commit -m "feat(admin): 請求計算の描画を方式分岐に（応援=出面×単価/請負=金額手入力＋合計）"
```

---

### Task 5: admin.html — Excel出力を現場ごと payload に

**Files:** Modify `admin.html`（`exportBillingXlsx`）

- [ ] **Step 1: `exportBillingXlsx` を置き換え**

`async function exportBillingXlsx(){ ... }` を以下に置き換える：

```javascript
async function exportBillingXlsx(){
  const genba=document.getElementById('bill-genba').value;
  const month=document.getElementById('bill-month').value;
  if(!genba||!month){alert('元請と月を選んでください');return;}
  const btn=document.getElementById('bill-export-btn');
  btn.disabled=true;btn.textContent='出力中...';
  try{
    // 1) 編集した単価をマスタ保存（元請×現場）
    for(const key of Object.keys(billingDirty)){
      const [g,l]=key.split('|||');
      const rate=billingDirty[key];
      if(rate==='')continue;
      await fetch(getGasUrl(),{method:'POST',body:JSON.stringify({action:'save_billing_rate',genba:g,loc:l,rate:Number(rate),updatedBy:getUsername()}),headers:{'Content-Type':'text/plain'}});
    }
    // 2) 明細を現場ごとに（応援＝出面×単価＋経費／請負＝今回請求額）
    const lines=getBillingLines().map(ln=>{
      if(ln.method==='応援'){
        const rate=billingRate(genba,ln.loc);
        return {loc:ln.loc,method:'応援',manDays:ln.manDays,rate:(rate===''?0:Number(rate)),expense:Number(billingExp[ln.loc]||0)};
      }else{
        return {loc:ln.loc,method:'請負',amount:Number(billingAmt[ln.loc]||0)};
      }
    });
    // 3) xlsx生成リクエスト
    const res=await fetch(getGasUrl(),{method:'POST',body:JSON.stringify({action:'generate_billing_calc_xlsx',genba,month,lines}),headers:{'Content-Type':'text/plain'}});
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
    // 5) 保存済み単価をローカルにも反映
    for(const key of Object.keys(billingDirty)){const[g,l]=key.split('|||');const hit=billingRates.find(r=>r.genba===g&&r.loc===l);if(hit)hit.rate=Number(billingDirty[key]);else billingRates.push({genba:g,loc:l,rate:Number(billingDirty[key])});}
    billingDirty={};
  }catch(e){alert('出力エラー：'+e.message);}
  finally{btn.disabled=false;btn.textContent='Excelで出力（関数入り）';}
}
```

- [ ] **Step 2: インラインJS構文チェック**

Run: 同 Task 4 Step 5 のコマンド
Expected: `blocks:1 errors:0`

- [ ] **Step 3: Commit & Push**

```bash
git add admin.html
git commit -m "feat(admin): 請求計算Excel出力を現場ごとに（応援=出面×単価/請負=金額）"
git push origin main
```

---

### Task 6: デプロイ（v37）＆検証＆引き継ぎ

- [ ] **Step 1: pre-deploy チェック**

Run: `python "C:/Users/akira/.claude/skills/gas-deploy/scripts/predeploy_check.py" "C:/Users/akira/OneDrive/Desktop/Claude/予定管理アプリ作成"`
Expected: PASS

- [ ] **Step 2: GAS v37 デプロイ（clasp）**

clasp は前回正アカウントでログイン済み・Apps Script API 有効。手順：
```bash
mkdir -p "_local/gas_deploy_tmp"
(cd "_local/gas_deploy_tmp" && clasp clone "1BXSKkYbrU4nhuFVi_YsujzP19zMpBHmG_xxMhysUp-365yg0BaeSMV5t")
cp "gas.js" "_local/gas_deploy_tmp/コード.js"
(cd "_local/gas_deploy_tmp" && clasp push --force)
(cd "_local/gas_deploy_tmp" && clasp deploy -i "AKfycbxp2eUcpIjCj0ZWyAPPD9m3egJrKdWmXRK2AVnFrmBm4iO1QHCk-FZEH5LFFv7OloqcjQ" -d "v37 請求計算を現場ごとに")
rm -rf "_local/gas_deploy_tmp"
```
Expected: `Deployed AKfycbxp2eUc… @37`。デプロイID/URL不変。

- [ ] **Step 3: 本番スモークテスト（curl）**

Run: `curl -sL -H "Content-Type: text/plain" --data '{"action":"get_billing_rates"}' "https://script.google.com/macros/s/AKfycbxp2eUcpIjCj0ZWyAPPD9m3egJrKdWmXRK2AVnFrmBm4iO1QHCk-FZEH5LFFv7OloqcjQ/exec"`
Expected: `{"status":"ok","rates":[...]}`（rates の各要素に `name` が無く `genba/loc/rate` のみ）

- [ ] **Step 4: 手動UI検証**

1. admin を Ctrl+Shift+R →「💰請求計算」→ 元請＋月
2. **現場ごとに1行**（職人ごとに並ばない）。応援現場は延べ出面が出る
3. 単価入力 → 金額＝出面×単価が即計算（フォーカス維持）、合計更新
4. 方式「請負」に切替 → 出面/単価/経費が「—」、金額欄が手入力できる
5. 請負の金額を入力 → 請求合計に入る
6. 経費（応援）→ 合計反映
7. Excel出力 → 応援行は出面/単価を変えると金額・合計が自動再計算、請負行は手入力額
8. 再読込 → 応援の単価が自動表示（元請×現場で記憶）。請求単価マスタが4列（職人なし）

- [ ] **Step 5: 引き継ぎ更新**

`引き継ぎ.md` 冒頭に「請求計算 現場ごと化（第1段階・改）＋GAS v37」を追記し commit & push。

---

## 自己レビュー（Spec カバレッジ）

- 単価を元請×現場で記憶 → Task 1 ✅
- 現場ごと集約・延べ出面 → Task 3 Step4 ✅
- 応援＝出面×単価＋経費 → Task 4（renderの応援分岐）＋ Task 2（=B*C）✅
- 請負＝金額手入力・出面/単価/経費は無効 → Task 4（請負分岐）＋ Task 2（amount数値）✅
- Excel関数入り（応援）／請負は数値 → Task 2 ✅
- 職人キー廃止 → Task 1/3 ✅
- 部分検収の書式生成・請負進捗は範囲外（次段階）✅

**型整合性:**
- payload 応援 `{loc,method,manDays,rate,expense}` / 請負 `{loc,method,amount}` を Task 5（送信）と Task 2（受信 `ln.manDays/ln.rate/ln.expense/ln.amount`）で一致 ✅
- 単価キー `genba|||loc` を `billingRate`/`onBillingRateInput`/保存ループで統一 ✅
- 列インデックス：出面=テーブル3列目（`tr.children[2]`）、xlsx 金額=D列・経費=E列 ✅
