# きんでん請求書を事務タブから生成 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: superpowers:executing-plans。**Task 0（実機検証）が必須ゲート**。PASSしなければ Task 1 以降に進まず、ローカルPython（既存・実証済み）に留める。

**Goal:** admin の事務タブに「💴 きんでん請求書」を追加し、局を選び当月迄出来高を入れると、GASが各局の実ファイル(Drive)を土台にxlsxを生成して即ダウンロードできる。

**Architecture:** GAS が Drive のひな型xlsxを `Utilities.unzip` → シートXMLの対象セルだけ文字列置換 → `Utilities.zip` で再圧縮（画像/印影/数式/結合は無加工で温存）。きんでんアクションはPIN必須（Script Property、公開リポに出さない）。累計(前月迄)は「きんでん局マスタ」シートで管理。

**Tech Stack:** GAS(Apps Script V8, Utilities.unzip/zip) / admin.html / Google Sheets / Google Drive

**Spec:** `docs/superpowers/specs/2026-05-30-kinden-west-invoice-design.md`（★改訂版）

**de-risk 済み:** マルチエージェント調査（`gas-zip-fidelity-research`）で「unzip は getName がフルパス／zip はパスで階層復元／未編集バイナリはbyte温存」を tanaikech 実例で確認。判定 = safe_with_mitigations。下記 mitigations を厳守。

---

## 検証で確定した必須mitigations（コードに織込済み／厳守）
1. 編集するのは **テキストXMLパートのみ**（sheet1/sheet2/workbook）。**バイナリ(画像/EMF/VML/printerSettings.bin)は絶対に getDataAsString/文字列化しない**＝unzipが返したBlobをそのまま zip へ。
2. 編集したXMLは `Utilities.newBlob(xml, MimeType.XML, '<フルパス>')` で**フルパス名**を維持（`xl/worksheets/sheet1.xml`）。余計なトップ階層を足さない。
3. 全エントリを漏れなく書き戻す（`[Content_Types].xml`/`_rels`/`xl/media`/`xl/drawings`/`printerSettings.bin`/`calcChain` 等）。
4. 文字列セルは **inlineStr 化**（sharedStrings索引を踏まない）、数値は `<v>` 直書き、`& < >` をエスケープ。数式は `<f>` を残し `fullCalcOnLoad="1"`。
5. 出力Blobに `.setContentType(MimeType.MICROSOFT_EXCEL)`。
6. `Utilities.unzip` の 50MB 上限内（きんでんは数百KB＝問題なし）。

---

### Task 0 ★必須ゲート：GAS実機検証（これがPASSしないと先に進まない）

**前提（あなたの操作）:** clasp が別アカウントに戻っているので、まず正アカウントに入り直す。
```bash
clasp logout
clasp login          # ブラウザで「予定管理」オーナーのGoogleアカウントを選び許可
clasp list-scripts   # 予定管理 が見えること（見えない=アカウント違い）
```

- [ ] **Step 1: 検証アクションを一時追加（gas.js）**

検証用spikeは `_local/zip_test/Code.js` に完成済み。これを本番で1回試すため、`generate_billing_calc_xlsx` の直後に一時アクションを足す（検証後に削除）：

```javascript
    // 【一時・検証用】Driveのひな型1つをGASで往復し、結果をbase64で返す。検証後に削除。
    if (action === 'kinden_zip_test') {
      var f = DriveApp.getFileById(String(body.fileId));   // 道頓堀ひな型のDrive ID
      var filled = fillKindenXlsx_(f.getBlob(), Number(body.made), Number(body.zen), String(body.ym));
      return ok({base64: Utilities.base64Encode(filled.getBytes()), filename: 'verify.xlsx'});
    }
```
そして `exportSheetAsXlsxBase64_` の直前に、`_local/zip_test/Code.js` の `kindenStyleOf_`/`kindenSetCell_`/`kindenFullCalc_`/`fillKindenXlsx_` をコピーして追加（※`fillKindenXlsx_` の最後を `.setContentType(MimeType.MICROSOFT_EXCEL)` 付きに）。

- [ ] **Step 2: 道頓堀ひな型をDriveに置き、IDを得る**

`_local/きんでん/template.xlsx` をDriveにアップロードし、ファイルIDを控える。

- [ ] **Step 3: デプロイ＆検証実行**

```bash
# 予定管理へ push & deploy（既存デプロイID更新・v39）
# 既存手順：clasp clone→cp gas.js コード.js→clasp push --force→clasp deploy -i AKfycbxp2eUc...
curl -sL -H "Content-Type: text/plain" \
  --data '{"action":"kinden_zip_test","fileId":"<DRIVE_ID>","made":2200000,"zen":1700000,"ym":"2026-06"}' \
  "https://script.google.com/macros/s/AKfycbxp2eUcpIjCj0ZWyAPPD9m3egJrKdWmXRK2AVnFrmBm4iO1QHCk-FZEH5LFFv7OloqcjQ/exec" -o "_local/verify_resp.json"
python -c "import json,base64;d=json.load(open(r'_local/verify_resp.json'));open(r'_local/きんでん/out/verify.xlsx','wb').write(base64.b64decode(d['base64']));print('saved')"
```

- [ ] **Step 4: ★判定（Excelで目視）**

`_local/きんでん/out/verify.xlsx` を**Excelで開く**。
- **PASS条件**：(1)修復ダイアログが出ない (2)印影/ロゴ画像が残っている (3)当月出来高=500,000・消費税・合計が計算される (4)結合セル・体裁が崩れていない。
- **PASS** → Task 1 へ進む（GAS方式で本実装）。
- **FAIL（修復ダイアログ等）** → **ここで中止**。GAS方式は不採用、ローカルPython（`tools/kinden/`）＝依頼ベースを正式手段として継続。一時アクションは削除。

---

### Task 1: GAS — きんでん局マスタ＋PIN＋本アクション（Task0 PASS後）

**Files:** Modify `gas.js`

- [ ] **Step 1: 定数とマスタシート**

```javascript
const KINDEN_MASTER_SHEET = 'きんでん局マスタ';
function getOrCreateKindenMasterSheet_(ss){
  var sh = ss.getSheetByName(KINDEN_MASTER_SHEET);
  if(!sh){ sh = ss.insertSheet(KINDEN_MASTER_SHEET); sh.appendRow(['局名','DriveファイルID','注文金額','前月迄出来高']); }
  return sh;
}
```

- [ ] **Step 2: PINゲートと list / generate アクション**

`doPost` 内に追加（PINは Script Property `KINDEN_PIN`、公開リポに出さない）：

```javascript
    if (action === 'kinden_list' || action === 'kinden_generate') {
      var pin = PropertiesService.getScriptProperties().getProperty('KINDEN_PIN');
      if (!pin || String(body.pin||'') !== pin) return error('PINが違います');
      var sh = getOrCreateKindenMasterSheet_(ss);
      var data = sh.getDataRange().getValues();
      if (action === 'kinden_list') {
        var list = [];
        for (var i=1;i<data.length;i++){ if(!String(data[i][0]).trim()) continue;
          list.push({name:String(data[i][0]), order:Number(data[i][2]||0), zen:Number(data[i][3]||0)}); }
        return ok({kyoku:list});
      }
      // generate
      var name = String(body.kyoku||''), made = Number(body.made||0), ym = String(body.ym||'');
      for (var r=1;r<data.length;r++){
        if (String(data[r][0]).trim() === name){
          var fileId = String(data[r][1]).trim();
          var zen = Number(data[r][3]||0);
          var blob = DriveApp.getFileById(fileId).getBlob();
          var filled = fillKindenXlsx_(blob, made, zen, ym);
          sh.getRange(r+1, 4).setValue(made);  // 前月迄を当月迄に更新
          logOperation_(ss, 'kinden_generate', name, ym+' 当月迄='+made, body.updatedBy);
          return ok({base64: Utilities.base64Encode(filled.getBytes()),
                     filename: 'きんでん_'+name+'_'+ym+'.xlsx'});
        }
      }
      return error('局マスタに該当局がありません: '+name);
    }
```

- [ ] **Step 3: Commit**（push/deployはTask4でまとめて）

```bash
git add gas.js && git commit -m "feat(gas): きんでん請求書アクション（局マスタ＋PIN＋zip生成）"
```

---

### Task 2: admin.html — 事務タブ「💴 きんでん請求書」UI

**Files:** Modify `admin.html`（事務画面 `screen-jimu` 内のカード群に追加。請求計算タブとは別、事務タブ内のカードでよい）

- [ ] **Step 1: HTMLカードを追加**（`screen-jimu` 内の適所）

```html
<div class="card" id="kinden-card">
  <div style="font-weight:600;margin-bottom:6px">💴 きんでん請求書</div>
  <div id="kinden-locked">
    <input type="password" id="kinden-pin" placeholder="PIN" style="width:120px">
    <button class="btn btn-sm" onclick="kindenUnlock()">解除</button>
  </div>
  <div id="kinden-body" style="display:none">
    <select id="kinden-kyoku" style="margin:6px 0"></select>
    <div><label>当月迄出来高（きんでん指示）</label>
      <input type="number" id="kinden-made" style="width:140px"></div>
    <div id="kinden-zan" style="font-size:12px;color:#666;margin:4px 0"></div>
    <input type="month" id="kinden-ym" style="margin:6px 0">
    <button class="btn" id="kinden-gen-btn" onclick="kindenGenerate()" style="background:#1D9E75">請求書を作成（ダウンロード）</button>
  </div>
</div>
```

- [ ] **Step 2: JS（PIN解除・一覧・生成→DL）**

請求計算JSの近くに追加。`KINDEN_PIN` は端末に保存せずセッション変数で保持：

```javascript
let kindenPin='', kindenList=[];
async function kindenUnlock(){
  const pin=document.getElementById('kinden-pin').value;
  if(!pin)return;
  try{
    const res=await fetch(getGasUrl(),{method:'POST',body:JSON.stringify({action:'kinden_list',pin}),headers:{'Content-Type':'text/plain'}});
    const j=await res.json();
    if(j.status!=='ok'){alert(j.message||'PINが違います');return;}
    kindenPin=pin; kindenList=j.kyoku||[];
    document.getElementById('kinden-locked').style.display='none';
    document.getElementById('kinden-body').style.display='block';
    const sel=document.getElementById('kinden-kyoku');
    sel.innerHTML=kindenList.map(k=>`<option value="${esc(k.name)}">${esc(k.name)}（残 ${(k.order-k.zen).toLocaleString()}円）</option>`).join('');
    kindenUpdateZan();
    sel.onchange=kindenUpdateZan;
  }catch(e){alert('通信エラー');}
}
function kindenUpdateZan(){
  const name=document.getElementById('kinden-kyoku').value;
  const k=kindenList.find(x=>x.name===name);
  document.getElementById('kinden-zan').textContent=k?`注文 ${k.order.toLocaleString()}／前月迄 ${k.zen.toLocaleString()}／残 ${(k.order-k.zen).toLocaleString()}円`:'';
}
async function kindenGenerate(){
  const kyoku=document.getElementById('kinden-kyoku').value;
  const made=Number(document.getElementById('kinden-made').value||0);
  const ym=document.getElementById('kinden-ym').value;
  if(!kyoku||!made||!ym){alert('局・当月迄出来高・年月を入れてください');return;}
  const btn=document.getElementById('kinden-gen-btn'); btn.disabled=true; btn.textContent='作成中...';
  try{
    const res=await fetch(getGasUrl(),{method:'POST',body:JSON.stringify({action:'kinden_generate',pin:kindenPin,kyoku,made,ym,updatedBy:getUsername()}),headers:{'Content-Type':'text/plain'}});
    const j=await res.json();
    if(j.status!=='ok'){alert('エラー：'+j.message);return;}
    const bin=atob(j.base64); const bytes=new Uint8Array(bin.length);
    for(let i=0;i<bin.length;i++)bytes[i]=bin.charCodeAt(i);
    const blob=new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=j.filename; document.body.appendChild(a); a.click(); a.remove();
    const k=kindenList.find(x=>x.name===kyoku); if(k)k.zen=made; kindenUpdateZan();
  }catch(e){alert('通信エラー');}
  finally{btn.disabled=false; btn.textContent='請求書を作成（ダウンロード）';}
}
```

- [ ] **Step 3: Commit**

```bash
git add admin.html && git commit -m "feat(admin): 事務タブにきんでん請求書UI（PIN・局選択・生成DL）"
```

---

### Task 3: あなたの一度きりの有効化（GASエディタ/Drive/シート）

- [ ] **Step 1: Driveにひな型を置く** — フォルダ「きんでん請求ひな型」を作り、6局のファイルをアップロード。各ファイルのID（URLの /d/〜）を控える。
- [ ] **Step 2: 局マスタを記入** — スプレッドシートの「きんでん局マスタ」シートに、局名／DriveファイルID／注文金額／前月迄出来高 を6行（値は `_local/きんでん/局マスタ.json` と同じ。道頓堀=template.xlsxのID）。
- [ ] **Step 3: PIN設定** — Apps Scriptエディタ → プロジェクトの設定 → スクリプト プロパティ → `KINDEN_PIN` に好きな合言葉。

---

### Task 4: デプロイ＆通し確認

- [ ] **Step 1: 一時検証アクション(kinden_zip_test)を削除**
- [ ] **Step 2: v39 デプロイ**（既存ID更新）＆ push
- [ ] **Step 3: 実機** — admin 事務タブ→💴→PIN→局選択→当月迄入力→生成→DL→**Excelで体裁OK確認**
- [ ] **Step 4: 引き継ぎ更新**

---

## フォールバック
Task 0 がFAIL、または途中で詰まったら、**ローカルPython（`tools/kinden/`・実証済み17テストPASS）＝依頼ベースを正式手段として継続**。きんでん請求は今日時点でPythonで問題なく出せる。GAS事務タブ化は「あれば便利」の上乗せ。

## セキュリティ注記
- 予定表のGAS URLは公開リポに載る＝実質公開。きんでんアクションは **PIN必須**＋ひな型は **Drive(非公開)**。PINは共有秘密で強固ではないが「素で誰でも取れる」は防ぐ。より堅くするなら別途認証を検討（次段階）。
