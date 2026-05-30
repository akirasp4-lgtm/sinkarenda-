# きんでん西 指定請求書ジェネレーター Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans（自動テストは pytest 不要の素のassertスクリプト）。Steps use checkbox (`- [ ]`) syntax.

**Goal:** きんでんの指定xlsxをひな型に、局で変わるセル値だけを直接書き換えて、局ごとに完成請求書を出力する Python ツール。

**Architecture:** xlsx(zip)内のシートXML(`sheet1`=表紙 / `sheet2`=請求書)の対象セルの`<v>`/inlineStringだけを置換し、画像・印影・スタイル・数式は触らない。`workbook.xml`に`fullCalcOnLoad`を立てExcel再計算を保証。累計(前月迄出来高)は局マスタJSONに記憶。

**Tech Stack:** Python 3（標準ライブラリ zipfile/re/json/datetime のみ。生成はopenpyxl不使用＝画像保全。検証読み取りにopenpyxl）

**Spec:** `docs/superpowers/specs/2026-05-30-kinden-west-invoice-design.md`

---

## ファイル構成

| パス | 役割 | 公開 |
|---|---|---|
| `tools/kinden/kinden_invoice.py` | ジェネレーター本体（穴埋め・累計・CLI） | commit（コードのみ・機密なし） |
| `tools/kinden/test_kinden_invoice.py` | 自動テスト（素のassert） | commit |
| `_local/きんでん/template.xlsx` | きんでん指定ひな型（実ファイルのコピー） | **非公開**（_local） |
| `_local/きんでん/局マスタ.json` | 局ごとの固定情報＋前月迄累計 | **非公開** |
| `_local/きんでん/out/` | 出力xlsx（局ごと1枚） | **非公開** |

> ひな型と局マスタ・出力は `_local/`（gitignore）に置き、公開リポに上げない。コードのみ commit。

---

### Task 1: スキャフォールド（フォルダ＋ひな型＋局マスタ）

**Files:** Create `_local/きんでん/`（手動・非公開）, `tools/kinden/`（commit）

- [ ] **Step 1: フォルダとひな型を用意**

```bash
cd "C:/Users/akira/OneDrive/Desktop/Claude/予定管理アプリ作成"
mkdir -p "_local/きんでん/out" "tools/kinden"
cp "_local/kinden_seikyu.xlsx" "_local/きんでん/template.xlsx"
ls "_local/きんでん"
```
Expected: `template.xlsx` と `out/` が見える。

- [ ] **Step 2: 局マスタJSONを作成（サンプル＝道頓堀。値はひな型から）**

Create `_local/きんでん/局マスタ.json`:

```json
{
  "局": {
    "ＩＭＴ道頓堀": {
      "見積No": "K00386",
      "見積日": "2026-05-01",
      "注文金額": 2530000,
      "労務費": 1289560,
      "工事名": "Ｒ００４７３０１２０　支障移転廃局側　ＩＭＴ道頓堀",
      "施工場所": "大阪府 大阪市中央区",
      "着工": "2026-05-01",
      "完成": "2026-05-15",
      "工番中": "5",
      "工番末": "1317",
      "前月迄出来高": 0
    }
  }
}
```

- [ ] **Step 3: Commit（コード側フォルダのみ。_localは無視される）**

```bash
git add tools/kinden 2>/dev/null; git commit -q -m "chore(kinden): ジェネレーター用フォルダを用意" --allow-empty
```

---

### Task 2: コア穴埋め関数（セル値の直接置換）

**Files:** Create `tools/kinden/kinden_invoice.py`, `tools/kinden/test_kinden_invoice.py`

- [ ] **Step 1: 置換のコア関数を実装**

Create `tools/kinden/kinden_invoice.py`:

```python
# -*- coding: utf-8 -*-
"""きんでん指定請求書ジェネレーター（xlsx内セル値の直接書換・画像/体裁を温存）。"""
import zipfile, re, json, calendar
from datetime import date

S1 = 'xl/worksheets/sheet1.xml'   # きんでん情通C指定見積表紙
S2 = 'xl/worksheets/sheet2.xml'   # 請求書(インボイス）
WB = 'xl/workbook.xml'

def _xesc(s):
    return (str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))

def _style_of(xml, ref):
    m = re.search(r'<c r="' + re.escape(ref) + r'"([^>]*?)(?:/>|>)', xml)
    if not m:
        raise ValueError('cell not found: ' + ref)
    sm = re.search(r's="(\d+)"', m.group(1))
    return sm.group(1) if sm else None

def set_cell(xml, ref, kind, value):
    """kind: 'n'(数値) / 's'(文字列). 対象セルの<c>要素を丸ごと置換（styleは維持）。"""
    s = _style_of(xml, ref)
    sattr = (' s="%s"' % s) if s else ''
    if kind == 'n':
        new = '<c r="%s"%s><v>%s</v></c>' % (ref, sattr, value)
    else:
        new = '<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, sattr, _xesc(value))
    pat = re.compile(r'<c r="' + re.escape(ref) + r'"[^>]*?(?:/>|>.*?</c>)', re.S)
    out, n = pat.subn(new, xml, count=1)
    if n != 1:
        raise ValueError('replace failed for %s (n=%d)' % (ref, n))
    return out

def enable_full_calc(wbxml):
    if 'calcPr' in wbxml:
        if 'fullCalcOnLoad' in wbxml:
            return re.sub(r'fullCalcOnLoad="[^"]*"', 'fullCalcOnLoad="1"', wbxml, count=1)
        return re.sub(r'<calcPr', '<calcPr fullCalcOnLoad="1"', wbxml, count=1)
    return re.sub(r'(<workbookPr[^>]*/>)', r'\1<calcPr fullCalcOnLoad="1"/>', wbxml, count=1)
```

- [ ] **Step 2: コア関数のテストを書く**

Create `tools/kinden/test_kinden_invoice.py`:

```python
# -*- coding: utf-8 -*-
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
import kinden_invoice as K

p = f = 0
def check(name, cond):
    global p, f
    print(('PASS' if cond else 'FAIL') + ' : ' + name); p += (1 if cond else 0); f += (0 if cond else 1)

# 数値セル: 既存<v>を差し替え、styleは維持
xml = '<row r="8"><c r="G8" s="264"><v>2530000</v></c></row>'
o = K.set_cell(xml, 'G8', 'n', 9999)
check('数値セル置換', '<c r="G8" s="264"><v>9999</v></c>' in o)

# 空の自己終了セルに数値を入れる
xml2 = '<row r="18"><c r="AO18" s="102"/></row>'
o2 = K.set_cell(xml2, 'AO18', 'n', 500000)
check('空セルに数値', '<c r="AO18" s="102"><v>500000</v></c>' in o2)

# 文字列セル: inlineString化（共有文字列を触らない）、エスケープ
xml3 = '<row r="28"><c r="G28" s="186" t="s"><v>155</v></c></row>'
o3 = K.set_cell(xml3, 'G28', 's', 'A&B<工事>')
check('文字列inlineString化', 't="inlineStr"' in o3 and 'A&amp;B&lt;工事&gt;' in o3 and 's="186"' in o3)

# fullCalcOnLoad
check('calcPrにfullCalc付与', 'fullCalcOnLoad="1"' in K.enable_full_calc('<calcPr calcId="1"/>'))

print('\n==== %d passed, %d failed ====' % (p, f)); sys.exit(1 if f else 0)
```

- [ ] **Step 3: テスト実行**

Run: `cd "C:/Users/akira/OneDrive/Desktop/Claude/予定管理アプリ作成" && python tools/kinden/test_kinden_invoice.py`
Expected: `4 passed, 0 failed`

- [ ] **Step 4: Commit**

```bash
git add tools/kinden/kinden_invoice.py tools/kinden/test_kinden_invoice.py
git commit -q -m "feat(kinden): セル値直接置換のコア関数＋テスト"
```

---

### Task 3: 日付・工番・1局生成

**Files:** Modify `tools/kinden/kinden_invoice.py`, `tools/kinden/test_kinden_invoice.py`

- [ ] **Step 1: 日付分割と1局生成を実装**

`kinden_invoice.py` の末尾に追加：

```python
# 着工/完成の年(4桁)・月(2桁)・日(2桁)を1桁ずつ入れるセル群
START_CELLS = {'y': ['G32', 'H32', 'I32', 'J32'], 'm': ['L32', 'M32'], 'd': ['O32', 'P32']}
DONE_CELLS  = {'y': ['W32', 'X32', 'Y32', 'Z32'], 'm': ['AB32', 'AC32'], 'd': ['AE32', 'AF32']}

def _digits(val, width):
    s = str(int(val)).zfill(width)
    return [int(ch) for ch in s[-width:]]

def _set_date(xml, cells, ymd):
    y, m, d = ymd.year, ymd.month, ymd.day
    for grp, width, num in (('y', 4, y), ('m', 2, m), ('d', 2, d)):
        for cell, dig in zip(cells[grp], _digits(num, width)):
            xml = K_set(xml, cell, 'n', dig)
    return xml

# set_cell をモジュール内から呼ぶ別名（_set_date 用）
def K_set(xml, ref, kind, value):
    return set_cell(xml, ref, kind, value)

def _d(s):
    y, m, dd = (int(x) for x in s.split('-'))
    return date(y, m, dd)

def fill_workbook(template_path, out_path, kyoku, toushin_made, zen_made, seikyu_ym):
    """kyoku=局マスタの1局dict, toushin_made=当月迄出来高, zen_made=前月迄出来高, seikyu_ym='YYYY-MM'."""
    sy, sm = (int(x) for x in seikyu_ym.split('-'))
    last_day = calendar.monthrange(sy, sm)[1]
    mitsumori = _d(kyoku['見積日'])
    zin = zipfile.ZipFile(template_path, 'r')
    s1 = zin.read(S1).decode('utf-8')
    s2 = zin.read(S2).decode('utf-8')
    wb = zin.read(WB).decode('utf-8')
    # --- 表紙 ---
    s1 = set_cell(s1, 'AB4', 's', kyoku['見積No'])
    s1 = set_cell(s1, 'AA5', 'n', sy)       # 請求年（請求書が参照）
    s1 = set_cell(s1, 'AE5', 'n', sm)       # 請求月
    s1 = set_cell(s1, 'AH5', 'n', mitsumori.day)
    s1 = set_cell(s1, 'G8', 'n', kyoku['注文金額'])
    s1 = set_cell(s1, 'H14', 'n', kyoku['労務費'])
    s1 = set_cell(s1, 'G28', 's', kyoku['工事名'])
    s1 = set_cell(s1, 'G30', 's', kyoku['施工場所'])
    s1 = _set_date(s1, START_CELLS, _d(kyoku['着工']))
    s1 = _set_date(s1, DONE_CELLS, _d(kyoku['完成']))
    s1 = set_cell(s1, 'M54', 's', kyoku['工番中'])
    s1 = set_cell(s1, 'P54', 's', kyoku['工番末'])
    # --- 請求書 ---
    s2 = set_cell(s2, 'AO17', 'n', toushin_made)   # 当月迄出来高(A)
    s2 = set_cell(s2, 'AO18', 'n', zen_made)       # 前月迄出来高(B)
    s2 = set_cell(s2, 'CO1', 'n', last_day)        # 請求日(末日)
    wb = enable_full_calc(wb)
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == S1: data = s1.encode('utf-8')
            elif it.filename == S2: data = s2.encode('utf-8')
            elif it.filename == WB: data = wb.encode('utf-8')
            zout.writestr(it, data)
    zin.close()
```

- [ ] **Step 2: 1局生成のテストを追加（ひな型が要る／openpyxlで値と数式と画像を検証）**

`test_kinden_invoice.py` の `print('\n==== ...` の直前に追加：

```python
TEMPLATE = os.path.join(os.path.dirname(__file__), '..', '..', '_local', 'きんでん', 'template.xlsx')
if os.path.exists(TEMPLATE):
    import openpyxl, zipfile as zf, tempfile
    out = os.path.join(tempfile.gettempdir(), 'kinden_test_out.xlsx')
    ky = {'見積No': 'K00999', '見積日': '2026-05-01', '注文金額': 3000000, '労務費': 1500000,
          '工事名': 'テスト工事', '施工場所': '大阪市西区', '着工': '2026-05-02', '完成': '2026-05-20',
          '工番中': '7', '工番末': '2222'}
    K.fill_workbook(TEMPLATE, out, ky, toushin_made=1800000, zen_made=600000, seikyu_ym='2026-06')
    wb = openpyxl.load_workbook(out, data_only=False)
    t = wb['きんでん情通C指定見積表紙']; r = wb['請求書(インボイス）']
    check('注文金額G8', t['G8'].value == 3000000)
    check('労務費H14', t['H14'].value == 1500000)
    check('法定福利費の数式が残る', str(t['H16'].value) == '=H14*H15')
    check('当月迄AO17', r['AO17'].value == 1800000)
    check('前月迄AO18', r['AO18'].value == 600000)
    check('当月出来高の数式が残る', str(r['AO19'].value) == '=AO17-AO18')
    check('請求日CO1=月末(6月=30)', r['CO1'].value == 30)
    check('工番末P54', str(t['P54'].value) == '2222')
    # 画像が温存されている（PNG×2/JPEG/EMF=4点）
    media = [n for n in zf.ZipFile(out).namelist() if '/media/' in n]
    check('画像4点が温存', len(media) == 4)
else:
    print('SKIP: ひな型が無いので生成テストはスキップ（_local/きんでん/template.xlsx を用意）')
```

- [ ] **Step 3: テスト実行**

Run: `cd "C:/Users/akira/OneDrive/Desktop/Claude/予定管理アプリ作成" && python tools/kinden/test_kinden_invoice.py`
Expected: コア4件＋生成9件＝`13 passed, 0 failed`（ひな型がある場合）

- [ ] **Step 4: Commit**

```bash
git add tools/kinden/kinden_invoice.py tools/kinden/test_kinden_invoice.py
git commit -q -m "feat(kinden): 日付分割・工番・1局生成（画像/数式温存を検証）"
```

---

### Task 4: 月次実行＋累計繰越（局マスタ読み書き）

**Files:** Modify `tools/kinden/kinden_invoice.py`

- [ ] **Step 1: 月次ランを実装**

`kinden_invoice.py` 末尾に追加：

```python
def run_month(master_path, template_path, out_dir, seikyu_ym, instructions):
    """instructions = {局名: 当月迄出来高}. 各局を生成し、前月迄累計を更新して保存。"""
    import os
    with open(master_path, encoding='utf-8') as fp:
        master = json.load(fp)
    results = []
    for name, toushin_made in instructions.items():
        ky = master['局'][name]
        zen = ky.get('前月迄出来高', 0)
        out = os.path.join(out_dir, 'きんでん_%s_%s.xlsx' % (name, seikyu_ym))
        fill_workbook(template_path, out, ky, toushin_made, zen, seikyu_ym)
        ky['前月迄出来高'] = toushin_made   # 次回の(B)
        results.append({'局': name, '当月出来高': toushin_made - zen, 'file': out})
    with open(master_path, 'w', encoding='utf-8') as fp:
        json.dump(master, fp, ensure_ascii=False, indent=2)
    return results
```

- [ ] **Step 2: 累計繰越のテストを追加（マスタの一時コピーで）**

`test_kinden_invoice.py` の最終 `print('\n==== ...` の直前に追加：

```python
MASTER = os.path.join(os.path.dirname(__file__), '..', '..', '_local', 'きんでん', '局マスタ.json')
if os.path.exists(TEMPLATE) and os.path.exists(MASTER):
    import json as _json, shutil, tempfile
    tmp_master = os.path.join(tempfile.gettempdir(), 'kinden_master_test.json')
    shutil.copy(MASTER, tmp_master)
    name = list(_json.load(open(MASTER, encoding='utf-8'))['局'].keys())[0]
    outdir = tempfile.gettempdir()
    K.run_month(tmp_master, TEMPLATE, outdir, '2026-05', {name: 1000000})
    after1 = _json.load(open(tmp_master, encoding='utf-8'))['局'][name]['前月迄出来高']
    check('1回目で前月迄が当月迄に更新', after1 == 1000000)
    K.run_month(tmp_master, TEMPLATE, outdir, '2026-06', {name: 1600000})
    after2 = _json.load(open(tmp_master, encoding='utf-8'))['局'][name]['前月迄出来高']
    check('2回目で累計繰越(1.6M)', after2 == 1600000)
else:
    print('SKIP: 累計テスト（ひな型/マスタ未設置）')
```

- [ ] **Step 3: テスト実行**

Run: `cd "C:/Users/akira/OneDrive/Desktop/Claude/予定管理アプリ作成" && python tools/kinden/test_kinden_invoice.py`
Expected: `15 passed, 0 failed`（ひな型＋マスタがある場合）

- [ ] **Step 4: Commit**

```bash
git add tools/kinden/kinden_invoice.py tools/kinden/test_kinden_invoice.py
git commit -q -m "feat(kinden): 月次実行＋前月迄累計の繰越"
```

---

### Task 5: CLI＋サンプル生成＋Excel実機確認

**Files:** Modify `tools/kinden/kinden_invoice.py`

- [ ] **Step 1: CLIを追加**

`kinden_invoice.py` 末尾に追加：

```python
if __name__ == '__main__':
    import argparse, os
    base = os.path.join(os.path.dirname(__file__), '..', '..', '_local', 'きんでん')
    ap = argparse.ArgumentParser(description='きんでん指定請求書ジェネレーター')
    ap.add_argument('--ym', required=True, help='請求年月 YYYY-MM')
    ap.add_argument('--kyoku', required=True, help='局名')
    ap.add_argument('--made', required=True, type=int, help='当月迄出来高（きんでん指示）')
    ap.add_argument('--base', default=base)
    a = ap.parse_args()
    res = run_month(os.path.join(a.base, '局マスタ.json'),
                    os.path.join(a.base, 'template.xlsx'),
                    os.path.join(a.base, 'out'),
                    a.ym, {a.kyoku: a.made})
    for r in res:
        print('生成: %s（当月出来高 %s円）' % (r['file'], format(r['当月出来高'], ',')))
```

- [ ] **Step 2: サンプル生成（道頓堀・5月分=全額一括の例）**

```bash
cd "C:/Users/akira/OneDrive/Desktop/Claude/予定管理アプリ作成"
python tools/kinden/kinden_invoice.py --ym 2026-05 --kyoku ＩＭＴ道頓堀 --made 2530000
ls "_local/きんでん/out"
```
Expected: `きんでん_ＩＭＴ道頓堀_2026-05.xlsx` が出力される。

- [ ] **Step 3: Excel実機確認（利用者）**

出力ファイルをExcelで開き、(1)体裁・罫線・印影がひな型と同じ (2)法定福利費・当月出来高・消費税・請求合計が正しく計算 (3)注文金額/工事名/工期/工番が正しい、を確認。

- [ ] **Step 4: Commit & Push**

```bash
git add tools/kinden/kinden_invoice.py
git commit -q -m "feat(kinden): CLI追加（局名＋当月迄出来高で1局生成）"
git push origin main
```

- [ ] **Step 5: 引き継ぎ更新**

`引き継ぎ.md` 冒頭にきんでん西ジェネレーターの完成と使い方（`python tools/kinden/kinden_invoice.py --ym ... --kyoku ... --made ...`／ひな型・局マスタは `_local/きんでん/`）を追記し commit & push。

---

## 自己レビュー

**Spec カバレッジ:**
- ひな型穴埋め・画像温存 → Task 2（zip-XML置換）＋ Task 3（生成テストで画像4点確認）✅
- 局ごとの固定情報＋累計 → Task 1（局マスタ）＋ Task 4（繰越）✅
- 穴埋めセル（見積No/金額/労務費/工事名/施工場所/工期/工番/当月迄/前月迄/請求日）→ Task 3 ✅
- 出来高分割・きんでん指示の金額・前月迄自動 → Task 4（instructions＝当月迄、zen＝マスタ）✅
- 法定福利費・出来高・消費税・合計は数式のまま＋fullCalcOnLoad → Task 2/3 ✅
- 1局＝1ファイル → Task 4（ファイル名 局＋年月）✅
- 機密は_local・コードのみcommit → ファイル構成表 ✅

**未確定（実装後に利用者と確認）:**
- 請求書の年月の扱い：本実装は「請求年月」を表紙AA5/AE5に入れCO1=月末。見積日(month)を固定したい運用なら要調整（初回出力をきんでんに確認）。
- 局マスタの初期登録（道頓堀以外の局）。値は各局の指定見積から転記。

**型整合性:**
- `fill_workbook(template, out, kyoku, toushin_made, zen_made, seikyu_ym)` の引数を Task 3定義・Task 4呼び出しで一致 ✅
- 局マスタのキー（見積No/見積日/注文金額/労務費/工事名/施工場所/着工/完成/工番中/工番末/前月迄出来高）を Task 1・3・4 で統一 ✅
