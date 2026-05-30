# -*- coding: utf-8 -*-
import sys
import os
import tempfile
sys.path.insert(0, os.path.dirname(__file__))
import kinden_invoice as K

p = f = 0


def check(name, cond):
    global p, f
    print(('PASS' if cond else 'FAIL') + ' : ' + name)
    if cond:
        p += 1
    else:
        f += 1


# ---- コア：セル値置換 ----
xml = '<row r="8"><c r="G8" s="264"><v>2530000</v></c></row>'
check('数値セル置換', '<c r="G8" s="264"><v>9999</v></c>' in K.set_cell(xml, 'G8', 'n', 9999))

xml2 = '<row r="18"><c r="AO18" s="102"/></row>'
check('空セルに数値', '<c r="AO18" s="102"><v>500000</v></c>' in K.set_cell(xml2, 'AO18', 'n', 500000))

xml3 = '<row r="28"><c r="G28" s="186" t="s"><v>155</v></c></row>'
o3 = K.set_cell(xml3, 'G28', 's', 'A&B<工事>')
check('文字列inlineString化＋エスケープ＋style維持',
      't="inlineStr"' in o3 and 'A&amp;B&lt;工事&gt;' in o3 and 's="186"' in o3)

check('calcPrにfullCalc付与', 'fullCalcOnLoad="1"' in K.enable_full_calc('<calcPr calcId="1"/>'))

# ---- 月次生成：土台(道頓堀template)に出来高だけ差込。固定情報は不変 ----
BASE_DIR = os.path.join(os.path.dirname(__file__), '..', '..', '_local', 'きんでん')
TEMPLATE = os.path.join(BASE_DIR, 'template.xlsx')
if os.path.exists(TEMPLATE):
    import openpyxl
    import zipfile as zf
    out = os.path.join(tempfile.gettempdir(), 'kinden_test_out.xlsx')
    K.fill_monthly(TEMPLATE, out, toushin_made=2000000, zen_made=800000, seikyu_ym='2026-06')
    wb = openpyxl.load_workbook(out, data_only=False)
    t = wb['きんでん情通C指定見積表紙']
    r = wb['請求書(インボイス）']
    # 毎月セルは更新される
    check('当月迄AO17更新', r['AO17'].value == 2000000)
    check('前月迄AO18更新', r['AO18'].value == 800000)
    check('請求日CO1=月末(6月=30)', r['CO1'].value == 30)
    check('請求年AA5=2026', t['AA5'].value == 2026)
    check('請求月AE5=6', t['AE5'].value == 6)
    # 固定情報は土台のまま不変（道頓堀template: G8=2530000, H15=0.1636, P54=1317）
    check('注文金額G8は不変(2530000)', t['G8'].value == 2530000)
    check('料率H15は不変(0.1636)', t['H15'].value == 0.1636)
    check('工番P54は不変(1317)', str(t['P54'].value) == '1317')
    # 数式・画像は温存
    check('当月出来高の数式が残る', str(r['AO19'].value) == '=AO17-AO18')
    check('法定福利費の数式が残る', str(t['H16'].value) == '=H14*H15')
    media = [n for n in zf.ZipFile(out).namelist() if '/media/' in n]
    check('画像4点が温存', len(media) == 4)
else:
    print('SKIP: ひな型が無いので月次生成テストはスキップ')

# ---- 累計繰越（実マスタの一時コピーで）----
MASTER = os.path.join(BASE_DIR, '局マスタ.json')
if os.path.exists(TEMPLATE) and os.path.exists(MASTER):
    import json as _json
    import shutil
    tmp_master = os.path.join(tempfile.gettempdir(), 'kinden_master_test.json')
    shutil.copy(MASTER, tmp_master)
    m = _json.load(open(MASTER, encoding='utf-8'))
    # ファイルが実在する局を1つ選ぶ
    name = None
    for k, v in m['局'].items():
        if os.path.exists(os.path.join(BASE_DIR, v['file'])):
            name = k
            break
    if name:
        outdir = tempfile.gettempdir()
        res = K.run_month(tmp_master, BASE_DIR, outdir, '2026-06', {name: 1234567})
        after = _json.load(open(tmp_master, encoding='utf-8'))['局'][name]['前月迄出来高']
        check('実行後に前月迄が当月迄(1234567)へ更新', after == 1234567)
        check('出力ファイルが作られる', os.path.exists(res[0]['file']))
    else:
        print('SKIP: 実在ファイルの局が無い')
else:
    print('SKIP: 累計テスト（ひな型/マスタ未設置）')

print('\n==== %d passed, %d failed ====' % (p, f))
sys.exit(1 if f else 0)
