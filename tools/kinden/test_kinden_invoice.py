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

# ---- 1局生成（ひな型がある場合）----
TEMPLATE = os.path.join(os.path.dirname(__file__), '..', '..', '_local', 'きんでん', 'template.xlsx')
if os.path.exists(TEMPLATE):
    import openpyxl
    import zipfile as zf
    out = os.path.join(tempfile.gettempdir(), 'kinden_test_out.xlsx')
    ky = {'見積No': 'K00999', '見積日': '2026-05-01', '注文金額': 3000000, '労務費': 1500000,
          '工事名': 'テスト工事', '施工場所': '大阪市西区', '着工': '2026-05-02', '完成': '2026-05-20',
          '工番中': '7', '工番末': '2222'}
    K.fill_workbook(TEMPLATE, out, ky, toushin_made=1800000, zen_made=600000, seikyu_ym='2026-06')
    wb = openpyxl.load_workbook(out, data_only=False)
    t = wb['きんでん情通C指定見積表紙']
    r = wb['請求書(インボイス）']
    check('注文金額G8', t['G8'].value == 3000000)
    check('労務費H14', t['H14'].value == 1500000)
    check('法定福利費の数式が残る', str(t['H16'].value) == '=H14*H15')
    check('工事名G28', t['G28'].value == 'テスト工事')
    check('工番末P54', str(t['P54'].value) == '2222')
    check('当月迄AO17', r['AO17'].value == 1800000)
    check('前月迄AO18', r['AO18'].value == 600000)
    check('当月出来高の数式が残る', str(r['AO19'].value) == '=AO17-AO18')
    check('請求日CO1=月末(6月=30)', r['CO1'].value == 30)
    media = [n for n in zf.ZipFile(out).namelist() if '/media/' in n]
    check('画像4点が温存', len(media) == 4)
else:
    print('SKIP: ひな型が無いので生成テストはスキップ（_local/きんでん/template.xlsx を用意）')

# ---- 累計繰越（ひな型＋マスタがある場合）----
MASTER = os.path.join(os.path.dirname(__file__), '..', '..', '_local', 'きんでん', '局マスタ.json')
if os.path.exists(TEMPLATE) and os.path.exists(MASTER):
    import json as _json
    import shutil
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

print('\n==== %d passed, %d failed ====' % (p, f))
sys.exit(1 if f else 0)
