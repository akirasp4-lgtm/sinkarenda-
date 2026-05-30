# -*- coding: utf-8 -*-
"""きんでん指定請求書ジェネレーター。

xlsx(zip)内のシートXMLの対象セル値だけを直接書き換える。画像・印影・スタイル・
結合・数式は一切触らない（openpyxlはEMF画像が落ちるため使わない）。
"""
import zipfile
import re
import json
import calendar
import os
from datetime import date

S1 = 'xl/worksheets/sheet1.xml'   # きんでん情通C指定見積表紙
S2 = 'xl/worksheets/sheet2.xml'   # 請求書(インボイス）
WB = 'xl/workbook.xml'

# 着工/完成の 年(4桁)・月(2桁)・日(2桁) を1桁ずつ入れるセル群
START_CELLS = {'y': ['G32', 'H32', 'I32', 'J32'], 'm': ['L32', 'M32'], 'd': ['O32', 'P32']}
DONE_CELLS = {'y': ['W32', 'X32', 'Y32', 'Z32'], 'm': ['AB32', 'AC32'], 'd': ['AE32', 'AF32']}


def _xesc(s):
    return str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def _style_of(xml, ref):
    m = re.search(r'<c r="' + re.escape(ref) + r'"([^>]*?)(?:/>|>)', xml)
    if not m:
        raise ValueError('cell not found: ' + ref)
    sm = re.search(r's="(\d+)"', m.group(1))
    return sm.group(1) if sm else None


def set_cell(xml, ref, kind, value):
    """対象セルの <c> 要素を丸ごと置換する。style(s=) は維持。
    kind: 'n'=数値 / 's'=文字列(inlineString化)。
    """
    s = _style_of(xml, ref)
    sattr = (' s="%s"' % s) if s else ''
    if kind == 'n':
        new = '<c r="%s"%s><v>%s</v></c>' % (ref, sattr, value)
    else:
        new = '<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (
            ref, sattr, _xesc(value))
    pat = re.compile(r'<c r="' + re.escape(ref) + r'"[^>]*?(?:/>|>.*?</c>)', re.S)
    out, n = pat.subn(new, xml, count=1)
    if n != 1:
        raise ValueError('replace failed for %s (n=%d)' % (ref, n))
    return out


def enable_full_calc(wbxml):
    """開いた瞬間にExcelが全数式を再計算するよう fullCalcOnLoad を立てる。"""
    if 'calcPr' in wbxml:
        if 'fullCalcOnLoad' in wbxml:
            return re.sub(r'fullCalcOnLoad="[^"]*"', 'fullCalcOnLoad="1"', wbxml, count=1)
        return re.sub(r'<calcPr', '<calcPr fullCalcOnLoad="1"', wbxml, count=1)
    return re.sub(r'(<workbookPr[^>]*/>)', r'\1<calcPr fullCalcOnLoad="1"/>', wbxml, count=1)


def _digits(num, width):
    s = str(int(num)).zfill(width)
    return [int(ch) for ch in s[-width:]]


def _set_date(xml, cells, ymd):
    for grp, width, num in (('y', 4, ymd.year), ('m', 2, ymd.month), ('d', 2, ymd.day)):
        for cell, dig in zip(cells[grp], _digits(num, width)):
            xml = set_cell(xml, cell, 'n', dig)
    return xml


def _d(s):
    y, m, dd = (int(x) for x in s.split('-'))
    return date(y, m, dd)


def fill_workbook(template_path, out_path, kyoku, toushin_made, zen_made, seikyu_ym):
    """1局分の請求書を生成する。

    kyoku        : 局マスタの1局dict
    toushin_made : 当月迄出来高(A)（きんでん指示）
    zen_made     : 前月迄出来高(B)（マスタの累計）
    seikyu_ym    : 請求年月 'YYYY-MM'
    """
    sy, sm = (int(x) for x in seikyu_ym.split('-'))
    last_day = calendar.monthrange(sy, sm)[1]
    mitsumori = _d(kyoku['見積日'])
    zin = zipfile.ZipFile(template_path, 'r')
    s1 = zin.read(S1).decode('utf-8')
    s2 = zin.read(S2).decode('utf-8')
    wb = zin.read(WB).decode('utf-8')
    # --- 指定見積表紙 ---
    s1 = set_cell(s1, 'AB4', 's', kyoku['見積No'])
    s1 = set_cell(s1, 'AA5', 'n', sy)            # 請求年（請求書が参照）
    s1 = set_cell(s1, 'AE5', 'n', sm)            # 請求月
    s1 = set_cell(s1, 'AH5', 'n', mitsumori.day)
    s1 = set_cell(s1, 'G8', 'n', kyoku['注文金額'])
    s1 = set_cell(s1, 'H14', 'n', kyoku['労務費'])
    s1 = set_cell(s1, 'G28', 's', kyoku['工事名'])
    s1 = set_cell(s1, 'G30', 's', kyoku['施工場所'])
    s1 = _set_date(s1, START_CELLS, _d(kyoku['着工']))
    s1 = _set_date(s1, DONE_CELLS, _d(kyoku['完成']))
    s1 = set_cell(s1, 'M54', 's', kyoku['工番中'])
    s1 = set_cell(s1, 'P54', 's', kyoku['工番末'])
    # --- 工事代金請求明細書（インボイス）---
    s2 = set_cell(s2, 'AO17', 'n', toushin_made)   # 当月迄出来高(A)
    s2 = set_cell(s2, 'AO18', 'n', zen_made)       # 前月迄出来高(B)
    s2 = set_cell(s2, 'CO1', 'n', last_day)        # 請求日（末日）
    wb = enable_full_calc(wb)
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == S1:
                data = s1.encode('utf-8')
            elif it.filename == S2:
                data = s2.encode('utf-8')
            elif it.filename == WB:
                data = wb.encode('utf-8')
            zout.writestr(it, data)
    zin.close()


def run_month(master_path, template_path, out_dir, seikyu_ym, instructions):
    """instructions = {局名: 当月迄出来高}. 各局を生成し、前月迄累計を更新して保存。"""
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


if __name__ == '__main__':
    import argparse
    base_default = os.path.join(os.path.dirname(__file__), '..', '..', '_local', 'きんでん')
    ap = argparse.ArgumentParser(description='きんでん指定請求書ジェネレーター')
    ap.add_argument('--ym', required=True, help='請求年月 YYYY-MM')
    ap.add_argument('--kyoku', required=True, help='局名')
    ap.add_argument('--made', required=True, type=int, help='当月迄出来高（きんでん指示）')
    ap.add_argument('--base', default=base_default, help='_local/きんでん フォルダ')
    a = ap.parse_args()
    res = run_month(os.path.join(a.base, '局マスタ.json'),
                    os.path.join(a.base, 'template.xlsx'),
                    os.path.join(a.base, 'out'),
                    a.ym, {a.kyoku: a.made})
    for r in res:
        print('生成: %s（当月出来高 %s円）' % (r['file'], format(r['当月出来高'], ',')))
