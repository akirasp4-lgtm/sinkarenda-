# -*- coding: utf-8 -*-
"""きんでん指定請求書ジェネレーター（各局の実ファイルを土台に、毎月の出来高だけ差し込む）。

各局は固定情報（法定保険料率・工番・注文金額・労務費・工事名…が局ごとに異なる）を
自分のファイルに正しく持っているので、それを“土台”にして、毎月変わるセル
（当月迄出来高・前月迄出来高・請求日・請求年月）だけを直接書き換える。
画像・印影・スタイル・結合・数式・固定情報は一切触らない（openpyxl不使用＝画像保全）。
"""
import zipfile
import re
import json
import calendar
import os

S1 = 'xl/worksheets/sheet1.xml'   # きんでん情通C指定見積表紙
S2 = 'xl/worksheets/sheet2.xml'   # 請求書(インボイス）
WB = 'xl/workbook.xml'


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


def fill_monthly(base_path, out_path, toushin_made, zen_made, seikyu_ym):
    """1局分の当月請求書を生成する。base_path=その局の実ファイル(土台)。

    毎月変わるセルだけ書き換える：
      請求書 AO17=当月迄出来高(A) / AO18=前月迄出来高(B) / CO1=請求日(末日)
      表紙   AA5=請求年 / AE5=請求月（請求書がここを参照して年月表示）
    固定情報（料率H15・工番J54/M54/P54・注文金額G8・労務費H14・工事名G28…）は触らない。
    """
    sy, sm = (int(x) for x in seikyu_ym.split('-'))
    last_day = calendar.monthrange(sy, sm)[1]
    zin = zipfile.ZipFile(base_path, 'r')
    s1 = zin.read(S1).decode('utf-8')
    s2 = zin.read(S2).decode('utf-8')
    wb = zin.read(WB).decode('utf-8')
    s1 = set_cell(s1, 'AA5', 'n', sy)
    s1 = set_cell(s1, 'AE5', 'n', sm)
    s2 = set_cell(s2, 'AO17', 'n', toushin_made)
    s2 = set_cell(s2, 'AO18', 'n', zen_made)
    s2 = set_cell(s2, 'CO1', 'n', last_day)
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


def run_month(master_path, base_dir, out_dir, seikyu_ym, instructions):
    """instructions = {局名: 当月迄出来高}. 各局を生成し、前月迄累計を更新して保存。"""
    with open(master_path, encoding='utf-8') as fp:
        master = json.load(fp)
    results = []
    for name, toushin_made in instructions.items():
        ky = master['局'][name]
        base = os.path.join(base_dir, ky['file'])
        zen = ky.get('前月迄出来高', 0)
        out = os.path.join(out_dir, 'きんでん_%s_%s.xlsx' % (name, seikyu_ym))
        fill_monthly(base, out, toushin_made, zen, seikyu_ym)
        ky['前月迄出来高'] = toushin_made   # 次回の(B)
        order = ky.get('注文金額')
        results.append({'局': name, '当月出来高': toushin_made - zen,
                        '残額': (order - toushin_made) if order is not None else None,
                        'file': out})
    with open(master_path, 'w', encoding='utf-8') as fp:
        json.dump(master, fp, ensure_ascii=False, indent=2)
    return results


if __name__ == '__main__':
    import argparse
    base_default = os.path.join(os.path.dirname(__file__), '..', '..', '_local', 'きんでん')
    ap = argparse.ArgumentParser(description='きんでん指定請求書ジェネレーター（各局ファイルを土台に毎月の出来高を差込）')
    ap.add_argument('--ym', required=True, help='請求年月 YYYY-MM')
    ap.add_argument('--kyoku', required=True, help='局名（局マスタのキー）')
    ap.add_argument('--made', required=True, type=int, help='当月迄出来高（きんでん指示）')
    ap.add_argument('--base', default=base_default, help='_local/きんでん フォルダ')
    a = ap.parse_args()
    res = run_month(os.path.join(a.base, '局マスタ.json'), a.base,
                    os.path.join(a.base, 'out'), a.ym, {a.kyoku: a.made})
    for r in res:
        zan = ('／残額 %s円' % format(r['残額'], ',')) if r['残額'] is not None else ''
        print('生成: %s（当月出来高 %s円%s）' % (r['file'], format(r['当月出来高'], ','), zan))
