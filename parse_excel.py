# -*- coding: utf-8 -*-
"""
budai_planner Excel数据提取脚本
用openpyxl读取Excel，完整保留单元格颜色和日期信息
输出JSON供budai_planner.html加载

用法:
  python parse_excel.py "C:/CULINES/Claw Report/REX & FEEDER SCHEDULE 5.25.xlsx" 25_May
  python parse_excel.py <excel路径> <sheet名>
"""

import sys
import json
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime, date

def rgb_str(fill):
    """提取填充色RGB字符串"""
    try:
        if not fill or fill.fill_type is None:
            return '00000000'
        fg = fill.fgColor
        if fg and fg.rgb:
            s = str(fg.rgb).upper()
            # 去掉8位ARGB的alpha前缀: FFDAF2D0 -> DAF2D0
            if len(s) == 8 and s.startswith('FF'):
                s = s[2:]
            elif len(s) == 8 and s.startswith('00'):
                s = s[2:]
            return s
        return '00000000'
    except Exception:
        return '00000000'


def dt_to_str(val):
    """将日期值转为 YYYY-MM-DD 字符串"""
    if val is None:
        return None
    if isinstance(val, (datetime, date)):
        return val.strftime('%Y-%m-%d')
    # Excel serial number
    try:
        n = float(val)
        if n > 25000 and n < 60000:
            base = datetime(1899, 12, 30)
            dt = base + __import__('datetime').timedelta(days=int(n))
            return dt.strftime('%Y-%m-%d')
    except (ValueError, TypeError):
        pass
    s = str(val).strip()
    if len(s) == 10 and '-' in s:
        return s
    return None


def parse_excel(filepath, sheet_name):
    wb = load_workbook(filepath, data_only=True)
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.worksheets[0]

    # 找表头行 (搜索 CUL VESSELS 列)
    hdr_row = -1
    headers = []
    for r in range(1, min(ws.max_row + 1, 200)):
        for c in range(1, min(ws.max_column + 1, 50)):
            cell_val = ws.cell(r, c).value
            if cell_val and 'CUL VESSELS' in str(cell_val).upper():
                hdr_row = r
                break
        if hdr_row > 0:
            break

    if hdr_row < 0:
        print("ERROR: 找不到表头(CUL VESSELS列)")
        sys.exit(1)

    # 读表头
    col_map = {}
    for c in range(1, ws.max_column + 1):
        h = ws.cell(hdr_row, c).value
        if h:
            col_map[str(h).strip().upper()] = c

    print(f"表头行: {hdr_row}, 列映射: {list(col_map.keys())}")

    # 关键列索引（兼容不同列名）
    ves_col = col_map.get('CUL VESSELS', col_map.get('VESSEL', 0))
    pkg_col = col_map.get('PKG', 0)
    pkge_col = col_map.get('PKGE', 0)
    teu_col = col_map.get('EFF. TEUS', col_map.get('EFF.TEUS', col_map.get('TEU', 0)))
    svc_col = col_map.get('SERVICES', col_map.get('SVC', 0))

    # 额外有用的列
    pol_col = col_map.get('POL', 0)       # 起运港
    pod_col = col_map.get('POD', 0)       # 目的港
    disch_vol_col = col_map.get('DISCHARGE VOL', 0)   # 货量

    result = {'whites': [], 'feeders': [], 'pinks': [], 'meta': {
        'source': filepath,
        'sheet': sheet_name,
        'headerRow': hdr_row,
        'columns': {k: v for k, v in col_map.items()
                    if k in ['CUL VESSELS','PKG','PKGE','EFF. TEUS',
                             'SERVICES','POL','POD','DISCHARGE VOL']}
    }}

    for r in range(hdr_row + 1, ws.max_row + 1):
        ves = ws.cell(r, ves_col).value
        if not ves or not str(ves).strip():
            continue

        name = str(ves).strip()

        # 取船名单元格的颜色（作为整行的分类依据）
        color = rgb_str(ws.cell(r, ves_col).fill)

        # 日期
        pkg_date = dt_to_str(ws.cell(r, pkg_col).value) if pkg_col else None
        pkge_date = dt_to_str(ws.cell(r, pkge_col).value) if pkge_col else None

        # TEU / 货量
        teu_val = ws.cell(r, teu_col).value if teu_col else 0
        try:
            teu_num = int(float(teu_val)) if teu_val else 0
        except (ValueError, TypeError):
            teu_num = 0

        disch_val = ws.cell(r, disch_vol_col).value if disch_vol_col else None
        try:
            disch_num = int(float(disch_val)) if disch_val else 0
        except (ValueError, TypeError):
            disch_num = 0

        svc = str(ws.cell(r, svc_col).value).strip() if svc_col else ''
        pol = str(ws.cell(r, pol_col).value).strip() if pol_col else ''
        pod = str(ws.cell(r, pod_col).value).strip() if pod_col else ''

        entry = {
            'row': r,
            'name': name,
            'pkg': pkg_date,
            'pkge': pkge_date,
            'teu': teu_num,
            'dischargeVol': disch_num,
            'svc': svc,
            'color': color,
            'pol': pol or '—',
            'pod': pod or '—'
        }

        # 颜色分类
        if color == '00000000' or color == 'FFFFFFFF' or not color:
            entry['type'] = 'W'
            entry['typeLabel'] = '干线'
            entry['fromPort'] = pol or '中国各港'
            entry['toPort'] = ''
            result['whites'].append(entry)
        elif 'DAF2D0' in color:
            entry['type'] = 'P'
            entry['typeLabel'] = '直航'
            entry['fromPort'] = pol or '中国各港'
            entry['toPort'] = pod or '红海'
            result['pinks'].append(entry)
        elif 'FBE2D5' in color:
            entry['type'] = 'F'
            entry['typeLabel'] = '布袋'
            entry['fromPort'] = '巴生PKG'
            entry['toPort'] = pod or '红海'
            result['feeders'].append(entry)
        else:
            entry['type'] = '?'
            entry['typeLabel'] = '未知(' + color + ')'
            entry['fromPort'] = pol or '—'
            entry['toPort'] = pod or '—'
            # 未知颜色的也先放到whites里，前端会让用户手动分类
            result['whites'].append(entry)

    # 统计
    result['summary'] = {
        'whites': len(result['whites']),
        'feeders': len(result['feeders']),
        'pinks': len(result['pinks']),
        'total': len(result['whites']) + len(result['feeders']) + len(result['pinks'])
    }

    return result


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python parse_excel.py <xlsx文件> [sheet名]")
        print("例: python parse_excel.py schedule.xlsx 25_May")
        sys.exit(1)

    xlsx_path = sys.argv[1]
    sheet = sys.argv[2] if len(sys.argv) > 2 else None

    data = parse_excel(xlsx_path, sheet)

    # 输出JSON到同目录
    out_path = xlsx_path.rsplit('.', 1)[0] + '_parsed.json'
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"\n=== 解析完成 ===")
    print(f"干线条: {data['summary']['whites']} 艘")
    print(f"布袋船: {data['summary']['feeders']} 艘")
    print(f"直航船: {data['summary']['pinks']} 艘")
    print(f"总计:   {data['summary']['total']} 艘")
    print(f"\nJSON已保存: {out_path}")
