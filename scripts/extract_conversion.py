#!/usr/bin/env python3
"""
从「日维度_搜索方式转化情况」Excel 中提取指定搜索方式的关键指标，
生成带格式的汇总 Excel。

用法:
    python3 extract_conversion.py <excel文件路径> [起始日期] [结束日期]

日期格式: YYYYMMDD (如 20260907)
不指定日期时自动提取全部日期。
"""

import sys
import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# 固定提取的搜索方式及顺序
TARGET_METHODS = [
    "手动文字-AI搜",
    "底纹词-AI搜",
    "语音输入-AI搜",
    "试试这样问-AI搜",
    "你最近问过-AI搜",
    "推荐词-前置页-AI搜",
]

# 源 Excel 列映射: (输出列名, 源列索引1-based)
FIELD_MAP = [
    ("搜索pv", 4),
    ("搜索uv", 5),
    ("总提问条数", 7),
    ("直接订单数", 56),
]

OUTPUT_HEADERS = ["日期", "搜索方式", "搜索pv", "搜索uv", "总提问条数", "直接订单数"]


def extract_data(src_path, date_start=None, date_end=None):
    """从源 Excel 读取数据，返回 [(date, method, pv, uv, tqs, ordn), ...]"""
    wb = openpyxl.load_workbook(src_path, data_only=True)
    ws = wb["日维度_搜索方式转化情况"]

    target_set = set(TARGET_METHODS)
    rows = []
    last_date = None

    for r in range(2, ws.max_row + 1):
        d = ws.cell(r, 1).value
        if d is not None:
            last_date = str(d)
        method = ws.cell(r, 3).value
        if method not in target_set:
            continue

        # 日期过滤
        if date_start and last_date < date_start:
            continue
        if date_end and last_date > date_end:
            continue

        vals = []
        for _, col_idx in FIELD_MAP:
            v = ws.cell(r, col_idx).value
            vals.append(v if v is not None else 0)
        rows.append((last_date, method, *vals))

    return rows


def write_output(rows, out_path):
    """写入带格式的 Excel"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "搜索方式转化情况"

    # 样式
    header_font = Font(bold=True, size=11)
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # 表头
    for col, h in enumerate(OUTPUT_HEADERS, 1):
        cell = ws.cell(1, col, h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    # 数据
    for i, row_data in enumerate(rows, 2):
        for col, v in enumerate(row_data, 1):
            cell = ws.cell(i, col, v)
            cell.alignment = center
            cell.border = border

    # 列宽
    widths = [12, 18, 10, 10, 12, 12]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(out_path)


def main():
    if len(sys.argv) < 2:
        print("用法: python3 extract_conversion.py <excel文件路径> [起始日期] [结束日期]")
        print("日期格式: YYYYMMDD")
        sys.exit(1)

    src_path = sys.argv[1]
    date_start = sys.argv[2] if len(sys.argv) > 2 else None
    date_end = sys.argv[3] if len(sys.argv) > 3 else None

    if not os.path.exists(src_path):
        print(f"错误: 文件不存在 - {src_path}")
        sys.exit(1)

    rows = extract_data(src_path, date_start, date_end)
    if not rows:
        print("错误: 未找到匹配数据，请检查日期范围和搜索方式")
        sys.exit(1)

    # 排序: 日期升序 + 搜索方式固定顺序
    method_order = {m: i for i, m in enumerate(TARGET_METHODS)}
    rows.sort(key=lambda r: (r[0], method_order.get(r[1], 99)))

    # 生成输出路径
    dates = [r[0] for r in rows]
    d_min, d_max = min(dates), max(dates)
    if date_start or date_end:
        name_range = f"{date_start or d_min}-{date_end or d_max}"
    else:
        name_range = f"{d_min}-{d_max}"

    # 输出到 outputs 目录
    cwd = os.getcwd()
    out_dir = os.path.join(cwd, "outputs")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f"搜索方式转化情况_{name_range}.xlsx")

    write_output(rows, out_path)

    print(f"已生成: {out_path}")
    print(f"共 {len(rows)} 行数据 ({len(set(dates))} 天 × {len(TARGET_METHODS)} 种搜索方式)")
    print(f"日期范围: {d_min} ~ {d_max}")


if __name__ == "__main__":
    main()
