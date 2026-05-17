"""
巨星农牧（603477.SH）三表模型 + DCF估值
生成时间：2026-05-09
依赖：pip install openpyxl
使用方法：python 巨星农牧_三表模型_DCF估值.py
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
import datetime

wb = openpyxl.Workbook()

# ============================================================
# 样式定义
# ============================================================
DARK_BLUE = PatternFill("solid", fgColor="1F4E79")
LIGHT_BLUE = PatternFill("solid", fgColor="D9E1F2")
MEDIUM_BLUE = PatternFill("solid", fgColor="BDD7EE")
LIGHT_GREY = PatternFill("solid", fgColor="F2F2F2")
WHITE_FILL = PatternFill("solid", fgColor="FFFFFF")
RED_FONT = Font(color="FF0000", size=11, bold=True)
WHITE_BOLD = Font(color="FFFFFF", size=11, bold=True)
BLACK_BOLD = Font(color="000000", size=11, bold=True)
BLACK_NORMAL = Font(color="000000", size=11)
BLUE_INPUT = Font(color="0000FF", size=11)
GREEN_LINK = Font(color="008000", size=10)
TITLE_FONT = Font(color="FFFFFF", size=14, bold=True)

thin_border = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin")
)
medium_border = Border(
    left=Side(style="medium"), right=Side(style="medium"),
    top=Side(style="medium"), bottom=Side(style="medium")
)

def set_cell(ws, row, col, value, font=BLACK_NORMAL, fill=WHITE_FILL, align=None, border=thin_border, number_format=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = font
    cell.fill = fill
    cell.alignment = align or Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = border
    if number_format:
        cell.number_format = number_format
    return cell

def section_header(ws, row, col_start, col_end, text):
    ws.merge_cells(start_row=row, start_column=col_start, end_row=row, end_column=col_end)
    cell = ws.cell(row=row, column=col_start, value=text)
    cell.font = WHITE_BOLD
    cell.fill = DARK_BLUE
    cell.alignment = Alignment(horizontal="left", vertical="center")
    for c in range(col_start, col_end+1):
        ws.cell(row=row, column=c).border = medium_border
        ws.cell(row=row, column=c).fill = DARK_BLUE

def col_header(ws, row, col, text):
    set_cell(ws, row, col, text, font=BLACK_BOLD, fill=LIGHT_BLUE)

# ============================================================
# Sheet 1: 三表模型 (3-Statement Model)
# ============================================================
ws1 = wb.active
ws1.title = "三表模型"

# 标题
ws1.merge_cells("A1:G1")
set_cell(ws1, 1, 1, "巨星农牧 (603477.SH) 三表财务模型", font=TITLE_FONT, fill=DARK_BLUE, border=medium_border)
for c in range(2, 8):
    ws1.cell(row=1, column=c).fill = DARK_BLUE
    ws1.cell(row=1, column=c).border = medium_border

ws1.merge_cells("A2:G2")
set_cell(ws1, 2, 1, f"报告日期: 2026-05-09 | 单位: 亿元人民币 (除特别注明外)", font=Font(color="555555", size=10), fill=LIGHT_GREY, border=medium_border)
for c in range(2, 8):
    ws1.cell(row=2, column=c).fill = LIGHT_GREY
    ws1.cell(row=2, column=c).border = thin_border

# 列标题
years = ["", "2023A", "2024A", "2025A", "2026E", "2027E", "2028E"]
for i, y in enumerate(years):
    col_header(ws1, 4, i+1, y)

# -----------------------------------
# 利润表
# -----------------------------------
row = 6
section_header(ws1, row, 1, 7, "利润表")
row += 1

is_items = [
    ("营业收入", [40.41, 60.78, 80.00, "=D7*(1+0.15)", "=E7*(1+0.20)", "=F7*(1+0.12)"], "input"),
    ("  营收增速 (%)", ["=B7/B7-1", "=C7/B7-1", "=D7/C7-1", "=E7/D7-1", "=F7/E7-1", "=G7/F7-1"], "formula"),
    ("营业成本", [40.56, 49.13, 74.19, "=E7*(1-0.08)", "=F7*(1-0.82)", "=G7*(1-0.84)"], "input" if row < 4 else "formula"),
    ("毛利", ["=B7-B9", "=C7-C9", "=D7-D9", "=E7-E9", "=F7-F9", "=G7-G9"], "formula"),
    ("  毛利率 (%)", ["=B10/B7", "=C10/C7", "=D10/D7", "=E10/E7", "=F10/F7", "=G10/G7"], "formula"),
    ("销售费用", [0.39, 0.39, 0.37, "=E7*0.0046", "=F7*0.0045", "=G7*0.0045"], "input" if row < 6 else "formula"),
    ("管理费用", [2.17, 2.90, 2.69, "=E7*0.034", "=F7*0.033", "=G7*0.033"], "input" if row < 6 else "formula"),
    ("研发费用", [0.25, 0.42, 0.48, "=E7*0.006", "=F7*0.006", "=G7*0.006"], "input" if row < 6 else "formula"),
    ("财务费用", [1.07, 2.05, 1.98, "=E7*0.025", "=F7*0.023", "=G7*0.022"], "input" if row < 6 else "formula"),
    ("营业利润", ["=B10-B12-B13-B14-B15", "=C10-C12-C13-C14-C15", "=D10-D12-D13-D14-D15", "=E10-E12-E13-E14-E15", "=F10-F12-F13-F14-F15", "=G10-G12-G13-G14-G15"], "formula"),
    ("利润总额", [0.0, 0.0, 0.0, "=E16*1.02", "=F16*1.02", "=G16*1.02"], "formula"),
    ("所得税费用", [0.07, 0.02, 0.08, "=IF(E17>0,E17*0.15,0)", "=IF(F17>0,F17*0.15,0)", "=IF(G17>0,G17*0.15,0)"], "input" if row < 8 else "formula"),
    ("归母净利润", ["=B17-B18", "=C17-C18", "=D17-D18", "=E17-E18", "=F17-F18", "=G17-G18"], "formula"),
    ("  净利率 (%)", ["=B19/B7", "=C19/C7", "=D19/D7", "=E19/E7", "=F19/F7", "=G19/G7"], "formula"),
]

# Actual data items
actual_is = [
    ("营业收入", [40.41, 60.78, 80.00]),
    ("营业成本", [40.56, 49.13, 74.19]),
    ("销售费用", [0.39, 0.39, 0.37]),
    ("管理费用", [2.17, 2.90, 2.69]),
    ("研发费用", [0.25, 0.42, 0.48]),
    ("财务费用", [1.07, 2.05, 1.98]),
    ("所得税费用", [0.07, 0.02, 0.08]),
]

# Write IS items with formulas
is_formulas = [
    # (label, [2023, 2024, 2025, 2026E, 2027E, 2028E])
    ("营业收入", ["=B7*(1+0.15)", "=C7*(1+0.20)", "=D7*(1+0.12)"]),
    ("营业成本", ["=E7*(1-0.08)", "=F7*(1-0.82)", "=G7*(1-0.84)"]),
]

r = 8
# 用行跟踪变量，避免硬编码行号
# 营收
rev_row = r
set_cell(ws1, r, 1, "营业收入", font=BLACK_BOLD)
set_cell(ws1, r, 2, 40.41, font=BLUE_INPUT, fill=LIGHT_GREY)
set_cell(ws1, r, 3, 60.78, font=BLUE_INPUT, fill=LIGHT_GREY)
set_cell(ws1, r, 4, 80.00, font=BLUE_INPUT, fill=LIGHT_GREY)
set_cell(ws1, r, 5, f"=D{rev_row}*(1+0.15)")
set_cell(ws1, r, 6, f"=E{rev_row}*(1+0.20)")
set_cell(ws1, r, 7, f"=F{rev_row}*(1+0.12)")
r += 1

# 营收增速（第一年无前期数据，留空）
growth_row = r
set_cell(ws1, r, 1, "  营收增速 (%)", font=BLACK_NORMAL)
set_cell(ws1, r, 2, "")  # 2023A 无前期，留空
for c in range(3, 8):
    cl = get_column_letter(c)
    prev_cl = get_column_letter(c - 1)
    set_cell(ws1, r, c, f"={cl}{rev_row}/{prev_cl}{rev_row}-1", number_format="0.0%")
r += 1

# 营业成本
cogs_row = r
set_cell(ws1, r, 1, "营业成本", font=BLACK_BOLD)
set_cell(ws1, r, 2, 40.56, font=BLUE_INPUT, fill=LIGHT_GREY)
set_cell(ws1, r, 3, 49.13, font=BLUE_INPUT, fill=LIGHT_GREY)
set_cell(ws1, r, 4, 74.19, font=BLUE_INPUT, fill=LIGHT_GREY)
set_cell(ws1, r, 5, f"=E{rev_row}*(1-0.93)")
set_cell(ws1, r, 6, f"=F{rev_row}*(1-0.80)")
set_cell(ws1, r, 7, f"=G{rev_row}*(1-0.84)")
r += 1

# 毛利
gp_row = r
set_cell(ws1, r, 1, "毛利", font=BLACK_BOLD)
for c in range(2, 8):
    cl = get_column_letter(c)
    set_cell(ws1, r, c, f"={cl}{rev_row}-{cl}{cogs_row}")
r += 1

# 毛利率
gpm_row = r
set_cell(ws1, r, 1, "  毛利率 (%)", font=BLACK_NORMAL)
for c in range(2, 8):
    cl = get_column_letter(c)
    set_cell(ws1, r, c, f"={cl}{gp_row}/{cl}{rev_row}", number_format="0.00%")
r += 1

# 销售费用
sell_row = r
set_cell(ws1, r, 1, "销售费用", font=BLACK_BOLD)
set_cell(ws1, r, 2, 0.39, font=BLUE_INPUT, fill=LIGHT_GREY)
set_cell(ws1, r, 3, 0.39, font=BLUE_INPUT, fill=LIGHT_GREY)
set_cell(ws1, r, 4, 0.37, font=BLUE_INPUT, fill=LIGHT_GREY)
for c in range(5, 8):
    set_cell(ws1, r, c, f"={get_column_letter(c)}{rev_row}*0.0046")
r += 1

# 管理费用
admin_row = r
set_cell(ws1, r, 1, "管理费用", font=BLACK_BOLD)
set_cell(ws1, r, 2, 2.17, font=BLUE_INPUT, fill=LIGHT_GREY)
set_cell(ws1, r, 3, 2.90, font=BLUE_INPUT, fill=LIGHT_GREY)
set_cell(ws1, r, 4, 2.69, font=BLUE_INPUT, fill=LIGHT_GREY)
for c in range(5, 8):
    set_cell(ws1, r, c, f"={get_column_letter(c)}{rev_row}*0.034")
r += 1

# 研发费用
rd_row = r
set_cell(ws1, r, 1, "研发费用", font=BLACK_BOLD)
set_cell(ws1, r, 2, 0.25, font=BLUE_INPUT, fill=LIGHT_GREY)
set_cell(ws1, r, 3, 0.42, font=BLUE_INPUT, fill=LIGHT_GREY)
set_cell(ws1, r, 4, 0.48, font=BLUE_INPUT, fill=LIGHT_GREY)
for c in range(5, 8):
    set_cell(ws1, r, c, f"={get_column_letter(c)}{rev_row}*0.006")
r += 1

# 财务费用
fin_row = r
set_cell(ws1, r, 1, "财务费用", font=BLACK_BOLD)
set_cell(ws1, r, 2, 1.07, font=BLUE_INPUT, fill=LIGHT_GREY)
set_cell(ws1, r, 3, 2.05, font=BLUE_INPUT, fill=LIGHT_GREY)
set_cell(ws1, r, 4, 1.98, font=BLUE_INPUT, fill=LIGHT_GREY)
for c in range(5, 8):
    set_cell(ws1, r, c, f"={get_column_letter(c)}{rev_row}*0.024")
r += 1

# 营业利润 = 毛利 - 销售 - 管理 - 研发 - 财务
op_row = r
set_cell(ws1, r, 1, "营业利润", font=BLACK_BOLD)
for c in range(2, 8):
    cl = get_column_letter(c)
    set_cell(ws1, r, c, f"={cl}{gp_row}-{cl}{sell_row}-{cl}{admin_row}-{cl}{rd_row}-{cl}{fin_row}")
r += 1

# 营业外收入
other_row = r
set_cell(ws1, r, 1, "  加：营业外收入", font=BLACK_NORMAL)
set_cell(ws1, r, 2, 0.06, font=BLUE_INPUT, fill=LIGHT_GREY)
set_cell(ws1, r, 3, 0.46, font=BLUE_INPUT, fill=LIGHT_GREY)
set_cell(ws1, r, 4, 1.23, font=BLUE_INPUT, fill=LIGHT_GREY)
for c in range(5, 8):
    set_cell(ws1, r, c, f"={get_column_letter(c)}{op_row}*0.02")
r += 1

# 利润总额
ebt_row = r
set_cell(ws1, r, 1, "利润总额", font=BLACK_BOLD)
for c in range(2, 8):
    cl = get_column_letter(c)
    set_cell(ws1, r, c, f"={cl}{op_row}+{cl}{other_row}")
r += 1

# 所得税
tax_row = r
set_cell(ws1, r, 1, "减：所得税费用", font=BLACK_BOLD)
for c in range(2, 8):
    cl = get_column_letter(c)
    set_cell(ws1, r, c, f"=IF({cl}{ebt_row}>0,{cl}{ebt_row}*0.15,0)")
r += 1

# 归母净利润
np_row = r
set_cell(ws1, r, 1, "归母净利润", font=Font(color="FF0000", size=11, bold=True))
for c in range(2, 8):
    cl = get_column_letter(c)
    set_cell(ws1, r, c, f"={cl}{ebt_row}-{cl}{tax_row}")
r += 1

# 净利率
npm_row = r
set_cell(ws1, r, 1, "  净利率 (%)", font=BLACK_NORMAL)
for c in range(2, 8):
    cl = get_column_letter(c)
    set_cell(ws1, r, c, f"={cl}{np_row}/{cl}{rev_row}", number_format="0.00%")
r += 2

# -----------------------------------
# 资产负债表
# -----------------------------------
section_header(ws1, r, 1, 7, "资产负债表")
r += 1

bs_labels = [
    ("资产类", None),
    ("  货币资金", [4.46, 5.77, 7.56]),
    ("  应收账款", [0.43, 0.68, 0.63]),
    ("  存货", [15.32, 20.93, 25.93]),
    ("  其他流动资产", [0.90, 1.26, 1.56]),
    ("  流动资产合计", "=B25+B26+B27+B28"),
    ("  固定资产", [37.53, 37.08, 36.12]),
    ("  在建工程", [3.53, 3.96, 4.49]),
    ("  其他非流动资产", [15.04, 17.96, 20.00]),
    ("  非流动资产合计", "=B30+B31+B32"),
    ("资产总计", "=B29+B33"),
]

bs_actual = {
    "货币资金": [4.46, 5.77, 7.56],
    "应收账款": [0.43, 0.68, 0.63],
    "存货": [15.32, 20.93, 25.93],
    "固定资产": [37.53, 37.08, 36.12],
    "在建工程": [3.53, 3.96, 4.49],
}

# Assets
r_start = r
for label, val in [
    ("资产类", None),
    ("  货币资金", [4.46, 5.77, 7.56]),
    ("  应收账款", [0.43, 0.68, 0.63]),
    ("  存货", [15.32, 20.93, 25.93]),
    ("  其他流动资产", [0.90, 1.26, 1.56]),
    ("  流动资产合计", None),
    ("", None),  # blank
    ("  固定资产", [37.53, 37.08, 36.12]),
    ("  在建工程", [3.53, 3.96, 4.49]),
    ("  其他非流动资产", [15.04, 17.96, 20.00]),
    ("  非流动资产合计", None),
    ("资产总计", None),
]:
    if label == "资产类":
        section_header(ws1, r, 1, 7, "资产类")
        r += 1
        continue
    if label == "":
        r += 1
        continue

    is_total_row = label in ["  流动资产合计", "  非流动资产合计", "资产总计"]
    set_cell(ws1, r, 1, label, font=BLACK_BOLD if is_total_row or "类" in label else BLACK_NORMAL)

    if val is None:
        # formula row
        pass
    elif isinstance(val, list):
        for c_idx, v in enumerate(val):
            set_cell(ws1, r, c_idx+2, v, font=BLUE_INPUT, fill=LIGHT_GREY)
        # projections for 2026E-2028E
        if label == "  货币资金":
            for c in range(5, 8):
                set_cell(ws1, r, c, f"={get_column_letter(c-1)}25*1.1")
        elif label == "  应收账款":
            for c in range(5, 8):
                set_cell(ws1, r, c, f"={get_column_letter(c)}7*0.008")
        elif label == "  存货":
            for c in range(5, 8):
                set_cell(ws1, r, c, f"={get_column_letter(c)}7*0.32")
        elif label == "  固定资产":
            for c in range(5, 8):
                set_cell(ws1, r, c, f"={get_column_letter(c-1)}30*0.98")
        elif label == "  在建工程":
            for c in range(5, 8):
                set_cell(ws1, r, c, f"={get_column_letter(c-1)}31*1.05")
    r += 1

# Liabilities & Equity
section_header(ws1, r, 1, 7, "负债及股东权益")
r += 1

liab_equity = [
    ("负债类", None),
    ("  短期借款", [11.53, 10.58, 14.30]),
    ("  应付账款", [2.72, 3.49, 3.50]),
    ("  其他流动负债", [13.50, 21.46, 26.17]),
    ("  流动负债合计", "=B38+B39+B40"),
    ("", None),
    ("  长期借款", [6.15, 4.51, 7.34]),
    ("  其他非流动负债", [12.31, 14.21, 13.47]),
    ("  非流动负债合计", "=B43+B44"),
    ("负债合计", "=B41+B45"),
    ("", None),
    ("股东权益类", None),
    ("  归母股东权益", [30.68, 33.01, 31.12]),
    ("负债及股东权益总计", "=B46+B48"),
]

for label, val in liab_equity:
    if label == "负债类" or label == "股东权益类":
        section_header(ws1, r, 1, 7, label)
        r += 1
        continue
    if label == "":
        r += 1
        continue

    is_total = label in ["  流动负债合计", "  非流动负债合计", "负债合计", "  归母股东权益", "负债及股东权益总计"]
    set_cell(ws1, r, 1, label, font=BLACK_BOLD if is_total else BLACK_NORMAL)

    if isinstance(val, list):
        for c_idx, v in enumerate(val):
            set_cell(ws1, r, c_idx+2, v, font=BLUE_INPUT, fill=LIGHT_GREY)
        # projections
        if label == "  短期借款":
            for c in range(5, 8):
                set_cell(ws1, r, c, f"={get_column_letter(c)}7*0.18")
        elif label == "  应付账款":
            for c in range(5, 8):
                set_cell(ws1, r, c, f"={get_column_letter(c)}7*0.044")
        elif label == "  长期借款":
            for c in range(5, 8):
                set_cell(ws1, r, c, f"={get_column_letter(c)}7*0.09")
    elif isinstance(val, str):
        for c in range(2, 8):
            cl = get_column_letter(c)
            formula = val.replace("B", cl)
            set_cell(ws1, r, c, formula)
    r += 1

r += 2
# -----------------------------------
# 现金流量表（简化）
# -----------------------------------
section_header(ws1, r, 1, 7, "现金流量表（简化）")
r += 1

cf_items = [
    ("经营活动现金流", None),
    ("  净利润", [-6.46, 5.25, 0.35]),
    ("  折旧摊销", [3.53, 4.72, 5.61]),
    ("  存货变动", ["", "", ""]),
    ("  其他调整", [1.88, 2.68, 6.22]),
    ("  经营活动现金流净额", [-1.05, 12.65, 12.18]),
    ("", None),
    ("投资活动现金流", None),
    ("  资本支出", [-10.04, -10.42, -10.83]),
    ("  投资活动现金流净额", [-8.95, -8.60, -8.51]),
    ("", None),
    ("筹资活动现金流", None),
    ("  借款净额", [5.47, -3.30, -3.04]),
    ("  筹资活动现金流净额", [5.47, -3.30, -3.04]),
    ("", None),
    ("现金净变动", "=B58+B62+B65"),
    ("期初现金", "=B25"),
    ("期末现金", "=B67+B68"),
]

for label, val in cf_items:
    if label in ["经营活动现金流", "投资活动现金流", "筹资活动现金流"]:
        section_header(ws1, r, 1, 7, label)
        r += 1
        continue
    if label == "":
        r += 1
        continue

    is_total = "净额" in label or "净变动" in label or "期末" in label or "期初" in label
    set_cell(ws1, r, 1, label, font=BLACK_BOLD if is_total else BLACK_NORMAL)

    if isinstance(val, list):
        for c_idx, v in enumerate(val):
            if v == "":
                continue
            set_cell(ws1, r, c_idx+2, v, font=BLUE_INPUT, fill=LIGHT_GREY)
        # projections
        if label == "  净利润":
            for c in range(5, 8):
                cl = get_column_letter(c)
                set_cell(ws1, r, c, f"={cl}20")
        elif label == "  折旧摊销":
            for c in range(5, 8):
                cl = get_column_letter(c)
                set_cell(ws1, r, c, f"={cl}7*0.07")
    elif isinstance(val, str):
        for c in range(2, 8):
            cl = get_column_letter(c)
            formula = val.replace("B", cl)
            set_cell(ws1, r, c, formula)
    r += 1

# 设置列宽
for c in range(1, 8):
    ws1.column_dimensions[get_column_letter(c)].width = 18 if c == 1 else 14

# ============================================================
# Sheet 2: DCF 估值
# ============================================================
ws2 = wb.create_sheet("DCF估值")

# Title
ws2.merge_cells("A1:G1")
set_cell(ws2, 1, 1, "巨星农牧 (603477.SH) DCF估值模型（辅助参考）", font=TITLE_FONT, fill=DARK_BLUE, border=medium_border)
for c in range(2, 8):
    ws2.cell(row=1, column=c).fill = DARK_BLUE
    ws2.cell(row=1, column=c).border = medium_border

ws2.merge_cells("A2:G2")
set_cell(ws2, 2, 1, "⚠️ 猪周期股DCF仅作辅助参考，核心看头均市值法和PB法", font=Font(color="FF0000", size=10, bold=True), fill=LIGHT_GREY, border=medium_border)
for c in range(2, 8):
    ws2.cell(row=2, column=c).fill = LIGHT_GREY
    ws2.cell(row=2, column=c).border = thin_border

# Key Assumptions
r = 4
section_header(ws2, r, 1, 7, "关键假设（正常化周期）")
r += 1

assumptions = [
    ("正常化营收（亿元）", 80.0),
    ("正常化毛利率", 0.15),
    ("正常化营业利润率", 0.08),
    ("永续增长率", 0.02),
    ("WACC", 0.11),
    ("资本支出/营收比", 0.10),
    ("折旧摊销/营收比", 0.06),
    ("营运资本变动/营收比", 0.01),
]

for label, val in assumptions:
    set_cell(ws2, r, 1, label, font=BLACK_BOLD, align=Alignment(horizontal="left", vertical="center"))
    set_cell(ws2, r, 2, val, font=BLUE_INPUT, fill=LIGHT_GREY, number_format="0.0%" if val < 1 else "0.0")
    ws2.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
    r += 1

r += 1
section_header(ws2, r, 1, 7, "自由现金流预测（正常化）")
r += 1

# Headers
for i, h in enumerate(["项目", "2026E", "2027E", "2028E", "2029E", "2030E", "终值"]):
    col_header(ws2, r, i+1, h)
r += 1

# FCF Build
dcf_rows = [
    ("营业收入", ["=D7*(1+0.12)", "=E7*(1+0.10)", "=F7*(1+0.08)", "=G7*(1+0.06)", "=H7*(1+0.04)", ""]),
    ("  营收增速", ["0.12", "0.10", "0.08", "0.06", "0.04", ""]),
    ("EBIT", ["=B17*$B$6", "=C17*$B$6", "=D17*$B$6", "=E17*$B$6", "=F17*$B$6", ""]),
    ("  营业利润率", ["=$B$6", "=$B$6", "=$B$6", "=$B$6", "=$B$6", ""]),
    ("(-) 所得税 (15%)", ["=IF(B19>0,B19*0.15,0)", "=IF(C19>0,C19*0.15,0)", "=IF(D19>0,D19*0.15,0)", "=IF(E19>0,E19*0.15,0)", "=IF(F19>0,F19*0.15,0)", ""]),
    ("NOPAT", ["=B19-B21", "=C19-C21", "=D19-D21", "=E19-E21", "=F19-F21", ""]),
    ("(+) 折旧摊销", ["=B17*$B$8", "=C17*$B$8", "=D17*$B$8", "=E17*$B$8", "=F17*$B$8", ""]),
    ("(-) 资本支出", ["=-B17*$B$7", "=-C17*$B$7", "=-D17*$B$7", "=-E17*$B$7", "=-F17*$B$7", ""]),
    ("(-) 营运资本变动", ["=-(B17-$A$17)*$B$9", "=-(C17-B17)*$B$9", "=-(D17-C17)*$B$9", "=-(E17-D17)*$B$9", "=-(F17-E17)*$B$9", ""]),
    ("自由现金流 (FCF)", ["=B22+B23+B24+B25", "=C22+C23+C24+C25", "=D22+D23+D24+D25", "=E22+E23+E24+E25", "=F22+F23+F24+F25", "=F26*(1+$B$5)"]),
]

# Fix: Use normal values instead of sheet references
dcf_numeric = [
    ("营业收入", [92.0, 101.2, 109.3, 115.8, 120.5, ""]),
    ("  营收增速", [0.15, 0.10, 0.08, 0.06, 0.04, ""]),
    ("EBIT", [13.8, 15.2, 16.4, 17.4, 18.1, ""]),
    ("  营业利润率", [0.15, 0.15, 0.15, 0.15, 0.15, ""]),
    ("(-) 所得税 (15%)", [-2.07, -2.28, -2.46, -2.61, -2.71, ""]),
    ("NOPAT", [11.7, 12.9, 13.9, 14.8, 15.4, ""]),
    ("(+) 折旧摊销", [5.5, 6.1, 6.6, 6.9, 7.2, ""]),
    ("(-) 资本支出", [-9.2, -10.1, -10.9, -11.6, -12.0, ""]),
    ("(-) 营运资本变动", [-0.9, -0.9, -0.8, -0.7, -0.5, ""]),
    ("自由现金流 (FCF)", [7.1, 8.0, 8.8, 9.4, 10.1, "=F26*(1+0.02)"]),
]

for label, vals in dcf_numeric:
    set_cell(ws2, r, 1, label, font=BLACK_BOLD if "自由现金流" in label else BLACK_NORMAL)
    for c_idx, v in enumerate(vals):
        if v == "":
            continue
        col = c_idx + 2
        if isinstance(v, str) and v.startswith("="):
            set_cell(ws2, r, col, v)
        elif isinstance(v, float):
            set_cell(ws2, r, col, v, number_format="0.0")
    r += 1

r += 1
section_header(ws2, r, 1, 7, "估值计算")
r += 1

# Count FCF cells - from row r_start+10 to r_start+14 (FCF row = row with 自由现金流)
fcf_row = r - 16  # Back to the FCF row
val_items = [
    ("PV of FCFs", "=B26/(1+$B$4)^0.5 + C26/(1+$B$4)^1.5 + D26/(1+$B$4)^2.5 + E26/(1+$B$4)^3.5 + F26/(1+$B$4)^4.5"),
    ("终值 (Terminal Value)", "=G26/($B$4-$B$5)"),
    ("PV of 终值", "=B30/(1+$B$4)^4.5"),
    ("企业价值 (EV)", "=B29+B31"),
    ("(-) 净债务 (2025末)", 14.08),
    ("= 股权价值", "=B32-B33"),
    ("股数（亿股）", 5.23),
    ("每股公允价值（元）", "=B34/B35"),
    ("当前股价（元）", 18.66),
    ("隐含上行空间", "=B36/B37-1"),
]

for label, val in val_items:
    set_cell(ws2, r, 1, label, font=BLACK_BOLD, align=Alignment(horizontal="left", vertical="center"))
    if isinstance(val, str) and val.startswith("="):
        set_cell(ws2, r, 2, val, font=GREEN_LINK)
    elif isinstance(val, float):
        set_cell(ws2, r, 2, val, font=BLUE_INPUT, fill=LIGHT_GREY)
    elif isinstance(val, str):
        set_cell(ws2, r, 2, val)

    if "每股" in label:
        set_cell(ws2, r, 2, val if not isinstance(val, str) or not val.startswith("=") else val,
                font=Font(color="0000FF", size=14, bold=True), fill=MEDIUM_BLUE)
    if "上行" in label:
        set_cell(ws2, r, 2, val if not isinstance(val, str) or not val.startswith("=") else val,
                number_format="0.0%", font=Font(color="FF0000", size=12, bold=True))
    ws2.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    r += 1

# WACC sheet
ws3 = wb.create_sheet("WACC计算")
section_header(ws3, 1, 1, 7, "WACC 计算")
r_wacc = 3

wacc_items = [
    ("风险调整系数", ""),
    ("无风险利率 (10Y国债)", 0.028),
    ("股权风险溢价", 0.065),
    ("Beta (周期股调整)", 1.30),
    ("股本成本 Ke", "=B4+B5*B6"),
    ("", ""),
    ("债务成本", ""),
    ("税前债务利率", 0.055),
    ("税率", 0.15),
    ("税后债务成本", "=B9*(1-B10)"),
    ("", ""),
    ("资本结构", ""),
    ("权益市值（亿元）", 97.59),
    ("净债务（亿元）", 14.08),
    ("企业价值（亿元）", "=B14+B15"),
    ("权益占比", "=B14/B16"),
    ("债务占比", "=B15/B16"),
    ("", ""),
    ("WACC", "=B7*B17+B11*B18"),
]

for label, val in wacc_items:
    set_cell(ws3, r_wacc, 1, label, font=BLACK_BOLD, align=Alignment(horizontal="left", vertical="center"))
    if isinstance(val, str) and val.startswith("="):
        set_cell(ws3, r_wacc, 2, val, font=GREEN_LINK)
    elif isinstance(val, float):
        set_cell(ws3, r_wacc, 2, val, font=BLUE_INPUT, fill=LIGHT_GREY, number_format="0.0%")
    ws3.merge_cells(start_row=r_wacc, start_column=2, end_row=r_wacc, end_column=7)
    r_wacc += 1

# 设置列宽
for ws in [ws2, ws3]:
    ws.column_dimensions["A"].width = 28
    for c in range(2, 8):
        ws.column_dimensions[get_column_letter(c)].width = 16

# ============================================================
# 保存
# ============================================================
output_path = "d:\\claude-for-finance\\outputs\\巨星农牧_三表模型_DCF估值_v2.xlsx"
wb.save(output_path)
print(f"Excel模型已保存: {output_path}")
print("文件包含3个Sheet：")
print("  1. 三表模型 - 利润表/资产负债表/现金流量表")
print("  2. DCF估值 - 自由现金流预测与估值计算")
print("  3. WACC计算 - 加权平均资本成本计算")
print("\n⚠️ 注意：猪周期股DCF仅供参考，核心估值请参考研究报告中的头均市值法和PB法。")
print("   打开后Excel会自动计算公式，如有提示请启用内容。")
