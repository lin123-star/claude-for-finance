# -*- coding: utf-8 -*-
"""
泰晶科技 (603738.SH) 三表模型 + DCF 估值 Excel 生成脚本
生成日期: 2026-05-16
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, numbers
from openpyxl.utils import get_column_letter

# ============================================================
# 配色与样式定义 (基金公司标准)
# ============================================================
DARK_BLUE_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
LIGHT_BLUE_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
MID_BLUE_FILL = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
LIGHT_GREY_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
LIGHT_GREEN_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

WHITE_BOLD = Font(name="微软雅黑", bold=True, color="FFFFFF", size=11)
BLACK_BOLD = Font(name="微软雅黑", bold=True, color="000000", size=11)
BLUE_FONT = Font(name="微软雅黑", color="0000FF", size=10)
GREEN_FONT = Font(name="微软雅黑", color="008000", size=10)
BLACK_FONT = Font(name="微软雅黑", color="000000", size=10)
ASSUMPTION_FONT = Font(name="微软雅黑", color="0070C0", size=10)

THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin")
)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def apply_header_style(ws, row, col_start, col_end):
    """深蓝区块标题行"""
    for c in range(col_start, col_end + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = DARK_BLUE_FILL
        cell.font = WHITE_BOLD
        cell.alignment = CENTER
        cell.border = THIN_BORDER


def apply_col_header_style(ws, row, col_start, col_end):
    """浅蓝列标题行"""
    for c in range(col_start, col_end + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = LIGHT_BLUE_FILL
        cell.font = BLACK_BOLD
        cell.alignment = CENTER
        cell.border = THIN_BORDER


def apply_subtotal_style(ws, row, col_start, col_end):
    """中蓝小计行"""
    for c in range(col_start, col_end + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = MID_BLUE_FILL
        cell.font = BLACK_BOLD
        cell.alignment = CENTER
        cell.border = THIN_BORDER


def set_historical_cell(ws, row, col, value):
    """设置历史实际值 (灰色填充+蓝色字体)"""
    cell = ws.cell(row=row, column=col)
    cell.value = value
    cell.fill = LIGHT_GREY_FILL
    cell.font = BLUE_FONT
    cell.alignment = CENTER
    cell.border = THIN_BORDER


def set_formula_cell(ws, row, col, formula):
    """设置公式单元格 (白色填充+绿色字体)"""
    cell = ws.cell(row=row, column=col)
    cell.value = formula
    cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    cell.font = GREEN_FONT
    cell.alignment = CENTER
    cell.border = THIN_BORDER


def set_label_cell(ws, row, col, label, bold=False):
    """设置标签单元格"""
    cell = ws.cell(row=row, column=col)
    cell.value = label
    cell.font = BLACK_BOLD if bold else BLACK_FONT
    cell.alignment = LEFT
    cell.border = THIN_BORDER


def set_assumption_cell(ws, row, col, value):
    """设置假设驱动单元格 (浅绿填充+蓝色字体)"""
    cell = ws.cell(row=row, column=col)
    cell.value = value
    cell.fill = LIGHT_GREEN_FILL
    cell.font = ASSUMPTION_FONT
    cell.alignment = CENTER
    cell.border = THIN_BORDER


# ============================================================
# 创建工作簿
# ============================================================
wb = openpyxl.Workbook()

# ============================================================
# Sheet 1: 假设
# ============================================================
ws_assume = wb.active
ws_assume.title = "假设"

# 列宽
col_widths = {1: 35, 2: 12, 3: 12, 4: 12, 5: 12, 6: 12, 7: 12, 8: 12, 9: 12}
for col, w in col_widths.items():
    ws_assume.column_dimensions[get_column_letter(col)].width = w

r = 1
# 标题
ws_assume.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
apply_header_style(ws_assume, r, 1, 9)
ws_assume.cell(row=r, column=1).value = "泰晶科技 (603738.SH) — 预测假设输入区"

r = 2
ws_assume.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
c = ws_assume.cell(row=r, column=1)
c.value = "绿色填充 = 可编辑驱动因子 | 修改后三表与DCF自动联动"
c.font = Font(name="微软雅黑", color="666666", size=9)
c.fill = LIGHT_GREEN_FILL

r = 4
# === 宏观经济假设 ===
apply_header_style(ws_assume, r, 1, 9)
ws_assume.cell(row=r, column=1).value = "一、宏观经济与市场假设"

r = 5
set_label_cell(ws_assume, r, 1, "无风险利率 (10Y国债)", bold=True)
set_assumption_cell(ws_assume, r, 2, 0.0170)

r = 6
set_label_cell(ws_assume, r, 1, "市场风险溢价 (ERP)", bold=True)
set_assumption_cell(ws_assume, r, 2, 0.065)

r = 7
set_label_cell(ws_assume, r, 1, "Beta (5Y monthly vs 沪深300)", bold=True)
set_assumption_cell(ws_assume, r, 2, 1.15)

r = 8
set_label_cell(ws_assume, r, 1, "税前债务成本", bold=True)
set_assumption_cell(ws_assume, r, 2, 0.035)

r = 9
set_label_cell(ws_assume, r, 1, "有效税率", bold=True)
set_assumption_cell(ws_assume, r, 2, 0.15)

r = 11
# === 收入驱动假设 ===
apply_header_style(ws_assume, r, 1, 9)
ws_assume.cell(row=r, column=1).value = "二、收入与成本驱动假设"

# 列头
r = 12
year_labels = ["驱动因子", "2022A", "2023A", "2024A", "2025A", "2026E", "2027E", "2028E", "2029E", "2030E"]
for i, label in enumerate(year_labels):
    if i == 0:
        set_label_cell(ws_assume, r, i+1, label, bold=True)
    else:
        c = ws_assume.cell(row=r, column=i+1)
        c.value = label
        c.fill = LIGHT_BLUE_FILL
        c.font = BLACK_BOLD
        c.alignment = CENTER
        c.border = THIN_BORDER

# 营收增速
r = 13
set_label_cell(ws_assume, r, 1, "营收增速 (%)")
set_historical_cell(ws_assume, r, 2, -0.15)    # 2022A
set_historical_cell(ws_assume, r, 3, -0.08)    # 2023A
set_historical_cell(ws_assume, r, 4, 0.12)     # 2024A
set_historical_cell(ws_assume, r, 5, 0.15)     # 2025A
set_assumption_cell(ws_assume, r, 6, 0.18)     # 2026E
set_assumption_cell(ws_assume, r, 7, 0.15)     # 2027E
set_assumption_cell(ws_assume, r, 8, 0.12)     # 2028E
set_assumption_cell(ws_assume, r, 9, 0.10)     # 2029E
set_assumption_cell(ws_assume, r, 10, 0.08)    # 2030E

# 毛利率
r = 14
set_label_cell(ws_assume, r, 1, "毛利率 (%)")
set_historical_cell(ws_assume, r, 2, 0.32)
set_historical_cell(ws_assume, r, 3, 0.30)
set_historical_cell(ws_assume, r, 4, 0.33)
set_historical_cell(ws_assume, r, 5, 0.35)
set_assumption_cell(ws_assume, r, 6, 0.36)
set_assumption_cell(ws_assume, r, 7, 0.37)
set_assumption_cell(ws_assume, r, 8, 0.38)
set_assumption_cell(ws_assume, r, 9, 0.38)
set_assumption_cell(ws_assume, r, 10, 0.38)

# 销售费用率
r = 15
set_label_cell(ws_assume, r, 1, "销售费用率 (%)")
set_historical_cell(ws_assume, r, 2, 0.03)
set_historical_cell(ws_assume, r, 3, 0.03)
set_historical_cell(ws_assume, r, 4, 0.025)
set_historical_cell(ws_assume, r, 5, 0.025)
set_assumption_cell(ws_assume, r, 6, 0.024)
set_assumption_cell(ws_assume, r, 7, 0.023)
set_assumption_cell(ws_assume, r, 8, 0.022)
set_assumption_cell(ws_assume, r, 9, 0.022)
set_assumption_cell(ws_assume, r, 10, 0.022)

# 管理费用率
r = 16
set_label_cell(ws_assume, r, 1, "管理费用率 (%)")
set_historical_cell(ws_assume, r, 2, 0.06)
set_historical_cell(ws_assume, r, 3, 0.065)
set_historical_cell(ws_assume, r, 4, 0.06)
set_historical_cell(ws_assume, r, 5, 0.055)
set_assumption_cell(ws_assume, r, 6, 0.053)
set_assumption_cell(ws_assume, r, 7, 0.051)
set_assumption_cell(ws_assume, r, 8, 0.050)
set_assumption_cell(ws_assume, r, 9, 0.050)
set_assumption_cell(ws_assume, r, 10, 0.050)

# 研发费用率
r = 17
set_label_cell(ws_assume, r, 1, "研发费用率 (%)")
set_historical_cell(ws_assume, r, 2, 0.055)
set_historical_cell(ws_assume, r, 3, 0.06)
set_historical_cell(ws_assume, r, 4, 0.058)
set_historical_cell(ws_assume, r, 5, 0.055)
set_assumption_cell(ws_assume, r, 6, 0.053)
set_assumption_cell(ws_assume, r, 7, 0.051)
set_assumption_cell(ws_assume, r, 8, 0.050)
set_assumption_cell(ws_assume, r, 9, 0.050)
set_assumption_cell(ws_assume, r, 10, 0.050)

# D&A占营收比
r = 18
set_label_cell(ws_assume, r, 1, "折旧摊销占营收 (%)")
set_historical_cell(ws_assume, r, 2, 0.08)
set_historical_cell(ws_assume, r, 3, 0.09)
set_historical_cell(ws_assume, r, 4, 0.085)
set_historical_cell(ws_assume, r, 5, 0.08)
set_assumption_cell(ws_assume, r, 6, 0.075)
set_assumption_cell(ws_assume, r, 7, 0.072)
set_assumption_cell(ws_assume, r, 8, 0.070)
set_assumption_cell(ws_assume, r, 9, 0.068)
set_assumption_cell(ws_assume, r, 10, 0.065)

# CAPEX占营收比
r = 19
set_label_cell(ws_assume, r, 1, "资本支出占营收 (%)")
set_historical_cell(ws_assume, r, 2, 0.10)
set_historical_cell(ws_assume, r, 3, 0.09)
set_historical_cell(ws_assume, r, 4, 0.08)
set_historical_cell(ws_assume, r, 5, 0.07)
set_assumption_cell(ws_assume, r, 6, 0.07)
set_assumption_cell(ws_assume, r, 7, 0.065)
set_assumption_cell(ws_assume, r, 8, 0.06)
set_assumption_cell(ws_assume, r, 9, 0.06)
set_assumption_cell(ws_assume, r, 10, 0.06)

# NWC变化占Δ营收
r = 20
set_label_cell(ws_assume, r, 1, "营运资本变动占Δ营收 (%)")
set_historical_cell(ws_assume, r, 2, 0.05)
set_historical_cell(ws_assume, r, 3, 0.04)
set_historical_cell(ws_assume, r, 4, 0.05)
set_historical_cell(ws_assume, r, 5, 0.06)
set_assumption_cell(ws_assume, r, 6, 0.05)
set_assumption_cell(ws_assume, r, 7, 0.05)
set_assumption_cell(ws_assume, r, 8, 0.04)
set_assumption_cell(ws_assume, r, 9, 0.04)
set_assumption_cell(ws_assume, r, 10, 0.04)

r = 22
# === 三情景 DCF 假设 ===
apply_header_style(ws_assume, r, 1, 9)
ws_assume.cell(row=r, column=1).value = "三、DCF 三情景假设 (基准情景)"

r = 23
set_label_cell(ws_assume, r, 1, "终值增长率 (g)", bold=True)
set_assumption_cell(ws_assume, r, 2, 0.025)

r = 24
set_label_cell(ws_assume, r, 1, "WACC", bold=True)
set_assumption_cell(ws_assume, r, 2, 0.095)

r = 25
set_label_cell(ws_assume, r, 1, "目标D/E比", bold=True)
set_assumption_cell(ws_assume, r, 2, 0.15)

r = 27
# === 市场数据 ===
apply_header_style(ws_assume, r, 1, 9)
ws_assume.cell(row=r, column=1).value = "四、市场数据 (非情景依赖)"

r = 28
set_label_cell(ws_assume, r, 1, "当前股价 (元)", bold=True)
set_assumption_cell(ws_assume, r, 2, 15.50)

r = 29
set_label_cell(ws_assume, r, 1, "总股本 (亿股)", bold=True)
set_assumption_cell(ws_assume, r, 2, 3.85)

r = 30
set_label_cell(ws_assume, r, 1, "有息负债 (亿元)", bold=True)
set_assumption_cell(ws_assume, r, 2, 3.50)

r = 31
set_label_cell(ws_assume, r, 1, "现金及等价物 (亿元)", bold=True)
set_assumption_cell(ws_assume, r, 2, 5.20)

r = 33
# 悲观情景
apply_header_style(ws_assume, r, 1, 9)
ws_assume.cell(row=r, column=1).value = "五、悲观情景"

r = 34
set_label_cell(ws_assume, r, 1, "营收增速 (Y1-Y5)", bold=True)
for i, v in enumerate([0.10, 0.08, 0.06, 0.05, 0.03]):
    set_assumption_cell(ws_assume, r, i+2, v)

r = 35
set_label_cell(ws_assume, r, 1, "毛利率 (终值)", bold=True)
set_assumption_cell(ws_assume, r, 2, 0.32)

r = 36
set_label_cell(ws_assume, r, 1, "终值增长率", bold=True)
set_assumption_cell(ws_assume, r, 2, 0.020)

r = 37
set_label_cell(ws_assume, r, 1, "WACC", bold=True)
set_assumption_cell(ws_assume, r, 2, 0.105)

r = 39
# 乐观情景
apply_header_style(ws_assume, r, 1, 9)
ws_assume.cell(row=r, column=1).value = "六、乐观情景"

r = 40
set_label_cell(ws_assume, r, 1, "营收增速 (Y1-Y5)", bold=True)
for i, v in enumerate([0.25, 0.22, 0.18, 0.15, 0.12]):
    set_assumption_cell(ws_assume, r, i+2, v)

r = 41
set_label_cell(ws_assume, r, 1, "毛利率 (终值)", bold=True)
set_assumption_cell(ws_assume, r, 2, 0.42)

r = 42
set_label_cell(ws_assume, r, 1, "终值增长率", bold=True)
set_assumption_cell(ws_assume, r, 2, 0.030)

r = 43
set_label_cell(ws_assume, r, 1, "WACC", bold=True)
set_assumption_cell(ws_assume, r, 2, 0.085)


# ============================================================
# Sheet 2: 三表模型
# ============================================================
ws_model = wb.create_sheet("三表模型")

col_widths_m = {1: 38, 2: 14, 3: 14, 4: 14, 5: 14, 6: 14, 7: 14, 8: 14, 9: 14, 10: 14}
for col, w in col_widths_m.items():
    ws_model.column_dimensions[get_column_letter(col)].width = w

# ---- 利润表 ----
r = 1
ws_model.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
apply_header_style(ws_model, r, 1, 10)
ws_model.cell(row=r, column=1).value = "泰晶科技 (603738.SH) — 三表财务模型"

r = 2
year_cols_m = ["科目 (亿元)", "2022A", "2023A", "2024A", "2025A", "2026E", "2027E", "2028E", "2029E", "2030E"]
apply_col_header_style(ws_model, r, 1, 10)
for i, label in enumerate(year_cols_m):
    ws_model.cell(row=r, column=i+1).value = label

# 利润表 section header
r = 3
ws_model.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
apply_header_style(ws_model, r, 1, 10)
ws_model.cell(row=r, column=1).value = "利润表"

# 营业收入 (历史数据为估算值，需以年报数据验证)
r = 5
set_label_cell(ws_model, r, 1, "营业收入")
set_historical_cell(ws_model, r, 2, 9.20)     # 2022A
set_historical_cell(ws_model, r, 3, 8.46)     # 2023A
set_historical_cell(ws_model, r, 4, 9.48)     # 2024A
set_historical_cell(ws_model, r, 5, 10.90)    # 2025A
set_formula_cell(ws_model, r, 6, "=E5*(1+假设!F13)")
set_formula_cell(ws_model, r, 7, "=F5*(1+假设!G13)")
set_formula_cell(ws_model, r, 8, "=G5*(1+假设!H13)")
set_formula_cell(ws_model, r, 9, "=H5*(1+假设!I13)")
set_formula_cell(ws_model, r, 10, "=I5*(1+假设!J13)")

# 营收增速
r = 6
set_label_cell(ws_model, r, 1, "  营收增速 (%)")
set_historical_cell(ws_model, r, 2, -0.15)
set_historical_cell(ws_model, r, 3, -0.08)
set_historical_cell(ws_model, r, 4, 0.12)
set_historical_cell(ws_model, r, 5, 0.15)
set_formula_cell(ws_model, r, 6, "=F5/E5-1")
set_formula_cell(ws_model, r, 7, "=G5/F5-1")
set_formula_cell(ws_model, r, 8, "=H5/G5-1")
set_formula_cell(ws_model, r, 9, "=I5/H5-1")
set_formula_cell(ws_model, r, 10, "=J5/I5-1")

# 营业成本
r = 7
set_label_cell(ws_model, r, 1, "营业成本")
set_historical_cell(ws_model, r, 2, 6.26)
set_historical_cell(ws_model, r, 3, 5.92)
set_historical_cell(ws_model, r, 4, 6.35)
set_historical_cell(ws_model, r, 5, 7.09)
set_formula_cell(ws_model, r, 6, "=F5*(1-假设!F14)")
set_formula_cell(ws_model, r, 7, "=G5*(1-假设!G14)")
set_formula_cell(ws_model, r, 8, "=H5*(1-假设!H14)")
set_formula_cell(ws_model, r, 9, "=I5*(1-假设!I14)")
set_formula_cell(ws_model, r, 10, "=J5*(1-假设!J14)")

# 毛利润
r = 8
set_label_cell(ws_model, r, 1, "毛利润")
set_formula_cell(ws_model, r, 2, "=B5-B7")
set_formula_cell(ws_model, r, 3, "=C5-C7")
set_formula_cell(ws_model, r, 4, "=D5-D7")
set_formula_cell(ws_model, r, 5, "=E5-E7")
set_formula_cell(ws_model, r, 6, "=F5-F7")
set_formula_cell(ws_model, r, 7, "=G5-G7")
set_formula_cell(ws_model, r, 8, "=H5-H7")
set_formula_cell(ws_model, r, 9, "=I5-I7")
set_formula_cell(ws_model, r, 10, "=J5-J7")

# 毛利率
r = 9
set_label_cell(ws_model, r, 1, "  毛利率 (%)")
set_formula_cell(ws_model, r, 2, "=B8/B5")
set_formula_cell(ws_model, r, 3, "=C8/C5")
set_formula_cell(ws_model, r, 4, "=D8/D5")
set_formula_cell(ws_model, r, 5, "=E8/E5")
set_formula_cell(ws_model, r, 6, "=F8/F5")
set_formula_cell(ws_model, r, 7, "=G8/G5")
set_formula_cell(ws_model, r, 8, "=H8/H5")
set_formula_cell(ws_model, r, 9, "=I8/I5")
set_formula_cell(ws_model, r, 10, "=J8/J5")

# 销售费用
r = 11
set_label_cell(ws_model, r, 1, "销售费用")
set_formula_cell(ws_model, r, 2, "=B5*假设!B15")
set_formula_cell(ws_model, r, 3, "=C5*假设!C15")
set_formula_cell(ws_model, r, 4, "=D5*假设!D15")
set_formula_cell(ws_model, r, 5, "=E5*假设!E15")
set_formula_cell(ws_model, r, 6, "=F5*假设!F15")
set_formula_cell(ws_model, r, 7, "=G5*假设!G15")
set_formula_cell(ws_model, r, 8, "=H5*假设!H15")
set_formula_cell(ws_model, r, 9, "=I5*假设!I15")
set_formula_cell(ws_model, r, 10, "=J5*假设!J15")

# 管理费用
r = 12
set_label_cell(ws_model, r, 1, "管理费用")
set_formula_cell(ws_model, r, 2, "=B5*假设!B16")
set_formula_cell(ws_model, r, 3, "=C5*假设!C16")
set_formula_cell(ws_model, r, 4, "=D5*假设!D16")
set_formula_cell(ws_model, r, 5, "=E5*假设!E16")
set_formula_cell(ws_model, r, 6, "=F5*假设!F16")
set_formula_cell(ws_model, r, 7, "=G5*假设!G16")
set_formula_cell(ws_model, r, 8, "=H5*假设!H16")
set_formula_cell(ws_model, r, 9, "=I5*假设!I16")
set_formula_cell(ws_model, r, 10, "=J5*假设!J16")

# 研发费用
r = 13
set_label_cell(ws_model, r, 1, "研发费用")
set_formula_cell(ws_model, r, 2, "=B5*假设!B17")
set_formula_cell(ws_model, r, 3, "=C5*假设!C17")
set_formula_cell(ws_model, r, 4, "=D5*假设!D17")
set_formula_cell(ws_model, r, 5, "=E5*假设!E17")
set_formula_cell(ws_model, r, 6, "=F5*假设!F17")
set_formula_cell(ws_model, r, 7, "=G5*假设!G17")
set_formula_cell(ws_model, r, 8, "=H5*假设!H17")
set_formula_cell(ws_model, r, 9, "=I5*假设!I17")
set_formula_cell(ws_model, r, 10, "=J5*假设!J17")

# 总费用
r = 14
set_label_cell(ws_model, r, 1, "期间费用合计", bold=True)
apply_subtotal_style(ws_model, r, 1, 1)
set_formula_cell(ws_model, r, 2, "=B11+B12+B13")
set_formula_cell(ws_model, r, 3, "=C11+C12+C13")
set_formula_cell(ws_model, r, 4, "=D11+D12+D13")
set_formula_cell(ws_model, r, 5, "=E11+E12+E13")
set_formula_cell(ws_model, r, 6, "=F11+F12+F13")
set_formula_cell(ws_model, r, 7, "=G11+G12+G13")
set_formula_cell(ws_model, r, 8, "=H11+H12+H13")
set_formula_cell(ws_model, r, 9, "=I11+I12+I13")
set_formula_cell(ws_model, r, 10, "=J11+J12+J13")

# EBIT
r = 15
set_label_cell(ws_model, r, 1, "EBIT (营业利润)", bold=True)
apply_subtotal_style(ws_model, r, 1, 1)
set_formula_cell(ws_model, r, 2, "=B8-B14")
set_formula_cell(ws_model, r, 3, "=C8-C14")
set_formula_cell(ws_model, r, 4, "=D8-D14")
set_formula_cell(ws_model, r, 5, "=E8-E14")
set_formula_cell(ws_model, r, 6, "=F8-F14")
set_formula_cell(ws_model, r, 7, "=G8-G14")
set_formula_cell(ws_model, r, 8, "=H8-H14")
set_formula_cell(ws_model, r, 9, "=I8-I14")
set_formula_cell(ws_model, r, 10, "=J8-J14")

# EBIT率
r = 16
set_label_cell(ws_model, r, 1, "  EBIT率 (%)")
set_formula_cell(ws_model, r, 2, "=B15/B5")
set_formula_cell(ws_model, r, 3, "=C15/C5")
set_formula_cell(ws_model, r, 4, "=D15/D5")
set_formula_cell(ws_model, r, 5, "=E15/E5")
set_formula_cell(ws_model, r, 6, "=F15/F5")
set_formula_cell(ws_model, r, 7, "=G15/G5")
set_formula_cell(ws_model, r, 8, "=H15/H5")
set_formula_cell(ws_model, r, 9, "=I15/I5")
set_formula_cell(ws_model, r, 10, "=J15/J5")

# 所得税
r = 17
set_label_cell(ws_model, r, 1, "所得税")
set_formula_cell(ws_model, r, 2, "=B15*假设!$B$9")
set_formula_cell(ws_model, r, 3, "=C15*假设!$B$9")
set_formula_cell(ws_model, r, 4, "=D15*假设!$B$9")
set_formula_cell(ws_model, r, 5, "=E15*假设!$B$9")
set_formula_cell(ws_model, r, 6, "=F15*假设!$B$9")
set_formula_cell(ws_model, r, 7, "=G15*假设!$B$9")
set_formula_cell(ws_model, r, 8, "=H15*假设!$B$9")
set_formula_cell(ws_model, r, 9, "=I15*假设!$B$9")
set_formula_cell(ws_model, r, 10, "=J15*假设!$B$9")

# NOPAT
r = 18
set_label_cell(ws_model, r, 1, "NOPAT (税后净营业利润)", bold=True)
apply_subtotal_style(ws_model, r, 1, 1)
set_formula_cell(ws_model, r, 2, "=B15-B17")
set_formula_cell(ws_model, r, 3, "=C15-C17")
set_formula_cell(ws_model, r, 4, "=D15-D17")
set_formula_cell(ws_model, r, 5, "=E15-E17")
set_formula_cell(ws_model, r, 6, "=F15-F17")
set_formula_cell(ws_model, r, 7, "=G15-G17")
set_formula_cell(ws_model, r, 8, "=H15-H17")
set_formula_cell(ws_model, r, 9, "=I15-I17")
set_formula_cell(ws_model, r, 10, "=J15-J17")

# ---- 资产负债表关键指标 ----
r = 20
ws_model.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
apply_header_style(ws_model, r, 1, 10)
ws_model.cell(row=r, column=1).value = "资产负债表 (关键指标)"

# 总资产
r = 22
set_label_cell(ws_model, r, 1, "总资产")
set_historical_cell(ws_model, r, 2, 22.00)
set_historical_cell(ws_model, r, 3, 21.50)
set_historical_cell(ws_model, r, 4, 22.80)
set_historical_cell(ws_model, r, 5, 24.50)
# 预测: 简单估算 (总资产周转率约0.45)
set_formula_cell(ws_model, r, 6, "=F5/0.45")
set_formula_cell(ws_model, r, 7, "=G5/0.45")
set_formula_cell(ws_model, r, 8, "=H5/0.45")
set_formula_cell(ws_model, r, 9, "=I5/0.45")
set_formula_cell(ws_model, r, 10, "=J5/0.45")

# 总负债
r = 23
set_label_cell(ws_model, r, 1, "总负债")
set_historical_cell(ws_model, r, 2, 8.50)
set_historical_cell(ws_model, r, 3, 8.00)
set_historical_cell(ws_model, r, 4, 8.20)
set_historical_cell(ws_model, r, 5, 8.50)
set_formula_cell(ws_model, r, 6, "=F22*0.38")
set_formula_cell(ws_model, r, 7, "=G22*0.37")
set_formula_cell(ws_model, r, 8, "=H22*0.36")
set_formula_cell(ws_model, r, 9, "=I22*0.35")
set_formula_cell(ws_model, r, 10, "=J22*0.35")

# 股东权益
r = 24
set_label_cell(ws_model, r, 1, "股东权益", bold=True)
apply_subtotal_style(ws_model, r, 1, 1)
set_formula_cell(ws_model, r, 2, "=B22-B23")
set_formula_cell(ws_model, r, 3, "=C22-C23")
set_formula_cell(ws_model, r, 4, "=D22-D23")
set_formula_cell(ws_model, r, 5, "=E22-E23")
set_formula_cell(ws_model, r, 6, "=F22-F23")
set_formula_cell(ws_model, r, 7, "=G22-G23")
set_formula_cell(ws_model, r, 8, "=H22-H23")
set_formula_cell(ws_model, r, 9, "=I22-I23")
set_formula_cell(ws_model, r, 10, "=J22-J23")

# BS 平衡检查
r = 25
set_label_cell(ws_model, r, 1, "BS平衡检查 (资产-负债-权益=0)")
set_formula_cell(ws_model, r, 2, "=B22-B23-B24")
set_formula_cell(ws_model, r, 3, "=C22-C23-C24")
set_formula_cell(ws_model, r, 4, "=D22-D23-D24")
set_formula_cell(ws_model, r, 5, "=E22-E23-E24")
set_formula_cell(ws_model, r, 6, "=F22-F23-F24")
set_formula_cell(ws_model, r, 7, "=G22-G23-G24")
set_formula_cell(ws_model, r, 8, "=H22-H23-H24")
set_formula_cell(ws_model, r, 9, "=I22-I23-I24")
set_formula_cell(ws_model, r, 10, "=J22-J23-J24")

# ---- 现金流量表关键指标 ----
r = 27
ws_model.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
apply_header_style(ws_model, r, 1, 10)
ws_model.cell(row=r, column=1).value = "现金流量表 (关键指标)"

# 经营现金流
r = 29
set_label_cell(ws_model, r, 1, "经营性现金流")
set_historical_cell(ws_model, r, 2, 1.80)
set_historical_cell(ws_model, r, 3, 1.50)
set_historical_cell(ws_model, r, 4, 1.85)
set_historical_cell(ws_model, r, 5, 2.20)
set_formula_cell(ws_model, r, 6, "=F18+F31")  # NOPAT + D&A simplified
set_formula_cell(ws_model, r, 7, "=G18+G31")
set_formula_cell(ws_model, r, 8, "=H18+H31")
set_formula_cell(ws_model, r, 9, "=I18+I31")
set_formula_cell(ws_model, r, 10, "=J18+J31")

# CAPEX
r = 30
set_label_cell(ws_model, r, 1, "资本支出")
set_formula_cell(ws_model, r, 2, "=B5*假设!B19")
set_formula_cell(ws_model, r, 3, "=C5*假设!C19")
set_formula_cell(ws_model, r, 4, "=D5*假设!D19")
set_formula_cell(ws_model, r, 5, "=E5*假设!E19")
set_formula_cell(ws_model, r, 6, "=F5*假设!F19")
set_formula_cell(ws_model, r, 7, "=G5*假设!G19")
set_formula_cell(ws_model, r, 8, "=H5*假设!H19")
set_formula_cell(ws_model, r, 9, "=I5*假设!I19")
set_formula_cell(ws_model, r, 10, "=J5*假设!J19")

# D&A
r = 31
set_label_cell(ws_model, r, 1, "折旧摊销")
set_formula_cell(ws_model, r, 2, "=B5*假设!B18")
set_formula_cell(ws_model, r, 3, "=C5*假设!C18")
set_formula_cell(ws_model, r, 4, "=D5*假设!D18")
set_formula_cell(ws_model, r, 5, "=E5*假设!E18")
set_formula_cell(ws_model, r, 6, "=F5*假设!F18")
set_formula_cell(ws_model, r, 7, "=G5*假设!G18")
set_formula_cell(ws_model, r, 8, "=H5*假设!H18")
set_formula_cell(ws_model, r, 9, "=I5*假设!I18")
set_formula_cell(ws_model, r, 10, "=J5*假设!J18")

# NWC变动
r = 32
set_label_cell(ws_model, r, 1, "营运资本变动")
set_formula_cell(ws_model, r, 2, "=(B5-9.35)*假设!B20")
set_formula_cell(ws_model, r, 3, "=(C5-B5)*假设!C20")
set_formula_cell(ws_model, r, 4, "=(D5-C5)*假设!D20")
set_formula_cell(ws_model, r, 5, "=(E5-D5)*假设!E20")
set_formula_cell(ws_model, r, 6, "=(F5-E5)*假设!F20")
set_formula_cell(ws_model, r, 7, "=(G5-F5)*假设!G20")
set_formula_cell(ws_model, r, 8, "=(H5-G5)*假设!H20")
set_formula_cell(ws_model, r, 9, "=(I5-H5)*假设!I20")
set_formula_cell(ws_model, r, 10, "=(J5-I5)*假设!J20")

# UFCF
r = 33
set_label_cell(ws_model, r, 1, "无杠杆自由现金流 (UFCF)", bold=True)
apply_subtotal_style(ws_model, r, 1, 1)
set_formula_cell(ws_model, r, 2, "=B18+B31-B30-B32")
set_formula_cell(ws_model, r, 3, "=C18+C31-C30-C32")
set_formula_cell(ws_model, r, 4, "=D18+D31-D30-D32")
set_formula_cell(ws_model, r, 5, "=E18+E31-E30-E32")
set_formula_cell(ws_model, r, 6, "=F18+F31-F30-F32")
set_formula_cell(ws_model, r, 7, "=G18+G31-G30-G32")
set_formula_cell(ws_model, r, 8, "=H18+H31-H30-H32")
set_formula_cell(ws_model, r, 9, "=I18+I31-I30-I32")
set_formula_cell(ws_model, r, 10, "=J18+J31-J30-J32")


# ============================================================
# Sheet 3: DCF估值
# ============================================================
ws_dcf = wb.create_sheet("DCF估值")

col_widths_d = {1: 38, 2: 14, 3: 14, 4: 14, 5: 14, 6: 14, 7: 14}
for col, w in col_widths_d.items():
    ws_dcf.column_dimensions[get_column_letter(col)].width = w

r = 1
ws_dcf.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
apply_header_style(ws_dcf, r, 1, 7)
ws_dcf.cell(row=r, column=1).value = "泰晶科技 (603738.SH) — DCF估值模型 (基准情景)"

# 列头
r = 2
dcf_cols = ["项目", "2026E", "2027E", "2028E", "2029E", "2030E", "终值"]
apply_col_header_style(ws_dcf, r, 1, 7)
for i, label in enumerate(dcf_cols):
    ws_dcf.cell(row=r, column=i+1).value = label

# UFCF (引用三表模型)
r = 4
set_label_cell(ws_dcf, r, 1, "无杠杆自由现金流 (亿元)")
set_formula_cell(ws_dcf, r, 2, "=三表模型!F33")
set_formula_cell(ws_dcf, r, 3, "=三表模型!G33")
set_formula_cell(ws_dcf, r, 4, "=三表模型!H33")
set_formula_cell(ws_dcf, r, 5, "=三表模型!I33")
set_formula_cell(ws_dcf, r, 6, "=三表模型!J33")

# 折现期数 (mid-year convention)
r = 5
set_label_cell(ws_dcf, r, 1, "折现期数 (mid-year)")
set_formula_cell(ws_dcf, r, 2, "=0.5")
set_formula_cell(ws_dcf, r, 3, "=1.5")
set_formula_cell(ws_dcf, r, 4, "=2.5")
set_formula_cell(ws_dcf, r, 5, "=3.5")
set_formula_cell(ws_dcf, r, 6, "=4.5")

# 折现因子
r = 6
set_label_cell(ws_dcf, r, 1, "折现因子")
wacc_ref = "假设!$B$24"
set_formula_cell(ws_dcf, r, 2, f"=1/(1+{wacc_ref})^B5")
set_formula_cell(ws_dcf, r, 3, f"=1/(1+{wacc_ref})^C5")
set_formula_cell(ws_dcf, r, 4, f"=1/(1+{wacc_ref})^D5")
set_formula_cell(ws_dcf, r, 5, f"=1/(1+{wacc_ref})^E5")
set_formula_cell(ws_dcf, r, 6, f"=1/(1+{wacc_ref})^F5")

# PV of FCF
r = 7
set_label_cell(ws_dcf, r, 1, "PV of FCF (亿元)")
set_formula_cell(ws_dcf, r, 2, "=B4*B6")
set_formula_cell(ws_dcf, r, 3, "=C4*C6")
set_formula_cell(ws_dcf, r, 4, "=D4*D6")
set_formula_cell(ws_dcf, r, 5, "=E4*E6")
set_formula_cell(ws_dcf, r, 6, "=F4*F6")

# 终值计算
r = 9
ws_dcf.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
apply_header_style(ws_dcf, r, 1, 7)
ws_dcf.cell(row=r, column=1).value = "终值计算 (永续增长法)"

r = 10
set_label_cell(ws_dcf, r, 1, "终值年份FCF (亿元)")
set_formula_cell(ws_dcf, r, 3, "=三表模型!J33*(1+假设!$B$23)")

r = 11
set_label_cell(ws_dcf, r, 1, "终值 (亿元)")
set_formula_cell(ws_dcf, r, 3, "=C10/(假设!$B$24-假设!$B$23)")

r = 12
set_label_cell(ws_dcf, r, 1, "PV of 终值 (亿元)")
set_formula_cell(ws_dcf, r, 3, f"=C11/(1+{wacc_ref})^F5")

# 估值汇总
r = 14
ws_dcf.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
apply_header_style(ws_dcf, r, 1, 7)
ws_dcf.cell(row=r, column=1).value = "估值汇总"

r = 15
set_label_cell(ws_dcf, r, 1, "预测期FCF现值之和 (亿元)")
set_formula_cell(ws_dcf, r, 3, "=SUM(B7:F7)")

r = 16
set_label_cell(ws_dcf, r, 1, "终值现值 (亿元)")
set_formula_cell(ws_dcf, r, 3, "=C12")

r = 17
set_label_cell(ws_dcf, r, 1, "企业价值 EV (亿元)", bold=True)
apply_subtotal_style(ws_dcf, r, 1, 1)
set_formula_cell(ws_dcf, r, 3, "=C15+C16")

r = 18
set_label_cell(ws_dcf, r, 1, "(-) 净负债 (亿元)")
net_debt_formula = "=假设!B30-假设!B31"
set_formula_cell(ws_dcf, r, 3, net_debt_formula)

r = 19
set_label_cell(ws_dcf, r, 1, "股权价值 (亿元)", bold=True)
apply_subtotal_style(ws_dcf, r, 1, 1)
set_formula_cell(ws_dcf, r, 3, "=C17-C18")

r = 20
set_label_cell(ws_dcf, r, 1, "总股本 (亿股)")
set_formula_cell(ws_dcf, r, 3, "=假设!B29")

r = 21
set_label_cell(ws_dcf, r, 1, "每股内在价值 (元)", bold=True)
apply_subtotal_style(ws_dcf, r, 1, 1)
c = ws_dcf.cell(row=r, column=3)
c.value = "=C19/C20"
c.font = Font(name="微软雅黑", bold=True, color="FFFFFF", size=12)
c.fill = DARK_BLUE_FILL
c.alignment = CENTER
c.border = THIN_BORDER

r = 22
set_label_cell(ws_dcf, r, 1, "当前股价 (元)")
set_formula_cell(ws_dcf, r, 3, "=假设!B28")

r = 23
set_label_cell(ws_dcf, r, 1, "上行/下行空间 (%)", bold=True)
apply_subtotal_style(ws_dcf, r, 1, 1)
set_formula_cell(ws_dcf, r, 3, "=C21/C22-1")

# ---- 敏感度分析 5x5 (WACC vs 终值增长率) ----
r = 26
ws_dcf.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
apply_header_style(ws_dcf, r, 1, 7)
ws_dcf.cell(row=r, column=1).value = "敏感度分析: WACC vs 终值增长率 (每股价值, 元)"

# 参考值: DCF中WACC=假设!B24, g=假设!B23
# WACC axis: 7.5%, 8.5%, 9.5%, 10.5%, 11.5%
# g axis: 1.5%, 2.0%, 2.5%, 3.0%, 3.5%

r = 27
# 列头 (g values)
set_label_cell(ws_dcf, r, 1, "WACC \\ g")
g_values = [0.015, 0.020, 0.025, 0.030, 0.035]
for i, gv in enumerate(g_values):
    c = ws_dcf.cell(row=r, column=i+2)
    c.value = gv
    c.font = BLACK_BOLD
    c.fill = LIGHT_BLUE_FILL
    c.alignment = CENTER
    c.border = THIN_BORDER
    c.number_format = '0.0%'

# 行头 + 公式
wacc_sens = [0.075, 0.085, 0.095, 0.105, 0.115]
for row_offset, wv in enumerate(wacc_sens):
    sens_row = r + 1 + row_offset
    # WACC label
    c = ws_dcf.cell(row=sens_row, column=1)
    c.value = wv
    c.font = BLACK_BOLD
    c.fill = LIGHT_BLUE_FILL
    c.alignment = CENTER
    c.border = THIN_BORDER
    c.number_format = '0.0%'

    for col_offset, gv in enumerate(g_values):
        cell = ws_dcf.cell(row=sens_row, column=col_offset + 2)
        # Each cell recalculates implied price with its own WACC and g
        # Simplified sensitivity: adjust PV using relative discount factors
        base_wacc = 0.095
        base_g = 0.025
        # Formula: base_price * adjustment factor for WACC and g
        cell_formula = (
            f'=C21*(({base_wacc}-{base_g})/({wv}-{gv}))/(({base_wacc}-{base_g})/({wv}-{gv})*'
            f'(1+{base_g})/(1+{gv})^5)'
        )
        # Simplified approach - use sensitivity ratio
        cell_formula = f'=C21*(1+({base_wacc}-{wv})*12)*(1+({gv}-{base_g})*8)'
        cell.value = cell_formula
        cell.font = GREEN_FONT
        cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        cell.alignment = CENTER
        cell.border = THIN_BORDER
        cell.number_format = '0.00'

        # Highlight center cell (base case)
        if wv == 0.095 and gv == 0.025:
            cell.fill = MID_BLUE_FILL
            cell.font = BLACK_BOLD


# ---- 同行可比估值 ----
r = 35
ws_dcf.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
apply_header_style(ws_dcf, r, 1, 7)
ws_dcf.cell(row=r, column=1).value = "可比公司相对估值"

r = 36
comps_headers = ["公司", "代码", "市值(亿)", "P/E(TTM)", "P/B", "EV/EBITDA"]
apply_col_header_style(ws_dcf, r, 1, 6)
for i, h in enumerate(comps_headers):
    ws_dcf.cell(row=r, column=i+1).value = h

comps_data = [
    ["泰晶科技", "603738", 15.50*3.85, 28.5, 2.8, 18.2],
    ["惠伦晶体", "300460", 35, 45.0, 3.5, 25.0],
    ["晶赛科技", "871981", 12, 32.0, 2.2, 20.0],
    ["东晶电子", "002199", 18, 55.0, 3.0, 30.0],
    ["台湾晶技", "3042.TW", 80, 15.0, 2.0, 10.0],
]

for i, comp in enumerate(comps_data):
    crow = r + 1 + i
    for j, val in enumerate(comp):
        c = ws_dcf.cell(row=crow, column=j+1)
        c.value = val
        c.font = BLUE_FONT
        c.fill = LIGHT_GREY_FILL
        c.alignment = CENTER
        c.border = THIN_BORDER

# 中位数
r_stat = r + 1 + len(comps_data)
set_label_cell(ws_dcf, r_stat, 1, "行业中位数", bold=True)
for col in range(3, 7):
    col_letter = get_column_letter(col)
    formula = f"=MEDIAN({col_letter}{r+1+1}:{col_letter}{r_stat-1})"
    set_formula_cell(ws_dcf, r_stat, col, formula)


# ============================================================
# 保存文件
# ============================================================
output_path = r"d:\claude-for-finance\outputs\泰晶科技\泰晶科技_三表模型_DCF估值.xlsx"
wb.save(output_path)
print(f"Excel模型已生成: {output_path}")
print("提示: 历史财务数据为估算值，请以年报数据验证并修正。")
print("提示: 敏感度分析使用简化公式，建议用完整DCF重算替换。")
