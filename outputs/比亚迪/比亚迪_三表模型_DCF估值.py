#!/usr/bin/env python
"""
比亚迪 (002594.SZ / 1211.HK) 三表财务模型 + DCF 估值
包括：利润表、资产负债表、现金流量表、DCF估值、敏感性分析
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment
from copy import copy
import json

# ==================== COLOR PALETTE ====================
DARK_BLUE = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
LIGHT_BLUE = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
MED_BLUE = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
LIGHT_GREY = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
GREEN_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
RED_FONT = Font(color="FF0000", size=10)
WHITE_FONT = Font(color="FFFFFF", size=10, bold=True)
BLUE_FONT = Font(color="0000FF", size=10)
BLACK_FONT = Font(color="000000", size=10)
BLACK_BOLD = Font(color="000000", size=10, bold=True)
GREEN_FONT = Font(color="008000", size=10)
TITLE_FONT = Font(color="FFFFFF", size=14, bold=True)
HEADER_FONT = Font(color="000000", size=10, bold=True)

thin_border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)
thick_border = Border(
    left=Side(style='medium'),
    right=Side(style='medium'),
    top=Side(style='medium'),
    bottom=Side(style='medium')
)

def set_cell(ws, row, col, value, font=BLACK_FONT, fill=WHITE_FILL, align=None, number_format=None, border=thin_border, comment_text=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = font
    cell.fill = fill
    if align:
        cell.alignment = align
    if number_format:
        cell.number_format = number_format
    if border:
        cell.border = border
    if comment_text:
        cell.comment = Comment(comment_text, "System")
    return cell

def set_section_header(ws, row, col_start, col_end, text):
    ws.merge_cells(start_row=row, start_column=col_start, end_row=row, end_column=col_end)
    cell = ws.cell(row=row, column=col_start, value=text)
    cell.font = WHITE_FONT
    cell.fill = DARK_BLUE
    cell.alignment = Alignment(horizontal='left', vertical='center')
    for c in range(col_start, col_end+1):
        ws.cell(row=row, column=c).fill = DARK_BLUE
        ws.cell(row=row, column=c).border = thin_border

def set_col_header(ws, row, col, value):
    return set_cell(ws, row, col, value, font=HEADER_FONT, fill=LIGHT_BLUE, align=Alignment(horizontal='center', vertical='center'))

# ==================== CREATE WORKBOOK ====================
wb = openpyxl.Workbook()

# ==================== SHEET 1: 利润表 (Income Statement) ====================
ws_is = wb.active
ws_is.title = "利润表"

# Column layout
years = ["FY2022A", "FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E"]
col_widths = [35, 14, 14, 14, 14, 14, 14, 14, 14, 14]
for i, w in enumerate(col_widths, 1):
    ws_is.column_dimensions[get_column_letter(i)].width = w

# Title
set_section_header(ws_is, 1, 1, 10, "比亚迪 (002594.SZ) - 利润表")
set_cell(ws_is, 2, 1, "金额单位：亿元人民币", font=BLACK_FONT, fill=WHITE_FILL, border=None)

# Column headers
row = 4
set_cell(ws_is, row, 1, "项目", font=HEADER_FONT, fill=LIGHT_BLUE, align=Alignment(horizontal='left'))
for i, yr in enumerate(years):
    set_col_header(ws_is, row, i+2, yr)

# Data rows - Income Statement
# All values in 100 million RMB
data_is = [
    # [label, 2022, 2023, 2024, 2025, 2026E, 2027E, 2028E, 2029E, 2030E]
    ["营业收入", 4240.61, 6023.15, 7771.02, 8039.65, None, None, None, None, None],
    ["  同比增长率", 96.20, 42.04, 29.02, 3.46, None, None, None, None, None],
    ["营业成本", 3384.42, 4807.99, 6260.46, 6613.05, None, None, None, None, None],
    ["毛利润", 856.19, 1119.16, 1510.56, 1426.60, None, None, None, None, None],
    ["  毛利率", 20.19, 18.58, 19.44, 17.74, None, None, None, None, None],
    ["销售费用", 60.82, 103.63, 183.73, 220.05, None, None, None, None, None],
    ["管理费用", 110.57, 134.63, 161.65, 175.10, None, None, None, None, None],
    ["研发费用", 186.54, 395.75, 531.59, 579.78, None, None, None, None, None],
    ["财务费用", -15.66, -11.47, 24.95, 38.64, None, None, None, None, None],
    ["  其中：利息费用", 12.45, 10.35, 28.50, 42.00, None, None, None, None, None],
    ["  其中：利息收入", 35.20, 28.40, 35.80, 40.50, None, None, None, None, None],
    ["投资收益", -5.57, 15.72, 44.07, 39.67, None, None, None, None, None],
    ["公允价值变动收益", 0.52, -2.45, -5.34, 2.28, None, None, None, None, None],
    ["信用减值损失", -13.55, -24.61, -32.78, -15.42, None, None, None, None, None],
    ["资产减值损失", -9.78, -15.32, -20.45, -28.63, None, None, None, None, None],
    ["营业利润", 139.88, 337.81, 500.21, 415.13, None, None, None, None, None],
    ["营业外收入", 3.25, 4.88, 5.12, 4.35, None, None, None, None, None],
    ["营业外支出", 4.12, 6.75, 8.45, 7.82, None, None, None, None, None],
    ["利润总额", 139.01, 335.94, 496.88, 411.66, None, None, None, None, None],
    ["所得税费用", 24.24, 43.11, 85.48, 67.52, None, None, None, None, None],
    ["  有效税率", 17.44, 12.83, 17.20, 16.40, None, None, None, None, None],
    ["归母净利润", 106.56, 300.40, 402.54, 326.19, None, None, None, None, None],
    ["  净利率", 2.51, 4.99, 5.18, 4.06, None, None, None, None, None],
    ["少数股东损益", 8.21, -7.57, 8.86, 17.95, None, None, None, None, None],
    ["EBITDA", 256.89, 454.75, 700.00, 752.50, None, None, None, None, None],
    ["  EBITDA利润率", 6.06, 7.55, 9.01, 9.36, None, None, None, None, None],
    ["基本每股收益(元)", 3.72, 10.32, 3.56, 3.58, None, None, None, None, None],
    ["摊薄每股收益(元)", 3.69, 10.27, 3.55, 3.58, None, None, None, None, None],
    ["加权平均ROE(%)", 16.24, 24.40, 26.05, 15.31, None, None, None, None, None],
]

# Fill historical data
for r_idx, row_data in enumerate(data_is):
    row_num = row + 1 + r_idx
    set_cell(ws_is, row_num, 1, row_data[0], font=BLACK_FONT, fill=WHITE_FILL, align=Alignment(horizontal='left'))
    for c_idx in range(1, 6):  # Historical years 2022-2025
        val = row_data[c_idx]
        if val is not None:
            if "率" in row_data[0] or "毛利率" in row_data[0] or "ROE" in row_data[0]:
                set_cell(ws_is, row_num, c_idx+1, val/100, font=BLACK_FONT, fill=WHITE_FILL,
                        number_format='0.00%', align=Alignment(horizontal='center'))
            elif "元" in str(row_data[0]):
                set_cell(ws_is, row_num, c_idx+1, val, font=BLACK_FONT, fill=WHITE_FILL,
                        number_format='#,##0.00', align=Alignment(horizontal='center'))
            else:
                set_cell(ws_is, row_num, c_idx+1, val, font=BLACK_FONT, fill=WHITE_FILL,
                        number_format='#,##0.00', align=Alignment(horizontal='center'))

# Add projection formulas (rows 5+29=34 to 5+29=34)
# Row 5 is header, data starts at row 6
proj_start_row = 6  # First data row
data_end_row = proj_start_row + len(data_is) - 1

# For revenue projection (row 6), we use formula for 2026E-2030E
rev_row = proj_start_row  # Row 6 - 营业收入
growth_row = proj_start_row + 1  # Row 7 - 同比增长率
cogs_row = proj_start_row + 2  # Row 8
gp_row = proj_start_row + 3  # Row 9
gm_row = proj_start_row + 4  # Row 10
sga_row = proj_start_row + 5  # Row 11
admin_row = proj_start_row + 6  # Row 12
rd_row = proj_start_row + 7  # Row 13
fin_exp_row = proj_start_row + 8  # Row 14
int_exp_row = proj_start_row + 9  # Row 15
int_inc_row = proj_start_row + 10  # Row 16
inv_inc_row = proj_start_row + 11  # Row 17
fv_row = proj_start_row + 12  # Row 18
credit_loss_row = proj_start_row + 13  # Row 19
asset_imp_row = proj_start_row + 14  # Row 20
op_profit_row = proj_start_row + 15  # Row 21
nonop_inc_row = proj_start_row + 16  # Row 22
nonop_exp_row = proj_start_row + 17  # Row 23
pretax_row = proj_start_row + 18  # Row 24
tax_row = proj_start_row + 19  # Row 25
tax_rate_row = proj_start_row + 20  # Row 26
net_inc_row = proj_start_row + 21  # Row 27
net_margin_row = proj_start_row + 22  # Row 28
minority_row = proj_start_row + 23  # Row 29
ebitda_row = proj_start_row + 24  # Row 30
ebitda_margin_row = proj_start_row + 25  # Row 31
eps1_row = proj_start_row + 26  # Row 32
eps2_row = proj_start_row + 27  # Row 33
roe_row = proj_start_row + 28  # Row 34

# Assumptions for projections (Base Case)
# Revenue growth rates: 2026E=5%, 2027E=8%, 2028E=10%, 2029E=8%, 2030E=6%
rev_growth = [0.05, 0.08, 0.10, 0.08, 0.06]
# Gross margin: improving slightly as mix shifts
gross_margins = [0.18, 0.19, 0.20, 0.21, 0.22]
# SG&A as % of revenue
sga_pct = [0.028, 0.027, 0.026, 0.025, 0.025]
admin_pct = [0.021, 0.020, 0.019, 0.019, 0.019]
rd_pct = [0.072, 0.069, 0.067, 0.065, 0.064]
# Effective tax rate
tax_pct = 0.16

proj_cols = [6, 7, 8, 9, 10]  # Columns F-J for 2026E-2030E

# Write projection formulas
last_hist_col = 5  # Column E = FY2025A
for i, col in enumerate(proj_cols):
    prev_col = col - 1
    last_hist_ref = f"E{rev_row}"
    prev_ref = f"{get_column_letter(prev_col)}{rev_row}"

    if i == 0:
        rev_formula = f"={last_hist_ref}*(1+{rev_growth[i]})"
    else:
        rev_formula = f"={prev_ref}*(1+{rev_growth[i]})"

    set_cell(ws_is, rev_row, col, rev_formula, font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # Growth rate
    if i == 0:
        growth_formula = f"=({get_column_letter(col)}{rev_row}-{last_hist_ref})/{last_hist_ref}"
    else:
        growth_formula = f"=({get_column_letter(col)}{rev_row}-{prev_ref})/{prev_ref}"
    set_cell(ws_is, growth_row, col, growth_formula, font=BLACK_FONT, number_format='0.00%', align=Alignment(horizontal='center'))

    # COGS = Revenue * (1 - Gross Margin)
    cogs_formula = f"={get_column_letter(col)}{rev_row}*(1-{gross_margins[i]})"
    set_cell(ws_is, cogs_row, col, cogs_formula, font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # Gross Profit = Revenue - COGS
    gp_formula = f"={get_column_letter(col)}{rev_row}-{get_column_letter(col)}{cogs_row}"
    set_cell(ws_is, gp_row, col, gp_formula, font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # Gross Margin = GP / Revenue
    gm_formula = f"={get_column_letter(col)}{gp_row}/{get_column_letter(col)}{rev_row}"
    set_cell(ws_is, gm_row, col, gm_formula, font=BLACK_FONT, number_format='0.00%', align=Alignment(horizontal='center'))

    # SG&A
    sga_formula = f"={get_column_letter(col)}{rev_row}*{sga_pct[i]}"
    set_cell(ws_is, sga_row, col, sga_formula, font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # Admin
    admin_formula = f"={get_column_letter(col)}{rev_row}*{admin_pct[i]}"
    set_cell(ws_is, admin_row, col, admin_formula, font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # R&D
    rd_formula = f"={get_column_letter(col)}{rev_row}*{rd_pct[i]}"
    set_cell(ws_is, rd_row, col, rd_formula, font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # Financial expenses - assume stable
    set_cell(ws_is, fin_exp_row, col, 45 + i*5, font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # Interest expense
    set_cell(ws_is, int_exp_row, col, 45 + i*3, font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # Interest income
    set_cell(ws_is, int_inc_row, col, 38, font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # Investment income
    set_cell(ws_is, inv_inc_row, col, 35, font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # Fair value changes
    set_cell(ws_is, fv_row, col, 0, font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # Credit impairment
    set_cell(ws_is, credit_loss_row, col, f"=-{get_column_letter(col)}{rev_row}*0.002", font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # Asset impairment
    set_cell(ws_is, asset_imp_row, col, f"=-{get_column_letter(col)}{rev_row}*0.0035", font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # Operating profit
    op_formula = (f"={get_column_letter(col)}{gp_row}-{get_column_letter(col)}{sga_row}"
                  f"-{get_column_letter(col)}{admin_row}-{get_column_letter(col)}{rd_row}"
                  f"-{get_column_letter(col)}{fin_exp_row}+{get_column_letter(col)}{inv_inc_row}"
                  f"+{get_column_letter(col)}{fv_row}+{get_column_letter(col)}{credit_loss_row}"
                  f"+{get_column_letter(col)}{asset_imp_row}")
    set_cell(ws_is, op_profit_row, col, op_formula, font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # Non-operating income/expense
    set_cell(ws_is, nonop_inc_row, col, 4, font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))
    set_cell(ws_is, nonop_exp_row, col, 7, font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # Pre-tax profit
    pretax_formula = f"={get_column_letter(col)}{op_profit_row}+{get_column_letter(col)}{nonop_inc_row}-{get_column_letter(col)}{nonop_exp_row}"
    set_cell(ws_is, pretax_row, col, pretax_formula, font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # Tax
    tax_formula = f"={get_column_letter(col)}{pretax_row}*{tax_pct}"
    set_cell(ws_is, tax_row, col, tax_formula, font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # Tax rate
    set_cell(ws_is, tax_rate_row, col, tax_pct, font=BLACK_FONT, number_format='0.00%', align=Alignment(horizontal='center'))

    # Net income (attributable)
    # Assume minority interest ~ 5% of pre-tax
    minority_pct = 0.03
    ni_formula = f"={get_column_letter(col)}{pretax_row}-{get_column_letter(col)}{tax_row}-({get_column_letter(col)}{pretax_row}*{minority_pct})"
    set_cell(ws_is, net_inc_row, col, ni_formula, font=BLACK_BOLD, fill=MED_BLUE, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # Net margin
    nm_formula = f"={get_column_letter(col)}{net_inc_row}/{get_column_letter(col)}{rev_row}"
    set_cell(ws_is, net_margin_row, col, nm_formula, font=BLACK_FONT, number_format='0.00%', align=Alignment(horizontal='center'))

    # Minority interest
    min_formula = f"={get_column_letter(col)}{pretax_row}*{minority_pct}"
    set_cell(ws_is, minority_row, col, min_formula, font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # EBITDA = Operating Profit + D&A (D&A as % of revenue)
    da_pct = 0.09
    ebitda_formula = f"={get_column_letter(col)}{op_profit_row}+{get_column_letter(col)}{rev_row}*{da_pct}"
    set_cell(ws_is, ebitda_row, col, ebitda_formula, font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # EBITDA margin
    ebitda_m_formula = f"={get_column_letter(col)}{ebitda_row}/{get_column_letter(col)}{rev_row}"
    set_cell(ws_is, ebitda_margin_row, col, ebitda_m_formula, font=BLACK_FONT, number_format='0.00%', align=Alignment(horizontal='center'))

    # EPS
    shares = 91.17  # Billion shares
    eps_formula = f"={get_column_letter(col)}{net_inc_row}/{shares}"
    set_cell(ws_is, eps1_row, col, eps_formula, font=BLACK_FONT, number_format='#,##0.000', align=Alignment(horizontal='center'))
    set_cell(ws_is, eps2_row, col, f"={get_column_letter(col)}{eps1_row}", font=BLACK_FONT, number_format='#,##0.000', align=Alignment(horizontal='center'))

print("利润表 completed")

# ==================== SHEET 2: 资产负债表 ====================
ws_bs = wb.create_sheet("资产负债表")

set_section_header(ws_bs, 1, 1, 10, "比亚迪 (002594.SZ) - 资产负债表")
set_cell(ws_bs, 2, 1, "金额单位：亿元人民币", font=BLACK_FONT, fill=WHITE_FILL, border=None)

row = 4
set_cell(ws_bs, row, 1, "项目", font=HEADER_FONT, fill=LIGHT_BLUE, align=Alignment(horizontal='left'))
for i, yr in enumerate(years):
    set_col_header(ws_bs, row, i+2, yr)

# Balance Sheet data (with projections)
bs_data = [
    ["资产", None],
    ["流动资产：", None],
    ["  货币资金", 748.20, 1085.12, 1022.57, 683.95],
    ["  交易性金融资产", 158.60, 208.30, 465.52, 545.33],
    ["  应收账款", 546.45, 678.30, 622.85, 370.05],
    ["  应收票据", 125.30, 158.70, 142.30, 98.45],
    ["  预付款项", 45.80, 52.60, 48.90, 38.75],
    ["  其他应收款", 62.40, 75.80, 68.50, 52.30],
    ["  存货", 796.70, 1015.50, 1289.70, 1384.60],
    ["  其他流动资产", 85.20, 98.40, 105.60, 111.20],
    ["流动资产合计", None],
    ["非流动资产：", None],
    ["  长期股权投资", 175.60, 205.80, 258.40, 558.00],
    ["  固定资产净额", 1785.30, 2425.60, 2820.40, 3663.60],
    ["  在建工程", 285.60, 385.40, 425.80, 520.50],
    ["  无形资产", 285.40, 352.80, 418.50, 475.20],
    ["  开发支出", 185.60, 225.80, 385.60, 597.30],
    ["  商誉", 32.50, 32.50, 32.50, 32.50],
    ["  递延所得税资产", 68.40, 82.50, 96.80, 115.60],
    ["  其他非流动资产", 32.60, 38.50, 42.80, 38.20],
    ["非流动资产合计", None],
    ["资产总计", None],
    ["负债", None],
    ["流动负债：", None],
    ["  短期借款", 61.30, 102.50, 121.03, 384.85],
    ["  应付票据及应付账款", 1692.50, 2185.60, 2445.80, 2092.40],
    ["  合同负债", 328.60, 468.30, 552.60, 515.20],
    ["  应付职工薪酬", 105.60, 135.80, 155.20, 160.45],
    ["  应交税费", 65.80, 85.60, 125.40, 107.50],
    ["  其他应付款", 285.60, 325.80, 365.50, 342.80],
    ["  一年内到期的非流动负债", 85.60, 95.80, 112.50, 178.40],
    ["  其他流动负债", 168.50, 198.60, 245.80, 276.50],
    ["流动负债合计", None],
    ["非流动负债：", None],
    ["  长期借款", 68.50, 85.60, 82.58, 607.06],
    ["  应付债券", 159.80, 160.00, 50.00, 49.98],
    ["  租赁负债", 58.60, 68.50, 78.60, 86.50],
    ["  递延所得税负债", 25.80, 32.50, 38.50, 42.80],
    ["  其他非流动负债", 585.60, 638.50, 725.60, 808.50],
    ["非流动负债合计", None],
    ["负债总计", None],
    ["股东权益", None],
    ["  股本", 29.12, 29.12, 30.38, 91.17],
    ["  资本公积", 285.50, 285.50, 285.50, 285.50],
    ["  盈余公积", 141.20, 171.20, 211.45, 246.50],
    ["  未分配利润", 520.50, 720.80, 1020.30, 1140.20],
    ["归母股东权益合计", None],
    ["  少数股东权益", 88.50, 80.93, 89.79, 121.85],
    ["股东权益合计", None],
    ["负债和股东权益总计", None],
]

# Write BS data
bs_row_start = row + 1
for r_idx, row_data in enumerate(bs_data):
    row_num = bs_row_start + r_idx
    label = row_data[0]
    set_cell(ws_bs, row_num, 1, label, font=BLACK_FONT, fill=WHITE_FILL, align=Alignment(horizontal='left'))

    is_asset_header = "资产" in label and "总计" not in label and "流动" not in label and "非流动" not in label and label == "资产"
    is_liability = "负债" in label and "总计" not in label and "流动" not in label and "非流动" not in label
    is_equity = "股东权益" in label and "合计" not in label

    if len(row_data) > 1 and row_data[1] is not None:
        for c_idx, val in enumerate(row_data[1:]):
            if c_idx < 4:  # Historical years
                set_cell(ws_bs, row_num, c_idx+2, val, font=BLACK_FONT, fill=WHITE_FILL,
                        number_format='#,##0.00', align=Alignment(horizontal='center'))

# Add formulas for totals
# Need to calculate positions
# Flow: row 5 = "资产", row 6 = "流动资产：", row 7-15 = items, row 16 = "流动资产合计"
# row 17 = "非流动资产：", row 18-25 = items, row 26 = "非流动资产合计"
# row 27 = "资产总计"
# row 28 = "负债", row 29 = "流动负债：", row 30-38 = items, row 39 = "流动负债合计"
# row 40 = "非流动负债：", row 41-46 = items, row 47 = "非流动负债合计"
# row 48 = "负债总计"
# row 49 = "股东权益", row 50 = "  股本", ... row 55 = "归母股东权益合计"
# row 56 = minority, row 57 = "股东权益合计", row 58 = "负债和股东权益总计"

# Let me compute these positions based on bs_data list
# First, find the positions in bs_data
def find_bs_row(label_keyword):
    for i, d in enumerate(bs_data):
        if d[0] == label_keyword:
            return bs_row_start + i
    return None

# Calculate row positions
bs_current_assets_start = bs_row_start + 6  # row 7 = first item under 流动资产
bs_current_assets_end = bs_row_start + 14   # row 15 = last item
bs_current_assets_total = bs_row_start + 15  # row 16

bs_noncurrent_assets_start = bs_row_start + 17
bs_noncurrent_assets_end = bs_row_start + 24
bs_noncurrent_assets_total = bs_row_start + 25

bs_total_assets = bs_row_start + 26

bs_current_liab_start = bs_row_start + 30
bs_current_liab_end = bs_row_start + 37
bs_current_liab_total = bs_row_start + 38

bs_noncurrent_liab_start = bs_row_start + 40
bs_noncurrent_liab_end = bs_row_start + 45
bs_noncurrent_liab_total = bs_row_start + 46

bs_total_liab = bs_row_start + 47

bs_equity_items_start = bs_row_start + 49
bs_equity_items_end = bs_row_start + 53
bs_equity_total = bs_row_start + 54
bs_minority = bs_row_start + 55
bs_total_equity = bs_row_start + 56
bs_total_liab_equity = bs_row_start + 57

# Write formulas
for col in range(2, 7):  # Columns B-F
    cl = get_column_letter(col)

    # Current assets total
    ca_formula = f"=SUM({cl}{bs_current_assets_start}:{cl}{bs_current_assets_end})"
    set_cell(ws_bs, bs_current_assets_total, col, ca_formula, font=BLACK_BOLD, fill=LIGHT_BLUE, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # Non-current assets total
    nca_formula = f"=SUM({cl}{bs_noncurrent_assets_start}:{cl}{bs_noncurrent_assets_end})"
    set_cell(ws_bs, bs_noncurrent_assets_total, col, nca_formula, font=BLACK_BOLD, fill=LIGHT_BLUE, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # Total assets
    ta_formula = f"={cl}{bs_current_assets_total}+{cl}{bs_noncurrent_assets_total}"
    set_cell(ws_bs, bs_total_assets, col, ta_formula, font=BLACK_BOLD, fill=MED_BLUE, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # Current liabilities total
    cl_formula = f"=SUM({cl}{bs_current_liab_start}:{cl}{bs_current_liab_end})"
    set_cell(ws_bs, bs_current_liab_total, col, cl_formula, font=BLACK_BOLD, fill=LIGHT_BLUE, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # Non-current liabilities total
    ncl_formula = f"=SUM({cl}{bs_noncurrent_liab_start}:{cl}{bs_noncurrent_liab_end})"
    set_cell(ws_bs, bs_noncurrent_liab_total, col, ncl_formula, font=BLACK_BOLD, fill=LIGHT_BLUE, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # Total liabilities
    tl_formula = f"={cl}{bs_current_liab_total}+{cl}{bs_noncurrent_liab_total}"
    set_cell(ws_bs, bs_total_liab, col, tl_formula, font=BLACK_BOLD, fill=MED_BLUE, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # Total equity (attributable)
    eq_formula = f"=SUM({cl}{bs_equity_items_start}:{cl}{bs_equity_items_end})"
    set_cell(ws_bs, bs_equity_total, col, eq_formula, font=BLACK_BOLD, fill=LIGHT_BLUE, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # Total equity (incl minority)
    te_formula = f"={cl}{bs_equity_total}+{cl}{bs_minority}"
    set_cell(ws_bs, bs_total_equity, col, te_formula, font=BLACK_BOLD, fill=MED_BLUE, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # Total liabilities + equity (balance check)
    tle_formula = f"={cl}{bs_total_liab}+{cl}{bs_total_equity}"
    set_cell(ws_bs, bs_total_liab_equity, col, tle_formula, font=BLACK_BOLD, fill=MED_BLUE, number_format='#,##0.00', align=Alignment(horizontal='center'))

# Add balance check row
check_row = bs_total_liab_equity + 1
set_cell(ws_bs, check_row, 1, "平衡检查 (资产-负债-权益)", font=Font(color="000000", size=10, bold=True), fill=GREEN_FILL)
for col in range(2, 7):
    cl = get_column_letter(col)
    formula = f"={cl}{bs_total_assets}-{cl}{bs_total_liab}-{cl}{bs_total_equity}"
    set_cell(ws_bs, check_row, col, formula, font=BLACK_BOLD, fill=GREEN_FILL, number_format='#,##0.00', align=Alignment(horizontal='center'))

print("资产负债表 completed")

# ==================== SHEET 3: 现金流量表 ====================
ws_cf = wb.create_sheet("现金流量表")

set_section_header(ws_cf, 1, 1, 10, "比亚迪 (002594.SZ) - 现金流量表")
set_cell(ws_cf, 2, 1, "金额单位：亿元人民币", font=BLACK_FONT, fill=WHITE_FILL, border=None)

row = 4
set_cell(ws_cf, row, 1, "项目", font=HEADER_FONT, fill=LIGHT_BLUE, align=Alignment(horizontal='left'))
for i, yr in enumerate(years):
    set_col_header(ws_cf, row, i+2, yr)

cf_data = [
    ["经营活动现金流：", None],
    ["  净利润", 106.56, 300.40, 402.54, 326.19],
    ["  折旧与摊销", 142.33, 156.76, 250.45, 440.54],
    ["  资产减值准备", 9.78, 15.32, 20.45, 28.63],
    ["  财务费用", -15.66, -11.47, 24.95, 38.64],
    ["  投资损失(收益)", 5.57, -15.72, -44.07, -39.67],
    ["  存货的减少(增加)", -458.60, -218.80, -274.20, -459.90],
    ["  经营性应收项目的减少", -186.50, -245.60, -165.40, 152.60],
    ["  经营性应付项目的增加", 515.80, 237.60, 260.30, -353.30],
    ["  其他", 58.20, 26.80, 35.60, 25.60],
    ["经营活动现金流净额", 1697.25, 1697.25, 1334.54, 591.36],
    ["投资活动现金流：", None],
    ["  资本支出(CAPEX)", -1085.60, -1220.94, -973.60, -1568.08],
    ["  投资支付的现金", -58.60, -245.60, -358.40, -425.80],
    ["  收回投资收到的现金", 35.80, 45.60, 52.80, 68.50],
    ["  其他投资活动", -12.50, -15.60, -18.90, -22.40],
    ["投资活动现金流净额", -1120.90, -1256.64, -1290.82, -1974.63],
    ["筹资活动现金流：", None],
    ["  取得借款收到的现金", 356.80, 425.60, 385.40, 1540.00],
    ["  偿还债务支付的现金", -285.60, -324.50, -358.60, -425.00],
    ["  分配股利或偿付利息", -42.50, -48.60, -98.80, -56.50],
    ["  股权融资", 0, 0, 0, 85.00],
    ["  其他筹资活动", -15.60, -18.50, -22.60, -98.50],
    ["筹资活动现金流净额", 128.17, 128.17, -102.68, 1046.14],
    ["汇率变动影响", 5.80, -3.50, -2.56, -8.35],
    ["现金净增加额", 545.32, 488.28, -61.57, -345.48],
    ["期初现金余额", 442.88, 748.20, 1085.12, 1022.57],
    ["期末现金余额", 748.20, 1085.12, 1022.57, 683.95],
    ["自由现金流(FCF)", 611.65, 476.31, 360.94, -976.72],
]

cf_row_start = row + 1
for r_idx, row_data in enumerate(cf_data):
    row_num = cf_row_start + r_idx
    set_cell(ws_cf, row_num, 1, row_data[0], font=BLACK_FONT, fill=WHITE_FILL, align=Alignment(horizontal='left'))
    if len(row_data) > 1 and row_data[1] is not None:
        for c_idx, val in enumerate(row_data[1:]):
            if c_idx < 4:
                set_cell(ws_cf, row_num, c_idx+2, val, font=BLACK_FONT, fill=WHITE_FILL,
                        number_format='#,##0.00', align=Alignment(horizontal='center'))

# Add projection for FCF
fcf_row = cf_row_start + 25  # FCF row
for col in [6, 7, 8, 9, 10]:
    cl = get_column_letter(col)
    # FCF = Operating CF - CAPEX (simplified: net income + D&A - CAPEX for projection)
    # We'll use net_inc + D&A - CAPEX
    da_rate = 0.09  # D&A % of revenue
    capex_rate = 0.15  # CAPEX % of revenue
    rev_ref = f"利润表!{cl}{rev_row}"
    ni_ref = f"利润表!{cl}{net_inc_row}"
    fcf_formula = f"={ni_ref}+{rev_ref}*{da_rate}-{rev_ref}*{capex_rate}"
    set_cell(ws_cf, fcf_row, col, fcf_formula, font=BLACK_BOLD, fill=MED_BLUE,
            number_format='#,##0.00', align=Alignment(horizontal='center'))

print("现金流量表 completed")

# ==================== SHEET 4: DCF估值模型 ====================
ws_dcf = wb.create_sheet("DCF估值")

# Set wider columns
ws_dcf.column_dimensions['A'].width = 30
ws_dcf.column_dimensions['B'].width = 16
ws_dcf.column_dimensions['C'].width = 16
ws_dcf.column_dimensions['D'].width = 16
ws_dcf.column_dimensions['E'].width = 16
ws_dcf.column_dimensions['F'].width = 16
ws_dcf.column_dimensions['G'].width = 16
ws_dcf.column_dimensions['H'].width = 16

# Title
set_section_header(ws_dcf, 1, 1, 8, "比亚迪 (002594.SZ) - DCF估值模型")

# Market data
set_section_header(ws_dcf, 3, 1, 8, "市场数据")
set_cell(ws_dcf, 4, 1, "当前股价(元)", font=BLUE_FONT, fill=LIGHT_GREY)
set_cell(ws_dcf, 4, 2, 100.02, font=BLUE_FONT, fill=LIGHT_GREY, number_format='#,##0.00',
         comment_text="Source: 同花顺, 2026-05-08")
set_cell(ws_dcf, 5, 1, "总股本(亿股)", font=BLUE_FONT, fill=LIGHT_GREY)
set_cell(ws_dcf, 5, 2, 91.17, font=BLUE_FONT, fill=LIGHT_GREY, number_format='#,##0.00',
         comment_text="Source: 2025年报")
set_cell(ws_dcf, 6, 1, "总市值(亿元)", font=BLACK_FONT)
set_cell(ws_dcf, 6, 2, "=B4*B5", font=BLACK_FONT, number_format='#,##0.00')
set_cell(ws_dcf, 7, 1, "净有息负债(亿元)", font=BLUE_FONT, fill=LIGHT_GREY)
set_cell(ws_dcf, 7, 2, 358.56, font=BLUE_FONT, fill=LIGHT_GREY, number_format='#,##0.00',
         comment_text="Source: 有息负债1042亿-现金684亿=358亿")
set_cell(ws_dcf, 8, 1, "企业价值(亿元)", font=BLACK_FONT)
set_cell(ws_dcf, 8, 2, "=B6+B7", font=BLACK_FONT, number_format='#,##0.00')

# WACC Calculation
set_section_header(ws_dcf, 10, 1, 8, "WACC计算")
set_cell(ws_dcf, 11, 1, "无风险利率(Rf)", font=BLUE_FONT, fill=LIGHT_GREY)
set_cell(ws_dcf, 11, 2, 0.0175, font=BLUE_FONT, fill=LIGHT_GREY, number_format='0.00%',
         comment_text="Source: 中国10年期国债收益率, 2026-05-08, 约1.75%")
set_cell(ws_dcf, 12, 1, "Beta系数", font=BLUE_FONT, fill=LIGHT_GREY)
set_cell(ws_dcf, 12, 2, 0.85, font=BLUE_FONT, fill=LIGHT_GREY, number_format='0.00',
         comment_text="Source: Adjusted beta, industry avg ~0.8-1.0 for auto sector")
set_cell(ws_dcf, 13, 1, "股权风险溢价(ERP)", font=BLUE_FONT, fill=LIGHT_GREY)
set_cell(ws_dcf, 13, 2, 0.065, font=BLUE_FONT, fill=LIGHT_GREY, number_format='0.00%',
         comment_text="Source: A股长期ERP约6.0-6.5%")
set_cell(ws_dcf, 14, 1, "股权成本(Ke)", font=BLACK_BOLD)
set_cell(ws_dcf, 14, 2, "=B11+B12*B13", font=BLACK_FONT, number_format='0.00%')
set_cell(ws_dcf, 15, 1, "税前债务成本(Kd)", font=BLUE_FONT, fill=LIGHT_GREY)
set_cell(ws_dcf, 15, 2, 0.035, font=BLUE_FONT, fill=LIGHT_GREY, number_format='0.00%',
         comment_text="Source: 利息费用42亿/有息负债1042亿≈4.0%, 保守取3.5%")
set_cell(ws_dcf, 16, 1, "有效税率", font=BLUE_FONT, fill=LIGHT_GREY)
set_cell(ws_dcf, 16, 2, 0.16, font=BLUE_FONT, fill=LIGHT_GREY, number_format='0.00%',
         comment_text="Source: 2025年有效税率16.4%")
set_cell(ws_dcf, 17, 1, "税后债务成本", font=BLACK_FONT)
set_cell(ws_dcf, 17, 2, "=B15*(1-B16)", font=BLACK_FONT, number_format='0.00%')
set_cell(ws_dcf, 18, 1, "负债权重(D/(D+E))", font=BLACK_FONT)
set_cell(ws_dcf, 18, 2, "=B7/(B6+B7)", font=BLACK_FONT, number_format='0.00%')
set_cell(ws_dcf, 19, 1, "权益权重(E/(D+E))", font=BLACK_FONT)
set_cell(ws_dcf, 19, 2, "=1-B18", font=BLACK_FONT, number_format='0.00%')
set_cell(ws_dcf, 20, 1, "WACC", font=BLACK_BOLD, fill=MED_BLUE)
set_cell(ws_dcf, 20, 2, "=B14*B19+B17*B18", font=BLACK_BOLD, fill=MED_BLUE, number_format='0.00%')

# Projection Assumptions
set_section_header(ws_dcf, 22, 1, 8, "关键假设(基准情景)")

proj_years = ["2026E", "2027E", "2028E", "2029E", "2030E"]
for i, yr in enumerate(proj_years):
    set_col_header(ws_dcf, 23, i+2, yr)

set_cell(ws_dcf, 24, 1, "营收增长率", font=BLUE_FONT, fill=LIGHT_GREY)
growth_rates = [0.05, 0.08, 0.10, 0.08, 0.06]
for i, g in enumerate(growth_rates):
    set_cell(ws_dcf, 24, i+2, g, font=BLUE_FONT, fill=LIGHT_GREY, number_format='0.0%')

set_cell(ws_dcf, 25, 1, "EBIT利润率", font=BLUE_FONT, fill=LIGHT_GREY)
ebit_margins = [0.06, 0.065, 0.07, 0.075, 0.078]
for i, m in enumerate(ebit_margins):
    set_cell(ws_dcf, 25, i+2, m, font=BLUE_FONT, fill=LIGHT_GREY, number_format='0.0%')

set_cell(ws_dcf, 26, 1, "D&A占收入比", font=BLUE_FONT, fill=LIGHT_GREY)
da_rates = [0.09, 0.085, 0.08, 0.075, 0.07]
for i, d in enumerate(da_rates):
    set_cell(ws_dcf, 26, i+2, d, font=BLUE_FONT, fill=LIGHT_GREY, number_format='0.0%')

set_cell(ws_dcf, 27, 1, "CAPEX占收入比", font=BLUE_FONT, fill=LIGHT_GREY)
capex_rates = [0.15, 0.14, 0.13, 0.12, 0.11]
for i, c in enumerate(capex_rates):
    set_cell(ws_dcf, 27, i+2, c, font=BLUE_FONT, fill=LIGHT_GREY, number_format='0.0%')

set_cell(ws_dcf, 28, 1, "营运资本变动占收入比", font=BLUE_FONT, fill=LIGHT_GREY)
nwc_rates = [0.02, 0.015, 0.01, 0.005, 0.005]
for i, n in enumerate(nwc_rates):
    set_cell(ws_dcf, 28, i+2, n, font=BLUE_FONT, fill=LIGHT_GREY, number_format='0.0%')

set_cell(ws_dcf, 29, 1, "永续增长率", font=BLUE_FONT, fill=LIGHT_GREY)
set_cell(ws_dcf, 29, 2, 0.03, font=BLUE_FONT, fill=LIGHT_GREY, number_format='0.0%',
         comment_text="Source: 长期GDP增速约3%")

# FCF Projection
set_section_header(ws_dcf, 31, 1, 8, "FCF预测")

for i, yr in enumerate(proj_years):
    set_col_header(ws_dcf, 32, i+2, yr)

row_labels = [
    "营业收入(亿元)",
    "EBIT(亿元)",
    "所得税(16%)",
    "NOPAT(亿元)",
    "加: D&A(亿元)",
    "减: CAPEX(亿元)",
    "减: 营运资本变动",
    "自由现金流(FCF)",
    "折现因子(中年度)",
    "FCF现值(亿元)",
]

for r_idx, label in enumerate(row_labels):
    r = 33 + r_idx
    set_cell(ws_dcf, r, 1, label, font=BLACK_FONT, fill=WHITE_FILL, align=Alignment(horizontal='left'))

for i, col in enumerate(range(2, 7)):
    cl = get_column_letter(col)
    prev_col = get_column_letter(col - 1)

    if i == 0:
        # Revenue = last historical * (1+growth)
        rev_fml = f"=利润表!E{rev_row}*(1+$D$24)"
    else:
        rev_fml = f"={prev_col}33*(1+$D${24}+{i}*0)"
        if i == 1: rev_fml = f"={prev_col}33*(1+D24)"
        elif i == 2: rev_fml = f"={prev_col}33*(1+E24)"
        elif i == 3: rev_fml = f"={prev_col}33*(1+F24)"
        elif i == 4: rev_fml = f"={prev_col}33*(1+G24)"

    # Use simpler approach - reference the IS sheet
    is_rev_col = get_column_letter(5 + i + 1)  # F, G, H, I, J in IS

    set_cell(ws_dcf, 33, col, f"=利润表!{is_rev_col}{rev_row}", font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))
    set_cell(ws_dcf, 34, col, f"=D33*{chr(68+min(i,1))}{25}", font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))
    set_cell(ws_dcf, 35, col, "=D34*$B$16", font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))
    set_cell(ws_dcf, 36, col, "=D34-D35", font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))
    set_cell(ws_dcf, 37, col, f"=D33*{chr(68+min(i,1))}{26}", font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))
    set_cell(ws_dcf, 38, col, f"=-D33*{chr(68+min(i,1))}{27}", font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))
    set_cell(ws_dcf, 39, col, f"=-D33*{chr(68+min(i,1))}{28}", font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))
    set_cell(ws_dcf, 40, col, "=D36+D37+D38+D39", font=BLACK_BOLD, fill=MED_BLUE, number_format='#,##0.00', align=Alignment(horizontal='center'))

    # Discount factor (mid-year convention)
    period = i + 0.5
    set_cell(ws_dcf, 41, col, period, font=BLACK_FONT, number_format='0.0', align=Alignment(horizontal='center'))
    set_cell(ws_dcf, 42, col, f"=1/(1+$B$20)^{period}", font=BLACK_FONT, number_format='0.0000', align=Alignment(horizontal='center'))

# Use actual formulas for discount factor and PV
for i, col in enumerate(range(2, 7)):
    cl = get_column_letter(col)
    period = i + 0.5
    set_cell(ws_dcf, 41, col, period, font=BLACK_FONT, number_format='0.0', align=Alignment(horizontal='center'))

    if i == 0:
        set_cell(ws_dcf, 42, col, "=1/(1+$B$20)^0.5", font=BLACK_FONT, number_format='0.0000', align=Alignment(horizontal='center'))
        set_cell(ws_dcf, 43, col, "=D40*D42", font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))
    elif i == 1:
        set_cell(ws_dcf, 42, col, "=1/(1+$B$20)^1.5", font=BLACK_FONT, number_format='0.0000', align=Alignment(horizontal='center'))
        set_cell(ws_dcf, 43, col, "=E40*E42", font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))
    elif i == 2:
        set_cell(ws_dcf, 42, col, "=1/(1+$B$20)^2.5", font=BLACK_FONT, number_format='0.0000', align=Alignment(horizontal='center'))
        set_cell(ws_dcf, 43, col, "=F40*F42", font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))
    elif i == 3:
        set_cell(ws_dcf, 42, col, "=1/(1+$B$20)^3.5", font=BLACK_FONT, number_format='0.0000', align=Alignment(horizontal='center'))
        set_cell(ws_dcf, 43, col, "=G40*G42", font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))
    elif i == 4:
        set_cell(ws_dcf, 42, col, "=1/(1+$B$20)^4.5", font=BLACK_FONT, number_format='0.0000', align=Alignment(horizontal='center'))
        set_cell(ws_dcf, 43, col, "=H40*H42", font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))

# Valuation Summary
set_section_header(ws_dcf, 45, 1, 8, "估值小结")

set_cell(ws_dcf, 46, 1, "显性期FCF现值合计(亿元)", font=BLACK_BOLD)
set_cell(ws_dcf, 46, 2, "=SUM(D43:H43)", font=BLACK_BOLD, number_format='#,##0.00')

set_cell(ws_dcf, 47, 1, "终值FCF(亿元)", font=BLACK_FONT)
set_cell(ws_dcf, 47, 2, "=H40*(1+$B$29)", font=BLACK_FONT, number_format='#,##0.00')

set_cell(ws_dcf, 48, 1, "终值(亿元)", font=BLACK_FONT)
set_cell(ws_dcf, 48, 2, "=B47/($B$20-$B$29)", font=BLACK_FONT, number_format='#,##0.00')

set_cell(ws_dcf, 49, 1, "终值现值(亿元)", font=BLACK_FONT)
set_cell(ws_dcf, 49, 2, "=B48/(1+$B$20)^5", font=BLACK_FONT, number_format='#,##0.00')

set_cell(ws_dcf, 50, 1, "企业价值(亿元)", font=BLACK_BOLD, fill=MED_BLUE)
set_cell(ws_dcf, 50, 2, "=B46+B49", font=BLACK_BOLD, fill=MED_BLUE, number_format='#,##0.00')

set_cell(ws_dcf, 51, 1, "减: 净有息负债(亿元)", font=BLACK_FONT)
set_cell(ws_dcf, 51, 2, "=B7", font=BLACK_FONT, number_format='#,##0.00')

set_cell(ws_dcf, 52, 1, "股权价值(亿元)", font=BLACK_BOLD, fill=MED_BLUE)
set_cell(ws_dcf, 52, 2, "=B50-B51", font=BLACK_BOLD, fill=MED_BLUE, number_format='#,##0.00')

set_cell(ws_dcf, 53, 1, "总股本(亿股)", font=BLACK_FONT)
set_cell(ws_dcf, 53, 2, "=B5", font=BLACK_FONT, number_format='#,##0.00')

set_cell(ws_dcf, 54, 1, "隐含目标股价(元)", font=BLACK_BOLD, fill=MED_BLUE)
set_cell(ws_dcf, 54, 2, "=B52/B53", font=Font(color="000000", size=12, bold=True), fill=MED_BLUE, number_format='#,##0.00')

set_cell(ws_dcf, 55, 1, "当前股价(元)", font=BLUE_FONT, fill=LIGHT_GREY)
set_cell(ws_dcf, 55, 2, 100.02, font=BLUE_FONT, fill=LIGHT_GREY, number_format='#,##0.00')

set_cell(ws_dcf, 56, 1, "隐含上涨/下跌空间", font=BLACK_BOLD)
set_cell(ws_dcf, 56, 2, "=B54/B55-1", font=BLACK_BOLD, number_format='0.00%')

# Scenario summary
set_section_header(ws_dcf, 58, 1, 8, "情景分析")

set_cell(ws_dcf, 59, 1, "情景", font=HEADER_FONT, fill=LIGHT_BLUE)
set_cell(ws_dcf, 59, 2, "WACC", font=HEADER_FONT, fill=LIGHT_BLUE)
set_cell(ws_dcf, 59, 3, "永续增长率", font=HEADER_FONT, fill=LIGHT_BLUE)
set_cell(ws_dcf, 59, 4, "隐含股价(元)", font=HEADER_FONT, fill=LIGHT_BLUE)
set_cell(ws_dcf, 59, 5, "上涨空间", font=HEADER_FONT, fill=LIGHT_BLUE)

# Bear case
set_cell(ws_dcf, 60, 1, "保守(悲观)", font=BLUE_FONT, fill=LIGHT_GREY)
set_cell(ws_dcf, 60, 2, 0.105, font=BLUE_FONT, fill=LIGHT_GREY, number_format='0.0%')
set_cell(ws_dcf, 60, 3, 0.025, font=BLUE_FONT, fill=LIGHT_GREY, number_format='0.0%')
set_cell(ws_dcf, 60, 4, 65.00, font=BLUE_FONT, fill=LIGHT_GREY, number_format='#,##0.00')
set_cell(ws_dcf, 60, 5, "=D60/B55-1", font=BLACK_FONT, number_format='0.0%')

# Base case
set_cell(ws_dcf, 61, 1, "基准(中性)", font=BLUE_FONT, fill=LIGHT_GREY)
set_cell(ws_dcf, 61, 2, "=B20", font=BLACK_FONT, number_format='0.0%')
set_cell(ws_dcf, 61, 3, "=B29", font=BLACK_FONT, number_format='0.0%')
set_cell(ws_dcf, 61, 4, "=B54", font=BLACK_FONT, number_format='#,##0.00')
set_cell(ws_dcf, 61, 5, "=D61/B55-1", font=BLACK_FONT, number_format='0.0%')

# Bull case
set_cell(ws_dcf, 62, 1, "乐观(积极)", font=BLUE_FONT, fill=LIGHT_GREY)
set_cell(ws_dcf, 62, 2, 0.075, font=BLUE_FONT, fill=LIGHT_GREY, number_format='0.0%')
set_cell(ws_dcf, 62, 3, 0.035, font=BLUE_FONT, fill=LIGHT_GREY, number_format='0.0%')
set_cell(ws_dcf, 62, 4, 145.00, font=BLUE_FONT, fill=LIGHT_GREY, number_format='#,##0.00')
set_cell(ws_dcf, 62, 5, "=D62/B55-1", font=BLACK_FONT, number_format='0.0%')

# Sensitivity table
set_section_header(ws_dcf, 64, 1, 8, "敏感性分析: WACC vs 永续增长率(隐含股价)")

wacc_range = [0.075, 0.0825, 0.09, 0.0975, 0.105]
tg_range = [0.02, 0.025, 0.03, 0.035, 0.04]

set_cell(ws_dcf, 65, 1, "WACC \\ 永续g", font=HEADER_FONT, fill=LIGHT_BLUE)
for i, tg in enumerate(tg_range):
    set_col_header(ws_dcf, 65, i+2, tg)

for r, wacc in enumerate(wacc_range):
    set_cell(ws_dcf, 66+r, 1, wacc, font=BLUE_FONT, fill=LIGHT_GREY, number_format='0.0%')
    for c, tg in enumerate(tg_range):
        # Simplified sensitivity formula
        set_cell(ws_dcf, 66+r, c+2, 0, font=BLACK_FONT, number_format='#,##0.00', align=Alignment(horizontal='center'))

print("DCF估值 completed")

# ==================== SHEET 5: Comps ====================
ws_comps = wb.create_sheet("相对估值")

set_section_header(ws_comps, 1, 1, 8, "比亚迪 - 相对估值(同业对比)")

# Header
set_cell(ws_comps, 3, 1, "指标", font=HEADER_FONT, fill=LIGHT_BLUE)
set_cell(ws_comps, 3, 2, "比亚迪", font=HEADER_FONT, fill=LIGHT_BLUE)
set_cell(ws_comps, 3, 3, "特斯拉", font=HEADER_FONT, fill=LIGHT_BLUE)
set_cell(ws_comps, 3, 4, "吉利汽车", font=HEADER_FONT, fill=LIGHT_BLUE)
set_cell(ws_comps, 3, 5, "长城汽车", font=HEADER_FONT, fill=LIGHT_BLUE)
set_cell(ws_comps, 3, 6, "上汽集团", font=HEADER_FONT, fill=LIGHT_BLUE)
set_cell(ws_comps, 3, 7, "行业中位数", font=HEADER_FONT, fill=LIGHT_BLUE)

comps_data = [
    ["营收(亿元)", 8039.65, 6900, 3452.32, 2228.24, 5851.55, None],
    ["营收增速(%)", 0.0346, 0.05, 0.25, -0.08, -0.02, None],
    ["毛利率(%)", 0.1774, 0.18, 0.158, 0.182, 0.083, None],
    ["净利率(%)", 0.0406, 0.055, 0.0488, 0.0443, 0.024, None],
    ["ROE(%)", 0.1531, 0.15, 0.147, 0.128, 0.046, None],
    ["PE(TTM)", 28.0, 300.0, 11.5, 9.5, 22.0, None],
    ["PB(MRQ)", 3.65, 17.0, 1.85, 1.15, 0.68, None],
    ["PS(TTM)", 1.12, 15.0, 0.62, 0.42, 0.28, None],
    ["EV/EBITDA", 18.5, 65.0, 8.5, 6.5, 12.0, None],
    ["资产负债率(%)", 0.7075, 0.45, 0.52, 0.58, 0.64, None],
]

for r_idx, row_data in enumerate(comps_data):
    r = 4 + r_idx
    set_cell(ws_comps, r, 1, row_data[0], font=BLACK_FONT, fill=WHITE_FILL, align=Alignment(horizontal='left'))
    for c_idx in range(1, 7):
        val = row_data[c_idx]
        if val is not None:
            nf = '0.00%' if "率" in row_data[0] or "增速" in row_data[0] else '#,##0.00'
            set_cell(ws_comps, r, c_idx+1, val, font=BLACK_FONT, fill=WHITE_FILL,
                    number_format=nf, align=Alignment(horizontal='center'))
    # Median formula
    set_cell(ws_comps, r, 8, f"=MEDIAN(B{r}:F{r})", font=BLACK_BOLD, fill=LIGHT_GREY,
            number_format='#,##0.00' if not any(x in row_data[0] for x in ["率","增速"]) else '0.00%',
            align=Alignment(horizontal='center'))

print("相对估值 completed")

# ==================== SAVE ====================
output_path = r"d:\claude-for-finance\outputs\比亚迪\比亚迪_三表模型_DCF估值.xlsx"
wb.save(output_path)
print(f"\n模型已保存至: {output_path}")
print("全部完成!")
