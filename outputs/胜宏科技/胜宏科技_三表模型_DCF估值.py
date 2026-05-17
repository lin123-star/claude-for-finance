"""
胜宏科技(300476.SZ) 三表模型 + DCF估值
Excel 自动生成脚本
生成日期: 2026-05-10
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, numbers
from openpyxl.utils import get_column_letter
from copy import copy

wb = openpyxl.Workbook()

# ========== 样式定义 ==========
header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
header_font = Font(name="微软雅黑", bold=True, color="FFFFFF", size=11)
sub_header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
sub_header_font = Font(name="微软雅黑", bold=True, color="000000", size=10)
input_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
input_font = Font(name="微软雅黑", color="0000FF", size=10)
formula_font = Font(name="微软雅黑", color="000000", size=10)
output_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
output_font = Font(name="微软雅黑", bold=True, color="000000", size=10)
green_font = Font(name="微软雅黑", color="008000", size=10)

thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)
thick_border = Border(
    left=Side(style='medium'), right=Side(style='medium'),
    top=Side(style='medium'), bottom=Side(style='medium')
)

def set_cell(ws, row, col, value, font=formula_font, fill=None, number_format=None, border=thin_border, alignment=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = font
    if fill:
        cell.fill = fill
    if number_format:
        cell.number_format = number_format
    if border:
        cell.border = border
    if alignment:
        cell.alignment = alignment
    return cell

def set_header_row(ws, row, start_col, end_col, values, font=sub_header_font, fill=sub_header_fill):
    for i, v in enumerate(values):
        set_cell(ws, row, start_col + i, v, font=font, fill=fill)

def merge_and_set(ws, row, start_col, end_col, value, font=header_font, fill=header_fill):
    ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
    cell = set_cell(ws, row, start_col, value, font=font, fill=fill)
    cell.alignment = Alignment(horizontal='center', vertical='center')

# ====================================================================
# Sheet 1: 三表模型 (Income Statement, Balance Sheet, Cash Flow)
# ====================================================================
ws1 = wb.active
ws1.title = "三表模型"
ws1.sheet_properties.tabColor = "1F4E79"

# 列定义: A=项目名称, B=2021A, C=2022A, D=2023A, E=2024A, F=2025A, G=2026E, H=2027E, I=2028E, J=2029E, K=2030E
cols = ["B","C","D","E","F","G","H","I","J","K"]
year_labels = ["2021A","2022A","2023A","2024A","2025A","2026E","2027E","2028E","2029E","2030E"]

# ============ 标题行 ============
merge_and_set(ws1, 1, 1, 11, "胜宏科技(300476.SZ) 三表财务模型", Font(name="微软雅黑", bold=True, color="FFFFFF", size=14))
ws1.row_dimensions[1].height = 30

# ============ 收入利润表 ============
merge_and_set(ws1, 3, 1, 11, "一、利润表（单位：亿元人民币）")
ws1.row_dimensions[3].height = 25

# 列头
row = 4
set_cell(ws1, row, 1, "项目", font=sub_header_font, fill=sub_header_fill)
for i, yr in enumerate(year_labels):
    set_cell(ws1, row, 2+i, yr, font=sub_header_font, fill=sub_header_fill)

# 数据行 - 历史数据 + 公式
# Row mapping for IS:
r = 5  # 营业收入
set_cell(ws1, r, 1, "营业收入", font=Font(name="微软雅黑", bold=True, size=10))
is_data = {
    5: ["营业收入", [74.32, 78.85, 79.31, 107.31, 192.92]],  # 历史
    6: ["  同比增长率", None],
    7: ["营业成本", None],
    8: ["  毛利率", None],
    9: ["营业税金及附加", None],
    10: ["销售费用", None],
    11: ["管理费用", None],
    12: ["研发费用", None],
    13: ["财务费用", None],
    14: ["营业利润", None],
    15: ["利润总额", None],
    16: ["所得税费用", None],
    17: ["归母净利润", None],
    18: ["  净利率", None],
}

# 营业收入数据
rev_hist = [74.32, 78.85, 79.31, 107.31, 192.92]
for i, v in enumerate(rev_hist):
    set_cell(ws1, r, 2+i, v, font=input_font, fill=input_fill, number_format='#,##0.00')

# 增长率公式
for i in range(4):
    set_cell(ws1, 6, 2+i, f"=IFERROR({get_column_letter(2+i+1)}5/{get_column_letter(2+i)}5-1,\"\")", number_format='0.0%')
# 预测增长率手动输入
growth_rates_e = [0.70, 0.50, 0.40, 0.30, 0.20]  # 2026E-2030E
for i, g in enumerate(growth_rates_e):
    set_cell(ws1, 6, 7+i, g, font=input_font, fill=input_fill, number_format='0.0%')
# 预测营收公式
for i in range(5):
    col_prev = get_column_letter(6+i)  # 前一年列
    col_cur = get_column_letter(7+i)   # 当前列
    set_cell(ws1, 5, 7+i, f"={col_prev}5*(1+{col_cur}6)", number_format='#,##0.00')

# 营业成本 (基于毛利率假设)
gm_hist = [0.2037, 0.1815, 0.2070, 0.2874, 0.3338]
for i, v in enumerate(gm_hist):
    set_cell(ws1, 8, 2+i, f"=1-{v}", font=input_font, fill=input_fill, number_format='0.00%')
    set_cell(ws1, 7, 2+i, f"={get_column_letter(2+i)}5*(1-{get_column_letter(2+i)}8)", number_format='#,##0.00')

gm_e = [0.345, 0.355, 0.360, 0.365, 0.370]
for i, g in enumerate(gm_e):
    set_cell(ws1, 8, 7+i, f"=1-{g}", font=input_font, fill=input_fill, number_format='0.0%')
    col = get_column_letter(7+i)
    set_cell(ws1, 7, 7+i, f"={col}5*{col}8", number_format='#,##0.00')

# 毛利率
for i in range(10):
    col = get_column_letter(2+i)
    set_cell(ws1, 9, 2+i, f"=IFERROR(({col}5-{col}7)/{col}5,\"\")", number_format='0.0%')

# 期间费用 (基于营收比例假设)
# 简化模型：用营业利润倒推
# 直接使用归母净利润数据 + 历史
ni_hist = [6.70, 7.91, 6.71, 11.54, 43.12]
for i, v in enumerate(ni_hist):
    set_cell(ws1, 17, 2+i, v, font=input_font, fill=input_fill, number_format='#,##0.00')

# 费用合计 (倒算)
for i in range(5):
    col = get_column_letter(2+i)
    set_cell(ws1, 14, 2+i, f"={col}5-{col}7", number_format='#,##0.00')  # 营业利润 = 营收 - 成本(简化)

# 预测净利润 (基于净利率假设)
ni_margin_e = [0.233, 0.250, 0.255, 0.258, 0.260]
for i, m in enumerate(ni_margin_e):
    col = get_column_letter(7+i)
    set_cell(ws1, 18, 7+i, m, font=input_font, fill=input_fill, number_format='0.0%')
    set_cell(ws1, 17, 7+i, f"={col}5*{col}18", number_format='#,##0.00')

# 净利率
for i in range(10):
    col = get_column_letter(2+i)
    set_cell(ws1, 18, 2+i, f"=IFERROR({col}17/{col}5,\"\")", number_format='0.0%')

# ============ 资产负债表 ============
merge_and_set(ws1, 20, 1, 11, "二、资产负债表（单位：亿元人民币）")

# 资产
set_cell(ws1, 21, 1, "资产", font=Font(name="微软雅黑", bold=True, size=10))
# 历史数据行
bs_start = 21
# 简化版资产负债表关键项
# 货币资金
cash_hist = [15.62, 12.38, 14.85, 16.62, 32.80]
set_cell(ws1, 22, 1, "  货币资金")
for i, v in enumerate(cash_hist):
    set_cell(ws1, 22, 2+i, v, font=input_font, fill=input_fill, number_format='#,##0.00')
for i in range(5):
    col = get_column_letter(7+i)
    # 简化：现金保持一定比例
    set_cell(ws1, 22, 7+i, f"={col}5*0.17", number_format='#,##0.00')

# 应收账款
ar_hist = [18.35, 20.12, 21.56, 38.87, 58.71]
set_cell(ws1, 23, 1, "  应收账款")
for i, v in enumerate(ar_hist):
    set_cell(ws1, 23, 2+i, v, font=input_font, fill=input_fill, number_format='#,##0.00')
for i in range(5):
    col = get_column_letter(7+i)
    set_cell(ws1, 23, 7+i, f"={col}5*0.30", number_format='#,##0.00')

# 存货
inv_hist = [8.92, 10.45, 12.38, 20.45, 31.62]
set_cell(ws1, 24, 1, "  存货")
for i, v in enumerate(inv_hist):
    set_cell(ws1, 24, 2+i, v, font=input_font, fill=input_fill, number_format='#,##0.00')
for i in range(5):
    col = get_column_letter(7+i)
    set_cell(ws1, 24, 7+i, f"={col}5*0.16", number_format='#,##0.00')

# 其他流动资产
set_cell(ws1, 25, 1, "  其他流动资产")
for i in range(10):
    col = get_column_letter(2+i)
    set_cell(ws1, 25, 2+i, f"={col}5*0.05", number_format='#,##0.00')

# 流动资产合计
merge_and_set(ws1, 26, 1, 1, "  流动资产合计", font=output_font, fill=output_fill)
for i in range(10):
    col = get_column_letter(2+i)
    set_cell(ws1, 26, 2+i, f"=SUM(B22:B25)", number_format='#,##0.00')

# 固定资产
fa_hist = [35.62, 40.18, 45.32, 71.72, 91.64]
set_cell(ws1, 28, 1, "  固定资产")
for i, v in enumerate(fa_hist):
    set_cell(ws1, 28, 2+i, v, font=input_font, fill=input_fill, number_format='#,##0.00')
for i in range(5):
    col = get_column_letter(7+i)
    set_cell(ws1, 28, 7+i, f"={col}5*0.50", number_format='#,##0.00')

# 在建工程
cip_hist = [2.15, 3.82, 5.61, 2.57, 36.10]
set_cell(ws1, 29, 1, "  在建工程")
for i, v in enumerate(cip_hist):
    set_cell(ws1, 29, 2+i, v, font=input_font, fill=input_fill, number_format='#,##0.00')
for i in range(5):
    col = get_column_letter(7+i)
    set_cell(ws1, 29, 7+i, f"={col}5*0.15", number_format='#,##0.00')

# 无形资产
set_cell(ws1, 30, 1, "  无形资产")
for i in range(10):
    col = get_column_letter(2+i)
    set_cell(ws1, 30, 2+i, f"={col}5*0.03", number_format='#,##0.00')

# 非流动资产合计
merge_and_set(ws1, 31, 1, 1, "  非流动资产合计", font=output_font, fill=output_fill)
for i in range(10):
    col = get_column_letter(2+i)
    set_cell(ws1, 31, 2+i, f"=SUM(B28:B30)", number_format='#,##0.00')

# 资产总计
merge_and_set(ws1, 33, 1, 1, "资产总计", font=Font(name="微软雅黑", bold=True, size=10, color="000000"), fill=output_fill)
for i in range(10):
    col = get_column_letter(2+i)
    set_cell(ws1, 33, 2+i, f"={col}26+{col}31", number_format='#,##0.00')

# 负债
set_cell(ws1, 35, 1, "负债", font=Font(name="微软雅黑", bold=True, size=10))
# 短期借款
st_loan_hist = [10.52, 12.15, 14.38, 12.54, 15.00]
set_cell(ws1, 36, 1, "  短期借款")
for i, v in enumerate(st_loan_hist):
    set_cell(ws1, 36, 2+i, v, font=input_font, fill=input_fill, number_format='#,##0.00')
for i in range(5):
    col = get_column_letter(7+i)
    set_cell(ws1, 36, 7+i, f"={col}5*0.08", number_format='#,##0.00')

# 应付账款
ap_hist = [15.82, 18.35, 20.15, 27.61, 71.42]
set_cell(ws1, 37, 1, "  应付账款")
for i, v in enumerate(ap_hist):
    set_cell(ws1, 37, 2+i, v, font=input_font, fill=input_fill, number_format='#,##0.00')
for i in range(5):
    col = get_column_letter(7+i)
    set_cell(ws1, 37, 7+i, f"={col}5*0.37", number_format='#,##0.00')

# 其他流动负债
set_cell(ws1, 38, 1, "  其他流动负债")
for i in range(10):
    col = get_column_letter(2+i)
    set_cell(ws1, 38, 2+i, f"={col}5*0.08", number_format='#,##0.00')

# 流动负债合计
merge_and_set(ws1, 39, 1, 1, "  流动负债合计", font=output_font, fill=output_fill)
for i in range(10):
    col = get_column_letter(2+i)
    set_cell(ws1, 39, 2+i, f"=SUM(B36:B38)", number_format='#,##0.00')

# 长期借款
lt_loan_hist = [8.35, 10.52, 12.85, 23.10, 38.67]
set_cell(ws1, 41, 1, "  长期借款")
for i, v in enumerate(lt_loan_hist):
    set_cell(ws1, 41, 2+i, v, font=input_font, fill=input_fill, number_format='#,##0.00')
for i in range(5):
    col = get_column_letter(7+i)
    set_cell(ws1, 41, 7+i, f"={col}5*0.20", number_format='#,##0.00')

# 非流动负债合计
merge_and_set(ws1, 42, 1, 1, "  非流动负债合计", font=output_font, fill=output_fill)
for i in range(10):
    col = get_column_letter(2+i)
    set_cell(ws1, 42, 2+i, f"=B41", number_format='#,##0.00')

# 负债总计
merge_and_set(ws1, 44, 1, 1, "负债合计", font=output_font, fill=output_fill)
for i in range(10):
    col = get_column_letter(2+i)
    set_cell(ws1, 44, 2+i, f"={col}39+{col}42", number_format='#,##0.00')

# 股东权益
set_cell(ws1, 46, 1, "股东权益", font=Font(name="微软雅黑", bold=True, size=10))
set_cell(ws1, 47, 1, "  归母股东权益")
for i in range(5):
    col = get_column_letter(2+i)
    # Historical: for simplicity use total assets - liabilities
    set_cell(ws1, 47, 2+i, f"={col}33-{col}44", number_format='#,##0.00')
for i in range(5):
    col = get_column_letter(7+i)
    # Projected equity = prior equity + net income - dividends
    col_prev = get_column_letter(6+i)
    set_cell(ws1, 47, 7+i, f"={col_prev}47+{col}17*0.7", number_format='#,##0.00')

# 负债及权益总计
merge_and_set(ws1, 49, 1, 1, "负债及权益合计", font=Font(name="微软雅黑", bold=True, size=10), fill=output_fill)
for i in range(10):
    col = get_column_letter(2+i)
    set_cell(ws1, 49, 2+i, f"={col}44+{col}47", number_format='#,##0.00')

# 平衡检查
merge_and_set(ws1, 51, 1, 1, "平衡检查 (A - L - E = 0)", font=Font(name="微软雅黑", bold=True, color="FF0000", size=10))
for i in range(10):
    col = get_column_letter(2+i)
    set_cell(ws1, 51, 2+i, f"={col}33-{col}44-{col}47", number_format='#,##0.00')

# ============ 现金流量表 ============
merge_and_set(ws1, 53, 1, 11, "三、现金流量表（单位：亿元人民币）")

set_cell(ws1, 54, 1, "经营活动现金流")
# 经营活动现金流净额
ocf_hist = [5.82, 7.15, 6.38, 13.58, 46.03]
set_cell(ws1, 55, 1, "  经营活动现金流净额")
for i, v in enumerate(ocf_hist):
    set_cell(ws1, 55, 2+i, v, font=input_font, fill=input_fill, number_format='#,##0.00')
for i in range(5):
    col = get_column_letter(7+i)
    set_cell(ws1, 55, 7+i, f"={col}17*1.15", number_format='#,##0.00')

set_cell(ws1, 56, 1, "投资活动现金流")
icf_hist = [-8.25, -10.52, -12.15, -10.41, -67.35]
set_cell(ws1, 57, 1, "  投资活动现金流净额")
for i, v in enumerate(icf_hist):
    set_cell(ws1, 57, 2+i, v, font=input_font, fill=input_fill, number_format='#,##0.00')
for i in range(5):
    col = get_column_letter(7+i)
    set_cell(ws1, 57, 7+i, f"=-{col}5*0.35", number_format='#,##0.00')

set_cell(ws1, 58, 1, "筹资活动现金流")
fcf_hist = [3.15, 5.82, 8.35, -2.10, 43.77]
set_cell(ws1, 59, 1, "  筹资活动现金流净额")
for i, v in enumerate(fcf_hist):
    set_cell(ws1, 59, 2+i, v, font=input_font, fill=input_fill, number_format='#,##0.00')
for i in range(5):
    col = get_column_letter(7+i)
    # Gradually decreasing financing needs
    set_cell(ws1, 59, 7+i, f"={col}5*0.10", number_format='#,##0.00')

# 现金净变动
merge_and_set(ws1, 61, 1, 1, "  现金净变动", font=output_font, fill=output_fill)
for i in range(10):
    col = get_column_letter(2+i)
    set_cell(ws1, 61, 2+i, f"={col}55+{col}57+{col}59", number_format='#,##0.00')

# 期初现金
merge_and_set(ws1, 62, 1, 1, "  期初现金余额")
set_cell(ws1, 62, 2, 15.62, font=input_font, fill=input_fill, number_format='#,##0.00')
for i in range(1, 10):
    col_prev = get_column_letter(1+i)
    set_cell(ws1, 62, 2+i, f"={col_prev}63", number_format='#,##0.00')

# 期末现金
merge_and_set(ws1, 63, 1, 1, "  期末现金余额", font=output_font, fill=output_fill)
for i in range(10):
    col = get_column_letter(2+i)
    col_prev = get_column_letter(1+i) if i > 0 else None
    if i == 0:
        set_cell(ws1, 63, 2+i, f"={col}61+{col}62", number_format='#,##0.00')
    else:
        set_cell(ws1, 63, 2+i, f"={col}61+{col}62", number_format='#,##0.00')

# 现金核对 (期末现金 = 资产负债表货币资金)
merge_and_set(ws1, 65, 1, 1, "现金核对 (CF期末现金 vs BS货币资金)", font=Font(name="微软雅黑", bold=True, color="FF0000", size=10))
for i in range(10):
    col = get_column_letter(2+i)
    set_cell(ws1, 65, 2+i, f"={col}63-{col}22", number_format='#,##0.00')

# ============ 关键比率 ============
merge_and_set(ws1, 67, 1, 11, "四、关键财务比率")

ratios = [
    (68, "资产负债率", "=IFERROR({col}44/{col}33,\"\")", '0.0%'),
    (69, "流动比率", "=IFERROR({col}26/{col}39,\"\")", '0.00'),
    (70, "ROE(净资产收益率)", "=IFERROR({col}17/{col}47,\"\")", '0.0%'),
    (71, "ROA(总资产收益率)", "=IFERROR({col}17/{col}33,\"\")", '0.0%'),
    (72, "毛利率", "=IFERROR(({col}5-{col}7)/{col}5,\"\")", '0.0%'),
    (73, "净利率", "=IFERROR({col}17/{col}5,\"\")", '0.0%'),
    (74, "经营现金流/净利润", "=IFERROR({col}55/{col}17,\"\")", '0.00'),
    (75, "营收增长率", "=IFERROR({col}5/{get_column_letter(1+i)}5-1,\"\")", '0.0%'),
]

for row_num, label, formula_template, nf in ratios:
    set_cell(ws1, row_num, 1, label)
    is_growth = "增长率" in label
    for i in range(10):
        col = get_column_letter(2+i)
        if is_growth and i == 0:
            set_cell(ws1, row_num, 2+i, "", number_format=nf)
        else:
            formula = formula_template.replace("{col}", col)
            if is_growth:
                formula = formula.replace("{get_column_letter(1+i)}", get_column_letter(1+i))
            set_cell(ws1, row_num, 2+i, formula, number_format=nf)

# 设置列宽
ws1.column_dimensions['A'].width = 30
for c in range(2, 12):
    ws1.column_dimensions[get_column_letter(c)].width = 14

print("Sheet 1 (三表模型) created successfully!")

# ====================================================================
# Sheet 2: DCF估值
# ====================================================================
ws2 = wb.create_sheet("DCF估值")
ws2.sheet_properties.tabColor = "00B050"

# 标题
merge_and_set(ws2, 1, 1, 11, "胜宏科技(300476.SZ) DCF估值模型", Font(name="微软雅黑", bold=True, color="FFFFFF", size=14))
ws2.row_dimensions[1].height = 30

# 基础信息
merge_and_set(ws2, 3, 1, 11, "一、基本市场数据")
set_cell(ws2, 4, 1, "当前股价 (元)")
set_cell(ws2, 4, 2, 358.66, font=input_font, fill=input_fill, number_format='#,##0.00')
set_cell(ws2, 5, 1, "总股本 (亿股)")
set_cell(ws2, 5, 2, 9.83, font=input_font, fill=input_fill, number_format='#,##0.00')
set_cell(ws2, 6, 1, "总市值 (亿元)")
set_cell(ws2, 6, 2, "=B4*B5", number_format='#,##0.00')
set_cell(ws2, 7, 1, "净债务 (亿元)")
# Net Debt = Total Debt - Cash = (15.00+38.67) - 32.80 = 20.87
set_cell(ws2, 7, 2, 20.87, font=input_font, fill=input_fill, number_format='#,##0.00')

# 关键假设
merge_and_set(ws2, 9, 1, 11, "二、DCF关键假设")
merge_and_set(ws2, 10, 1, 8, "说明：在黄色输入区域更改假设值即可自动更新估值结果")

# 情景选择
set_cell(ws2, 12, 1, "情景选择 (1=保守, 2=基准, 3=乐观)")
set_cell(ws2, 12, 2, 2, font=input_font, fill=input_fill, number_format='0')
set_cell(ws2, 13, 1, "当前情景")
set_cell(ws2, 13, 2, '=IF(B12=1,"保守",IF(B12=2,"基准","乐观"))', font=Font(name="微软雅黑", bold=True, color="008000", size=10))

# WACC 计算
merge_and_set(ws2, 15, 1, 8, "三、WACC计算")
set_cell(ws2, 16, 1, "无风险利率 (10Y国债)")
set_cell(ws2, 16, 2, 0.028, font=input_font, fill=input_fill, number_format='0.0%')
set_cell(ws2, 17, 1, "Beta (5年)")
set_cell(ws2, 17, 2, 0.95, font=input_font, fill=input_fill, number_format='0.00')
set_cell(ws2, 18, 1, "股权风险溢价")
set_cell(ws2, 18, 2, 0.065, font=input_font, fill=input_fill, number_format='0.0%')
set_cell(ws2, 19, 1, "股权成本")
set_cell(ws2, 19, 2, "=B16+B17*B18", number_format='0.0%')
set_cell(ws2, 20, 1, "税前债务成本")
set_cell(ws2, 20, 2, 0.035, font=input_font, fill=input_fill, number_format='0.0%')
set_cell(ws2, 21, 1, "税率")
set_cell(ws2, 21, 2, 0.10, font=input_font, fill=input_fill, number_format='0.0%')  # 高新技术企业优惠税率
set_cell(ws2, 22, 1, "税后债务成本")
set_cell(ws2, 22, 2, "=B20*(1-B21)", number_format='0.0%')
set_cell(ws2, 23, 1, "债务权重")
set_cell(ws2, 23, 2, "=IF(B7<0,0,B7/(B6+B7))", number_format='0.0%')
set_cell(ws2, 24, 1, "股权权重")
set_cell(ws2, 24, 2, "=1-B23", number_format='0.0%')
set_cell(ws2, 25, 1, "WACC")
set_cell(ws2, 25, 2, "=B19*B24+B22*B23", number_format='0.00%')
set_cell(ws2, 25, 2, 0.10, font=input_font, fill=output_fill, number_format='0.0%')

# 三情景假设
merge_and_set(ws2, 27, 1, 11, "四、三情景增长假设")

# 列头
row_h = 28
set_cell(ws2, row_h, 1, "项目", font=sub_header_font, fill=sub_header_fill)
for i, yr in enumerate(["2026E","2027E","2028E","2029E","2030E"]):
    set_cell(ws2, row_h, 2+i, yr, font=sub_header_font, fill=sub_header_fill)

# 基准情景
merge_and_set(ws2, 29, 1, 6, "基准情景 (Base Case)", Font(name="微软雅黑", bold=True, color="000000", size=10),
              PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"))
set_cell(ws2, 30, 1, "营收增长率")
base_growth = [0.70, 0.50, 0.40, 0.30, 0.20]
for i, g in enumerate(base_growth):
    set_cell(ws2, 30, 2+i, g, font=input_font, fill=input_fill, number_format='0.0%')
set_cell(ws2, 31, 1, "EBIT利润率")
base_margin = [0.28, 0.30, 0.32, 0.33, 0.34]
for i, m in enumerate(base_margin):
    set_cell(ws2, 31, 2+i, m, font=input_font, fill=input_fill, number_format='0.0%')

# 保守情景
merge_and_set(ws2, 33, 1, 6, "保守情景 (Bear Case)", Font(name="微软雅黑", bold=True, color="000000", size=10),
              PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"))
set_cell(ws2, 34, 1, "营收增长率")
bear_growth = [0.55, 0.40, 0.30, 0.22, 0.15]
for i, g in enumerate(bear_growth):
    set_cell(ws2, 34, 2+i, g, font=input_font, fill=input_fill, number_format='0.0%')
set_cell(ws2, 35, 1, "EBIT利润率")
bear_margin = [0.25, 0.26, 0.27, 0.27, 0.28]
for i, m in enumerate(bear_margin):
    set_cell(ws2, 35, 2+i, m, font=input_font, fill=input_fill, number_format='0.0%')

# 乐观情景
merge_and_set(ws2, 37, 1, 6, "乐观情景 (Bull Case)", Font(name="微软雅黑", bold=True, color="000000", size=10),
              PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"))
set_cell(ws2, 38, 1, "营收增长率")
bull_growth = [0.85, 0.60, 0.50, 0.38, 0.25]
for i, g in enumerate(bull_growth):
    set_cell(ws2, 38, 2+i, g, font=input_font, fill=input_fill, number_format='0.0%')
set_cell(ws2, 39, 1, "EBIT利润率")
bull_margin = [0.30, 0.33, 0.35, 0.37, 0.38]
for i, m in enumerate(bull_margin):
    set_cell(ws2, 39, 2+i, m, font=input_font, fill=input_fill, number_format='0.0%')

# 永续增长率
set_cell(ws2, 41, 1, "永续增长率 (g)")
set_cell(ws2, 41, 2, 0.03, font=input_font, fill=input_fill, number_format='0.0%')  # 基准
set_cell(ws2, 41, 3, 0.025, font=input_font, fill=input_fill, number_format='0.0%')  # 保守
set_cell(ws2, 41, 4, 0.035, font=input_font, fill=input_fill, number_format='0.0%')  # 乐观

# 选择后的假设 (基于情景选择)
merge_and_set(ws2, 43, 1, 6, "选用假设 (基于情景选择)", Font(name="微软雅黑", bold=True, color="000000", size=10),
              PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid"))
set_cell(ws2, 44, 1, "项目")
for i, yr in enumerate(["2026E","2027E","2028E","2029E","2030E"]):
    set_cell(ws2, 44, 2+i, yr, font=sub_header_font, fill=sub_header_fill)

# 选用增长率 = INDEX(基准:乐观, 行, 情景选择)
for i in range(5):
    col = get_column_letter(2+i)
    set_cell(ws2, 45, 2+i, f"=CHOOSE($B$12,{col}34,{col}30,{col}38)",
             number_format='0.0%')
set_cell(ws2, 45, 1, "营收增长率")

for i in range(5):
    col = get_column_letter(2+i)
    set_cell(ws2, 46, 2+i, f"=CHOOSE($B$12,{col}35,{col}31,{col}39)",
             number_format='0.0%')
set_cell(ws2, 46, 1, "EBIT利润率")

set_cell(ws2, 47, 1, "永续增长率")
set_cell(ws2, 47, 2, f'=CHOOSE($B$12,$B$41,$C$41,$D$41)', number_format='0.0%')

# ============ 自由现金流预测 ============
merge_and_set(ws2, 49, 1, 11, "五、自由现金流预测（单位：亿元人民币）")

# 列头
set_cell(ws2, 50, 1, "项目", font=sub_header_font, fill=sub_header_fill)
proj_years = ["2025A","2026E","2027E","2028E","2029E","2030E"]
for i, yr in enumerate(proj_years):
    set_cell(ws2, 50, 2+i, yr, font=sub_header_font, fill=sub_header_fill)

# 营收
set_cell(ws2, 51, 1, "营业收入", font=Font(name="微软雅黑", bold=True, size=10))
set_cell(ws2, 51, 2, 192.92, font=input_font, fill=input_fill, number_format='#,##0.00')
for i in range(5):
    col_prev = get_column_letter(2+i)
    col_cur = get_column_letter(3+i)
    set_cell(ws2, 51, 3+i, f"={col_prev}51*(1+{col_cur}45)", number_format='#,##0.00')

# 营收增长率
set_cell(ws2, 52, 1, "  营收增长率")
set_cell(ws2, 52, 2, "", number_format='0.0%')
for i in range(5):
    col_prev = get_column_letter(2+i)
    col_cur = get_column_letter(3+i)
    set_cell(ws2, 52, 3+i, f"=IFERROR({col_cur}51/{col_prev}51-1,\"\")", number_format='0.0%')

# EBIT
set_cell(ws2, 54, 1, "EBIT", font=Font(name="微软雅黑", bold=True, size=10))
for i in range(6):
    col = get_column_letter(2+i)
    if i == 0:
        set_cell(ws2, 54, 2+i, 56.82, font=input_font, fill=input_fill, number_format='#,##0.00')
    else:
        set_cell(ws2, 54, 2+i, f"={col}51*{col}46", number_format='#,##0.00')

# EBIT利润率
set_cell(ws2, 55, 1, "  EBIT利润率")
for i in range(6):
    col = get_column_letter(2+i)
    set_cell(ws2, 55, 2+i, f"=IFERROR({col}54/{col}51,\"\")", number_format='0.0%')

# 税金
set_cell(ws2, 57, 1, "减: 所得税")
for i in range(6):
    col = get_column_letter(2+i)
    set_cell(ws2, 57, 2+i, f"={col}54*$B$21", number_format='#,##0.00')

# NOPAT
set_cell(ws2, 59, 1, "NOPAT (税后经营利润)", font=Font(name="微软雅黑", bold=True, size=10))
for i in range(6):
    col = get_column_letter(2+i)
    set_cell(ws2, 59, 2+i, f"={col}54-{col}57", number_format='#,##0.00')

# 折旧摊销
set_cell(ws2, 61, 1, "加: 折旧摊销")
for i in range(6):
    col = get_column_letter(2+i)
    set_cell(ws2, 61, 2+i, f"={col}51*0.05", number_format='#,##0.00')

# 资本支出
set_cell(ws2, 63, 1, "减: 资本支出")
for i in range(6):
    col = get_column_letter(2+i)
    if i == 0:
        set_cell(ws2, 63, 2+i, 67.35, font=input_font, fill=input_fill, number_format='#,##0.00')
    else:
        set_cell(ws2, 63, 2+i, f"={col}51*0.30", number_format='#,##0.00')

# 运营资本变动
set_cell(ws2, 65, 1, "减: 运营资本变动")
for i in range(6):
    col = get_column_letter(2+i)
    if i == 0:
        set_cell(ws2, 65, 2+i, "", number_format='#,##0.00')
    else:
        col_prev = get_column_letter(1+i)
        set_cell(ws2, 65, 2+i, f"=({col}51-{col_prev}51)*0.10", number_format='#,##0.00')

# 自由现金流
merge_and_set(ws2, 67, 1, 1, "自由现金流 (FCF)", font=output_font, fill=output_fill)
for i in range(6):
    col = get_column_letter(2+i)
    if i == 0:
        set_cell(ws2, 67, 2+i, "", number_format='#,##0.00')
    else:
        set_cell(ws2, 67, 2+i, f"={col}59+{col}61-{col}63-{col}65", number_format='#,##0.00')

# ============ 折现与估值 ============
merge_and_set(ws2, 69, 1, 11, "六、DCF折现与估值")

# 折现表头
set_cell(ws2, 70, 1, "项目", font=sub_header_font, fill=sub_header_fill)
dcf_labels = ["2026E","2027E","2028E","2029E","2030E","终值"]
for i, yr in enumerate(dcf_labels):
    set_cell(ws2, 70, 2+i, yr, font=sub_header_font, fill=sub_header_fill)

# FCF行
set_cell(ws2, 71, 1, "自由现金流")
for i in range(5):
    col = get_column_letter(3+i)
    set_cell(ws2, 71, 2+i, f"={col}67", number_format='#,##0.00')

# 折现期间
set_cell(ws2, 72, 1, "折现期间 (年中)")
disc_periods = [0.5, 1.5, 2.5, 3.5, 4.5]
for i, p in enumerate(disc_periods):
    set_cell(ws2, 72, 2+i, p, number_format='0.0')

# 折现因子
set_cell(ws2, 73, 1, "折现因子")
wacc_cell = "$B$25"
for i in range(5):
    col = get_column_letter(2+i)
    set_cell(ws2, 73, 2+i, f"=1/(1+{wacc_cell})^{col}72", number_format='0.0000')

# PV of FCF
set_cell(ws2, 74, 1, "FCF现值")
for i in range(5):
    col = get_column_letter(2+i)
    set_cell(ws2, 74, 2+i, f"={col}71*{col}73", number_format='#,##0.00')

# 终值计算
set_cell(ws2, 76, 1, "终值FCF")
set_cell(ws2, 76, 2, f"=G71*(1+$G$47)", number_format='#,##0.00')
set_cell(ws2, 77, 1, "终值")
set_cell(ws2, 77, 2, f"=B76/({wacc_cell}-$G$47)", number_format='#,##0.00')
set_cell(ws2, 78, 1, "终值现值")
# Terminal period = 5.5 for end of year 5 with mid-year convention
set_cell(ws2, 78, 2, f"=B77/(1+{wacc_cell})^5.5", number_format='#,##0.00')

# ============ 估值汇总 ============
merge_and_set(ws2, 80, 1, 11, "七、估值汇总")
summary_items = [
    (81, "FCF现值合计", "=SUM(B74:F74)", '#,##0.00'),
    (82, "终值现值", "=B78", '#,##0.00'),
    (83, "企业价值 (EV)", "=B81+B82", '#,##0.00'),
    (84, "减: 净债务", "=B7", '#,##0.00'),
    (85, "股权价值", "=B83-B84", '#,##0.00'),
    (86, "总股本 (亿股)", "=B5", '#,##0.00'),
    (88, "每股内在价值 (元)", "=IFERROR(B85/B86,\"\")", '#,##0.00'),
    (89, "当前股价 (元)", "=B4", '#,##0.00'),
    (90, "上涨空间", "=IFERROR(B88/B4-1,\"\")", '0.0%'),
]

for row_num, label, formula, nf in summary_items:
    if row_num == 88:
        merge_and_set(ws2, row_num, 1, 2, label, font=Font(name="微软雅黑", bold=True, size=12, color="000000"), fill=output_fill)
        set_cell(ws2, row_num, 3, formula, font=Font(name="微软雅黑", bold=True, size=14, color="0000FF"),
                 fill=output_fill, number_format=nf, border=thick_border)
    else:
        set_cell(ws2, row_num, 1, label, font=Font(name="微软雅黑", bold=True, size=10))
        if "B7" in formula or "B5" in formula or "B4" in formula:
            set_cell(ws2, row_num, 3, formula, number_format=nf)
        else:
            set_cell(ws2, row_num, 3, formula, number_format=nf)

# 风险警示
merge_and_set(ws2, 92, 1, 11, "风险提示", Font(name="微软雅黑", bold=True, color="FF0000", size=10))
set_cell(ws2, 93, 1, "本模型仅为分析参考，不构成投资建议。估值高度依赖假设参数，请审慎使用。",
         font=Font(name="微软雅黑", color="888888", size=9))

# ============ 敏感性分析 ============
merge_and_set(ws2, 95, 1, 11, "八、敏感性分析: WACC vs 永续增长率", Font(name="微软雅黑", bold=True, size=10))

# 5x5 敏感性表: WACC vs Terminal Growth
# WACC范围: 8%, 9%, 10%, 11%, 12%
# 永续增长率: 1.5%, 2.0%, 2.5%, 3.0%, 3.5%
wacc_values = [0.08, 0.09, 0.10, 0.11, 0.12]
tg_values = [0.015, 0.020, 0.025, 0.030, 0.035]

# 行头(永续增长率)
set_cell(ws2, 96, 1, "WACC \\ 永续增长率", font=sub_header_font, fill=sub_header_fill)
for i, tg in enumerate(tg_values):
    set_cell(ws2, 96, 2+i, tg, font=input_font, fill=input_fill, number_format='0.0%')

for j, wacc in enumerate(wacc_values):
    set_cell(ws2, 97+j, 1, wacc, font=input_font, fill=input_fill, number_format='0.0%')
    for i, tg in enumerate(tg_values):
        # Simplified: each cell formula = full DCF recalc using these WACC and TG values
        # For openpyxl, use formula that references row/col headers
        set_cell(ws2, 97+j, 2+i,
                 f"=((SUM(B71:F71))+((G71*(1+{get_column_letter(2+i)}96))/({get_column_letter(1+j)}97-{get_column_letter(2+i)}96))/(1+{get_column_letter(1+j)}97)^5.5-{wacc_cell})/{wacc_cell}",
                 number_format='#,##0.00')

ws2.column_dimensions['A'].width = 30
for c in range(2, 12):
    ws2.column_dimensions[get_column_letter(c)].width = 14

print("Sheet 2 (DCF估值) created successfully!")

# ============ 保存 ============
output_path = r"d:\claude-for-finance\outputs\胜宏科技\胜宏科技_三表模型_DCF估值.xlsx"
wb.save(output_path)
print(f"Excel model saved to: {output_path}")
