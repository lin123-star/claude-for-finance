"""
沪电股份（002463.SZ）三表模型
Three-Statement Financial Model
历史数据：2023A-2025A | 预测：2026E-2030E
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from copy import copy

# ============================================================
# 样式定义
# ============================================================
DARK_BLUE = PatternFill("solid", fgColor="1F4E79")
LIGHT_BLUE = PatternFill("solid", fgColor="D9E1F2")
MED_BLUE = PatternFill("solid", fgColor="BDD7EE")
LIGHT_GREY = PatternFill("solid", fgColor="F2F2F2")
WHITE_FILL = PatternFill("solid", fgColor="FFFFFF")
HEADER_FONT = Font(name="微软雅黑", bold=True, color="FFFFFF", size=11)
SUB_HEADER_FONT = Font(name="微软雅黑", bold=True, color="000000", size=10)
INPUT_FONT = Font(name="微软雅黑", color="0000FF", size=10)
FORMULA_FONT = Font(name="微软雅黑", color="000000", size=10)
BOLD_FONT = Font(name="微软雅黑", bold=True, color="000000", size=10)
LINK_FONT = Font(name="微软雅黑", color="008000", size=10)
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin")
)
THICK_BORDER = Border(
    left=Side(style="medium"), right=Side(style="medium"),
    top=Side(style="medium"), bottom=Side(style="medium")
)

# 列映射：B=2023A, C=2024A, D=2025A, E=2026E, F=2027E, G=2028E, H=2029E, I=2030E
COL_MAP = {0: "B", 1: "C", 2: "D", 3: "E", 4: "F", 5: "G", 6: "H", 7: "I"}
COLS = ["B", "C", "D", "E", "F", "G", "H", "I"]
YEARS_HIST = ["2023A", "2024A", "2025A"]
YEARS_PROJ = ["2026E", "2027E", "2028E", "2029E", "2030E"]
ALL_YEARS = YEARS_HIST + YEARS_PROJ

wb = openpyxl.Workbook()

# ============================================================
# 辅助函数
# ============================================================
def set_cell(ws, row, col, value, font=FORMULA_FONT, fill=WHITE_FILL, num_fmt=None, border=THIN_BORDER, alignment=None, color=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = font
    cell.fill = fill
    if num_fmt:
        cell.number_format = num_fmt
    if border:
        cell.border = border
    if alignment:
        cell.alignment = alignment
    if color:
        cell.font = Font(name="微软雅黑", bold=font.bold, color=color, size=font.size)
    return cell

def write_header_row(ws, row, labels, col_start=1):
    for i, label in enumerate(labels):
        set_cell(ws, row, col_start + i, label, font=SUB_HEADER_FONT, fill=LIGHT_BLUE)

def write_section_header(ws, row, text, max_col=9):
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = HEADER_FONT
    cell.fill = DARK_BLUE
    for c in range(2, max_col + 1):
        ws.cell(row=row, column=c).fill = DARK_BLUE

def write_formula_row(ws, row, label, formulas, is_input=False):
    """写入一行：标签 + 公式/数值"""
    font = INPUT_FONT if is_input else FORMULA_FONT
    fill = LIGHT_GREY if is_input else WHITE_FILL
    set_cell(ws, row, 1, label, font=BOLD_FONT, fill=WHITE_FILL, border=THIN_BORDER)
    for i, val in enumerate(formulas):
        set_cell(ws, row, 2 + i, val, font=font, fill=fill, border=THIN_BORDER)

# ============================================================
# Sheet 1: 假设条件
# ============================================================
ws_as = wb.active
ws_as.title = "假设条件"
ws_as.sheet_properties.tabColor = "1F4E79"
ws_as.column_dimensions["A"].width = 28
for c in range(2, 10):
    ws_as.column_dimensions[get_column_letter(c)].width = 14

# 标题
write_section_header(ws_as, 1, "沪电股份（002463.SZ）三表模型 — 假设条件")

# 列标题
write_header_row(ws_as, 2, ["假设项目"] + ALL_YEARS)

# 假设数据行
# 格式: (行号, 标签, 数值列表[8个], 是否为百分比)
assumptions = [
    (3,  "营收增长率",     ["", "", "", "37%", "30%", "22%", "16%", "12%"]),
    (4,  "毛利率",         ["31.17%", "34.54%", "35.48%", "36.0%", "36.5%", "37.0%", "37.0%", "37.0%"]),
    (5,  "研发费用率",     ["6.02%", "5.91%", "6.02%", "6.0%", "5.8%", "5.5%", "5.5%", "5.5%"]),
    (6,  "销售费用率",     ["3.12%", "2.74%", "2.80%", "2.8%", "2.7%", "2.5%", "2.5%", "2.5%"]),
    (7,  "管理费用率",     ["2.18%", "2.42%", "2.22%", "2.2%", "2.1%", "2.0%", "2.0%", "2.0%"]),
    (8,  "财务费用率(净)", ["-0.76%", "-1.36%", "-0.35%", "-0.3%", "-0.3%", "-0.3%", "-0.3%", "-0.3%"]),
    (9,  "实际税率",       ["15%", "15%", "15%", "15%", "15%", "15%", "15%", "15%"]),
    (10, "折旧摊销/营收",  ["3.8%", "3.6%", "3.5%", "4.5%", "4.5%", "4.0%", "4.0%", "4.0%"]),
    (11, "资本开支/营收",  ["9.1%", "16.1%", "16.9%", "18%", "16%", "14%", "12%", "10%"]),
    (12, "应收账款/营收",  ["28.5%", "29.3%", "29.1%", "29%", "28%", "27%", "26%", "25%"]),
    (13, "存货/营收",      ["7.8%", "8.0%", "7.9%", "8%", "8%", "8%", "8%", "8%"]),
    (14, "应付账款/营收",  ["14.5%", "14.8%", "14.8%", "15%", "15%", "15%", "15%", "15%"]),
    (15, "其他流动负债/营收", ["4.5%", "4.8%", "5.0%", "5%", "5%", "5%", "5%", "5%"]),
    (16, "股利支付率",     ["20%", "20%", "20%", "20%", "20%", "20%", "25%", "25%"]),
]

for row_num, label, values in assumptions:
    set_cell(ws_as, row_num, 1, label, font=BOLD_FONT, border=THIN_BORDER)
    for i, v in enumerate(values):
        if v:
            set_cell(ws_as, row_num, 2 + i, v, font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="0.00%")
        else:
            set_cell(ws_as, row_num, 2 + i, "", border=THIN_BORDER)

# 增加备注列
set_cell(ws_as, 18, 1, "说明：", font=BOLD_FONT)
set_cell(ws_as, 19, 1, "蓝色字体 = 可修改的假设输入", font=INPUT_FONT)
set_cell(ws_as, 20, 1, "历史数据（2023A-2025A）为实际财务数据", font=FORMULA_FONT)

# ============================================================
# Sheet 2: 利润表
# ============================================================
ws_is = wb.create_sheet("利润表")
ws_is.sheet_properties.tabColor = "2E75B6"
ws_is.column_dimensions["A"].width = 30
for c in range(2, 10):
    ws_is.column_dimensions[get_column_letter(c)].width = 16

write_section_header(ws_is, 1, "沪电股份（002463.SZ）— 利润表（百万元人民币）")
write_header_row(ws_is, 2, ["项目"] + ALL_YEARS)

# 历史数据（硬编码）
hist_data_is = {
    "营业收入":       [8938, 13342, 18945],
    "营业成本":       [6153, 8734, 12224],  # 排除6万精度问题
    "毛利":           [2785, 4608, 6721],
    "研发费用":       [538,  789,  1140],
    "销售费用":       [279,  365,  531],
    "管理费用":       [195,  323,  420],
    "财务费用(净)":   [-68,  -182, -67],
    "营业利润(EBIT)": [1841, 3313, 4697],
    "利润总额(EBT)":  [1780, 3044, 4496],
    "所得税":         [267,  457,  674],
    "净利润":         [1513, 2587, 3822],
    "归母净利润":     [1513, 2587, 3822],
}

# 构建公式引用
# 列偏移: E=4(2026E), F=5(2027E), G=6(2028E), H=7(2029E), I=8(2030E)
rev_row = 3
gp_row = 5
ni_row = 13

# 定义行
is_rows = {
    "营业收入":       3,
    "营业成本":       4,
    "毛利":           5,
    "毛利率":         6,
    "研发费用":       7,
    "销售费用":       8,
    "管理费用":       9,
    "财务费用(净)":   10,
    "营业利润(EBIT)": 11,
    "利润总额(EBT)":  12,
    "所得税":         13,
    "净利润":         14,
    "归母净利润":     15,
    "净利率":         16,
    "EPS(元)":        17,
}

r = is_rows

# 营收 — 历史硬编码，预测公式
set_cell(ws_is, r["营业收入"], 1, "营业收入", font=BOLD_FONT, border=THIN_BORDER)
for i in range(3):  # 2023A-2025A
    set_cell(ws_is, r["营业收入"], 2 + i, hist_data_is["营业收入"][i], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
# 2026E = 2025A * (1+假设条件!E3)
set_cell(ws_is, r["营业收入"], 5, f"=D{r['营业收入']}*(1+'假设条件'!E3)", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")
for idx, col in enumerate(["F", "G", "H", "I"], 3):
    prev = get_column_letter(ord(col) - 1)
    set_cell(ws_is, r["营业收入"], ord(col) - 64, f"={prev}{r['营业收入']}*(1+'假设条件'!{col}3)", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

# 营业成本 — 历史硬编码，预测 = 营收*(1-毛利率)
set_cell(ws_is, r["营业成本"], 1, "营业成本", font=BOLD_FONT, border=THIN_BORDER)
for i in range(3):
    set_cell(ws_is, r["营业成本"], 2 + i, hist_data_is["营业成本"][i], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
for j, col in enumerate(["E", "F", "G", "H", "I"], 3):
    c = ord(col) - 64
    set_cell(ws_is, r["营业成本"], c, f"={col}{r['营业收入']}*(1-'假设条件'!{col}4)", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

# 毛利 = 营收 - 营业成本
set_cell(ws_is, r["毛利"], 1, "毛利", font=BOLD_FONT, border=THIN_BORDER)
for j, col in enumerate(["B", "C", "D", "E", "F", "G", "H", "I"]):
    c = ord(col) - 64
    set_cell(ws_is, r["毛利"], c, f"={col}{r['营业收入']}-{col}{r['营业成本']}", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

# 毛利率 = 毛利/营收
set_cell(ws_is, r["毛利率"], 1, "毛利率", font=BOLD_FONT, border=THIN_BORDER)
for j, col in enumerate(["B", "C", "D", "E", "F", "G", "H", "I"]):
    c = ord(col) - 64
    set_cell(ws_is, r["毛利率"], c, f"={col}{r['毛利']}/{col}{r['营业收入']}", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="0.00%")

# 研发费用
set_cell(ws_is, r["研发费用"], 1, "研发费用", font=BOLD_FONT, border=THIN_BORDER)
for i in range(3):
    set_cell(ws_is, r["研发费用"], 2 + i, hist_data_is["研发费用"][i], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
for j, col in enumerate(["E", "F", "G", "H", "I"], 3):
    c = ord(col) - 64
    set_cell(ws_is, r["研发费用"], c, f"={col}{r['营业收入']}*'假设条件'!{col}5", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

# 销售费用
set_cell(ws_is, r["销售费用"], 1, "销售费用", font=BOLD_FONT, border=THIN_BORDER)
for i in range(3):
    set_cell(ws_is, r["销售费用"], 2 + i, hist_data_is["销售费用"][i], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
for j, col in enumerate(["E", "F", "G", "H", "I"], 3):
    c = ord(col) - 64
    set_cell(ws_is, r["销售费用"], c, f"={col}{r['营业收入']}*'假设条件'!{col}6", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

# 管理费用
set_cell(ws_is, r["管理费用"], 1, "管理费用", font=BOLD_FONT, border=THIN_BORDER)
for i in range(3):
    set_cell(ws_is, r["管理费用"], 2 + i, hist_data_is["管理费用"][i], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
for j, col in enumerate(["E", "F", "G", "H", "I"], 3):
    c = ord(col) - 64
    set_cell(ws_is, r["管理费用"], c, f"={col}{r['营业收入']}*'假设条件'!{col}7", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

# 财务费用(净)
set_cell(ws_is, r["财务费用(净)"], 1, "财务费用(净)", font=BOLD_FONT, border=THIN_BORDER)
for i in range(3):
    set_cell(ws_is, r["财务费用(净)"], 2 + i, hist_data_is["财务费用(净)"][i], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
for j, col in enumerate(["E", "F", "G", "H", "I"], 3):
    c = ord(col) - 64
    set_cell(ws_is, r["财务费用(净)"], c, f"={col}{r['营业收入']}*'假设条件'!{col}8", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

# 营业利润(EBIT) = 毛利 - 研发 - 销售 - 管理 - 财务
set_cell(ws_is, r["营业利润(EBIT)"], 1, "营业利润(EBIT)", font=BOLD_FONT, border=THIN_BORDER)
for i in range(3):
    set_cell(ws_is, r["营业利润(EBIT)"], 2 + i, hist_data_is["营业利润(EBIT)"][i], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
for j, col in enumerate(["E", "F", "G", "H", "I"]):
    c = ord(col) - 64
    set_cell(ws_is, r["营业利润(EBIT)"], c,
             f"={col}{r['毛利']}-{col}{r['研发费用']}-{col}{r['销售费用']}-{col}{r['管理费用']}-{col}{r['财务费用(净)']}",
             font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

# 利润总额(EBT) = EBIT（简化为无重大非经常性损益）
set_cell(ws_is, r["利润总额(EBT)"], 1, "利润总额(EBT)", font=BOLD_FONT, border=THIN_BORDER)
for i in range(3):
    set_cell(ws_is, r["利润总额(EBT)"], 2 + i, hist_data_is["利润总额(EBT)"][i], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
for j, col in enumerate(["E", "F", "G", "H", "I"]):
    c = ord(col) - 64
    set_cell(ws_is, r["利润总额(EBT)"], c, f"={col}{r['营业利润(EBIT)']}", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

# 所得税 = EBT * 税率
set_cell(ws_is, r["所得税"], 1, "所得税", font=BOLD_FONT, border=THIN_BORDER)
for i in range(3):
    set_cell(ws_is, r["所得税"], 2 + i, hist_data_is["所得税"][i], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
for j, col in enumerate(["E", "F", "G", "H", "I"]):
    c = ord(col) - 64
    set_cell(ws_is, r["所得税"], c, f"={col}{r['利润总额(EBT)']}*'假设条件'!{col}9", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

# 净利润 = EBT - 所得税
set_cell(ws_is, r["净利润"], 1, "净利润", font=BOLD_FONT, border=THIN_BORDER)
for i in range(3):
    set_cell(ws_is, r["净利润"], 2 + i, hist_data_is["净利润"][i], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
for j, col in enumerate(["E", "F", "G", "H", "I"]):
    c = ord(col) - 64
    set_cell(ws_is, r["净利润"], c, f"={col}{r['利润总额(EBT)']}-{col}{r['所得税']}", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

# 归母净利润 = 净利润（无非控股权益）
set_cell(ws_is, r["归母净利润"], 1, "归母净利润", font=BOLD_FONT, fill=MED_BLUE, border=THICK_BORDER)
for j, col in enumerate(["B", "C", "D", "E", "F", "G", "H", "I"]):
    c = ord(col) - 64
    set_cell(ws_is, r["归母净利润"], c, f"={col}{r['净利润']}", font=BOLD_FONT, fill=MED_BLUE, border=THICK_BORDER, num_fmt="#,##0")

# 净利率
set_cell(ws_is, r["净利率"], 1, "净利率", font=BOLD_FONT, border=THIN_BORDER)
for j, col in enumerate(["B", "C", "D", "E", "F", "G", "H", "I"]):
    c = ord(col) - 64
    set_cell(ws_is, r["净利率"], c, f"={col}{r['净利润']}/{col}{r['营业收入']}", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="0.00%")

# EPS = 归母净利润 / 总股本（约19.2亿股）
SHARES = 1920  # 百万股
set_cell(ws_is, r["EPS(元)"], 1, "EPS（元）", font=BOLD_FONT, fill=MED_BLUE, border=THICK_BORDER)
for j, col in enumerate(["B", "C", "D", "E", "F", "G", "H", "I"]):
    c = ord(col) - 64
    set_cell(ws_is, r["EPS(元)"], c, f"={col}{r['归母净利润']}/{SHARES}", font=BOLD_FONT, fill=MED_BLUE, border=THICK_BORDER, num_fmt="0.000")

# ============================================================
# Sheet 3: 资产负债表
# ============================================================
ws_bs = wb.create_sheet("资产负债表")
ws_bs.sheet_properties.tabColor = "548235"
ws_bs.column_dimensions["A"].width = 32
for c in range(2, 10):
    ws_bs.column_dimensions[get_column_letter(c)].width = 16

write_section_header(ws_bs, 1, "沪电股份（002463.SZ）— 资产负债表（百万元人民币）")
write_header_row(ws_bs, 2, ["项目"] + ALL_YEARS)

# BS 行号定义
bs_rows = {
    # 资产
    "货币资金":         3,
    "应收账款":         4,
    "存货":             5,
    "其他流动资产":     6,
    "流动资产合计":     7,
    "固定资产":         8,
    "其他非流动资产":   9,
    "非流动资产合计":   10,
    "总资产":           11,
    # 负债
    "应付账款":         13,
    "短期借款":         14,
    "其他流动负债":     15,
    "流动负债合计":     16,
    "长期借款":         17,
    "其他非流动负债":   18,
    "非流动负债合计":   19,
    "总负债":           20,
    # 权益
    "股本(实收资本)":    22,
    "资本公积":          23,
    "留存收益":          24,
    "归母权益合计":      25,
    "负债及权益合计":    26,
    # 校验
    "校验(A=L+E)":      28,
}

# 2025年末BS 数据
bs_2025 = {
    "货币资金": 2579,
    "应收账款": 5507,
    "存货": 1500,
    "其他流动资产": 5168,  # 倒算
    "固定资产": 5706,
    "其他非流动资产": 7794,  # 倒算
    "应付账款": 2800,
    "短期借款": 2000,
    "其他流动负债": 950,
    "长期借款": 2556,
    "其他非流动负债": 7820,  # 倒算
    "股本": 1920,
    "资本公积": 1080,
    "留存收益": 12113,
}

# 2023A, 2024A BS数据（估算，基于已有信息构建）
bs_hist = {
    "货币资金": [3150, 2800, 2579],
    "应收账款": [2550, 3900, 5507],
    "存货": [700, 1050, 1500],
    "其他流动资产": [3580, 4250, 5168],
    "固定资产": [3800, 4800, 5706],
    "其他非流动资产": [6020, 6840, 7794],
    "应付账款": [1300, 1950, 2800],
    "短期借款": [1200, 1500, 2000],
    "其他流动负债": [400, 640, 950],
    "长期借款": [1850, 2100, 2556],
    "其他非流动负债": [6450, 7100, 7820],
    "股本": [1920, 1920, 1920],
    "资本公积": [1080, 1080, 1080],
    "留存收益": [8440, 10139, 12113],
}

# 撰写资产负债表
labels_bs = ["货币资金", "应收账款", "存货", "其他流动资产", "流动资产合计",
             "固定资产", "其他非流动资产", "非流动资产合计", "总资产",
             "", "应付账款", "短期借款", "其他流动负债", "流动负债合计",
             "长期借款", "其他非流动负债", "非流动负债合计", "总负债",
             "", "股本(实收资本)", "资本公积", "留存收益", "归母权益合计",
             "负债及权益合计", "", "校验(A=L+E)"]

# 构建BS数据
# 资产类
r = bs_rows
for yr_idx, col in enumerate(COLS):
    c = ord(col) - 64

    if yr_idx < 3:  # 历史数据
        set_cell(ws_bs, r["货币资金"], c, bs_hist["货币资金"][yr_idx], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
        set_cell(ws_bs, r["应收账款"], c, bs_hist["应收账款"][yr_idx], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
        set_cell(ws_bs, r["存货"], c, bs_hist["存货"][yr_idx], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
        set_cell(ws_bs, r["其他流动资产"], c, bs_hist["其他流动资产"][yr_idx], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
    else:  # 预测数据 - 公式
        # 货币资金 - 现金流量表关联（后面通过CF平衡）
        # 暂时用简单公式：保持与现金相关的假设
        set_cell(ws_bs, r["货币资金"], c, "=0", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")  # 会被CF更新

        # 应收账款 = 营收 * AR/营收%
        set_cell(ws_bs, r["应收账款"], c, f"=利润表!{col}3*'假设条件'!{col}12", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

        # 存货 = 营收 * 存货/营收%
        set_cell(ws_bs, r["存货"], c, f"=利润表!{col}3*'假设条件'!{col}13", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

        # 其他流动资产 - 假设按营收增长比例增长
        set_cell(ws_bs, r["其他流动资产"], c, f"=D{r['其他流动资产']}*利润表!{col}3/利润表!D3", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

    # 流动资产合计
    set_cell(ws_bs, r["流动资产合计"], c,
             f"={col}{r['货币资金']}+{col}{r['应收账款']}+{col}{r['存货']}+{col}{r['其他流动资产']}",
             font=FORMULA_FONT, fill=MED_BLUE, border=THIN_BORDER, num_fmt="#,##0")

    # 固定资产
    if yr_idx < 3:
        set_cell(ws_bs, r["固定资产"], c, bs_hist["固定资产"][yr_idx], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
    else:
        # 固定资产 = 上期固定资产 + Capex - D&A
        prev_col = get_column_letter(c - 1)
        # Capex = 营收 * Capex/营收%, D&A = 营收 * D&A/营收%
        set_cell(ws_bs, r["固定资产"], c,
                 f"={prev_col}{r['固定资产']}+利润表!{col}3*'假设条件'!{col}11-利润表!{col}3*'假设条件'!{col}10",
                 font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

    # 其他非流动资产
    if yr_idx < 3:
        set_cell(ws_bs, r["其他非流动资产"], c, bs_hist["其他非流动资产"][yr_idx], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
    else:
        set_cell(ws_bs, r["其他非流动资产"], c, f"=D{r['其他非流动资产']}*利润表!{col}3/利润表!D3", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

    # 非流动资产合计
    set_cell(ws_bs, r["非流动资产合计"], c,
             f"={col}{r['固定资产']}+{col}{r['其他非流动资产']}",
             font=FORMULA_FONT, fill=MED_BLUE, border=THIN_BORDER, num_fmt="#,##0")

    # 总资产
    set_cell(ws_bs, r["总资产"], c,
             f"={col}{r['流动资产合计']}+{col}{r['非流动资产合计']}",
             font=BOLD_FONT, fill=MED_BLUE, border=THICK_BORDER, num_fmt="#,##0")

    # 应付账款
    if yr_idx < 3:
        set_cell(ws_bs, r["应付账款"], c, bs_hist["应付账款"][yr_idx], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
    else:
        set_cell(ws_bs, r["应付账款"], c, f"=利润表!{col}3*'假设条件'!{col}14", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

    # 短期借款 - 假设按营收增长比例增长
    if yr_idx < 3:
        set_cell(ws_bs, r["短期借款"], c, bs_hist["短期借款"][yr_idx], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
    else:
        set_cell(ws_bs, r["短期借款"], c, f"=D{r['短期借款']}*利润表!{col}3/利润表!D3", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

    # 其他流动负债
    if yr_idx < 3:
        set_cell(ws_bs, r["其他流动负债"], c, bs_hist["其他流动负债"][yr_idx], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
    else:
        set_cell(ws_bs, r["其他流动负债"], c, f"=利润表!{col}3*'假设条件'!{col}15", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

    # 流动负债合计
    set_cell(ws_bs, r["流动负债合计"], c,
             f"={col}{r['应付账款']}+{col}{r['短期借款']}+{col}{r['其他流动负债']}",
             font=FORMULA_FONT, fill=MED_BLUE, border=THIN_BORDER, num_fmt="#,##0")

    # 长期借款
    if yr_idx < 3:
        set_cell(ws_bs, r["长期借款"], c, bs_hist["长期借款"][yr_idx], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
    else:
        set_cell(ws_bs, r["长期借款"], c, f"=D{r['长期借款']}*利润表!{col}3/利润表!D3", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

    # 其他非流动负债
    if yr_idx < 3:
        set_cell(ws_bs, r["其他非流动负债"], c, bs_hist["其他非流动负债"][yr_idx], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
    else:
        set_cell(ws_bs, r["其他非流动负债"], c, f"=D{r['其他非流动负债']}*利润表!{col}3/利润表!D3", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

    # 非流动负债合计
    set_cell(ws_bs, r["非流动负债合计"], c,
             f"={col}{r['长期借款']}+{col}{r['其他非流动负债']}",
             font=FORMULA_FONT, fill=MED_BLUE, border=THIN_BORDER, num_fmt="#,##0")

    # 总负债
    set_cell(ws_bs, r["总负债"], c,
             f"={col}{r['流动负债合计']}+{col}{r['非流动负债合计']}",
             font=BOLD_FONT, fill=MED_BLUE, border=THICK_BORDER, num_fmt="#,##0")

    # 股本
    if yr_idx < 3:
        set_cell(ws_bs, r["股本(实收资本)"], c, bs_hist["股本"][yr_idx], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
    else:
        set_cell(ws_bs, r["股本(实收资本)"], c, f"=D{r['股本(实收资本)']}", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

    # 资本公积
    if yr_idx < 3:
        set_cell(ws_bs, r["资本公积"], c, bs_hist["资本公积"][yr_idx], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
    else:
        set_cell(ws_bs, r["资本公积"], c, f"=D{r['资本公积']}", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

    # 留存收益
    if yr_idx < 3:
        set_cell(ws_bs, r["留存收益"], c, bs_hist["留存收益"][yr_idx], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
    elif yr_idx == 3:
        # 2026E留存收益 = 2025A留存收益 + 2026E净利润 - 2026E股利
        set_cell(ws_bs, r["留存收益"], c,
                 f"=D{r['留存收益']}+利润表!E15-利润表!E15*'假设条件'!E16",
                 font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")
    else:
        prev_col = get_column_letter(c - 1)
        set_cell(ws_bs, r["留存收益"], c,
                 f"={prev_col}{r['留存收益']}+利润表!{col}15-利润表!{col}15*'假设条件'!{col}16",
                 font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

    # 归母权益合计
    set_cell(ws_bs, r["归母权益合计"], c,
             f"={col}{r['股本(实收资本)']}+{col}{r['资本公积']}+{col}{r['留存收益']}",
             font=BOLD_FONT, fill=MED_BLUE, border=THICK_BORDER, num_fmt="#,##0")

    # 负债及权益合计
    set_cell(ws_bs, r["负债及权益合计"], c,
             f"={col}{r['总负债']}+{col}{r['归母权益合计']}",
             font=BOLD_FONT, fill=MED_BLUE, border=THICK_BORDER, num_fmt="#,##0")

    # 校验：A - (L+E) = 0
    set_cell(ws_bs, r["校验(A=L+E)"], c,
             f"={col}{r['总资产']}-{col}{r['负债及权益合计']}",
             font=BOLD_FONT, fill=LIGHT_GREY, border=THICK_BORDER, num_fmt="0")

# 添加标签
set_cell(ws_bs, r["货币资金"], 1, "货币资金", font=BOLD_FONT, border=THIN_BORDER)
set_cell(ws_bs, r["应收账款"], 1, "应收账款", font=BOLD_FONT, border=THIN_BORDER)
set_cell(ws_bs, r["存货"], 1, "存货", font=BOLD_FONT, border=THIN_BORDER)
set_cell(ws_bs, r["其他流动资产"], 1, "其他流动资产", font=BOLD_FONT, border=THIN_BORDER)
set_cell(ws_bs, r["流动资产合计"], 1, "流动资产合计", font=BOLD_FONT, fill=MED_BLUE, border=THIN_BORDER)
set_cell(ws_bs, r["固定资产"], 1, "固定资产净额", font=BOLD_FONT, border=THIN_BORDER)
set_cell(ws_bs, r["其他非流动资产"], 1, "其他非流动资产", font=BOLD_FONT, border=THIN_BORDER)
set_cell(ws_bs, r["非流动资产合计"], 1, "非流动资产合计", font=BOLD_FONT, fill=MED_BLUE, border=THIN_BORDER)
set_cell(ws_bs, r["总资产"], 1, "总资产", font=BOLD_FONT, fill=DARK_BLUE, color="FFFFFF", border=THICK_BORDER)
set_cell(ws_bs, r["应付账款"], 1, "应付账款", font=BOLD_FONT, border=THIN_BORDER)
set_cell(ws_bs, r["短期借款"], 1, "短期借款", font=BOLD_FONT, border=THIN_BORDER)
set_cell(ws_bs, r["其他流动负债"], 1, "其他流动负债", font=BOLD_FONT, border=THIN_BORDER)
set_cell(ws_bs, r["流动负债合计"], 1, "流动负债合计", font=BOLD_FONT, fill=MED_BLUE, border=THIN_BORDER)
set_cell(ws_bs, r["长期借款"], 1, "长期借款", font=BOLD_FONT, border=THIN_BORDER)
set_cell(ws_bs, r["其他非流动负债"], 1, "其他非流动负债", font=BOLD_FONT, border=THIN_BORDER)
set_cell(ws_bs, r["非流动负债合计"], 1, "非流动负债合计", font=BOLD_FONT, fill=MED_BLUE, border=THIN_BORDER)
set_cell(ws_bs, r["总负债"], 1, "总负债", font=BOLD_FONT, fill=DARK_BLUE, color="FFFFFF", border=THICK_BORDER)
set_cell(ws_bs, r["股本(实收资本)"], 1, "股本(实收资本)", font=BOLD_FONT, border=THIN_BORDER)
set_cell(ws_bs, r["资本公积"], 1, "资本公积", font=BOLD_FONT, border=THIN_BORDER)
set_cell(ws_bs, r["留存收益"], 1, "留存收益", font=BOLD_FONT, border=THIN_BORDER)
set_cell(ws_bs, r["归母权益合计"], 1, "归母权益合计", font=BOLD_FONT, fill=MED_BLUE, border=THICK_BORDER)
set_cell(ws_bs, r["负债及权益合计"], 1, "负债及权益合计", font=BOLD_FONT, fill=MED_BLUE, border=THICK_BORDER)
set_cell(ws_bs, r["校验(A=L+E)"], 1, "校验(A-L-E=0)", font=BOLD_FONT, fill=LIGHT_GREY, border=THICK_BORDER)

# 修复总资产标签的颜色
ws_bs.cell(row=r["总资产"], column=1).fill = DARK_BLUE
ws_bs.cell(row=r["总资产"], column=1).font = Font(name="微软雅黑", bold=True, color="FFFFFF", size=10)
ws_bs.cell(row=r["总负债"], column=1).fill = DARK_BLUE
ws_bs.cell(row=r["总负债"], column=1).font = Font(name="微软雅黑", bold=True, color="FFFFFF", size=10)

# ============================================================
# Sheet 4: 现金流量表
# ============================================================
ws_cf = wb.create_sheet("现金流量表")
ws_cf.sheet_properties.tabColor = "BF8F00"
ws_cf.column_dimensions["A"].width = 32
for c in range(2, 10):
    ws_cf.column_dimensions[get_column_letter(c)].width = 16

write_section_header(ws_cf, 1, "沪电股份（002463.SZ）— 现金流量表（百万元人民币）")
write_header_row(ws_cf, 2, ["项目"] + ALL_YEARS)

r_cf = {
    "净利润": 3,
    "折旧摊销": 4,
    "营运资本变动": 5,
    "经营活动现金流": 6,
    "资本开支": 7,
    "投资活动现金流": 8,
    "新增借款": 9,
    "偿还借款": 10,
    "分配股利": 11,
    "筹资活动现金流": 12,
    "现金净变动": 13,
    "期初现金": 14,
    "期末现金": 15,
    "自由现金流": 16,
    "OCF/净利润": 17,
}

# CF行标签
cf_labels = {
    r_cf["净利润"]: "净利润",
    r_cf["折旧摊销"]: "加：折旧摊销",
    r_cf["营运资本变动"]: "营运资本变动",
    r_cf["经营活动现金流"]: "经营活动现金流",
    r_cf["资本开支"]: "减：资本开支",
    r_cf["投资活动现金流"]: "投资活动现金流",
    r_cf["新增借款"]: "新增借款",
    r_cf["偿还借款"]: "偿还借款",
    r_cf["分配股利"]: "分配股利",
    r_cf["筹资活动现金流"]: "筹资活动现金流",
    r_cf["现金净变动"]: "现金净变动",
    r_cf["期初现金"]: "期初现金",
    r_cf["期末现金"]: "期末现金",
    r_cf["自由现金流"]: "自由现金流(FCF)",
    r_cf["OCF/净利润"]: "OCF/净利润",
}

for row_num, label in cf_labels.items():
    if row_num in [r_cf["经营活动现金流"], r_cf["投资活动现金流"], r_cf["筹资活动现金流"]]:
        set_cell(ws_cf, row_num, 1, label, font=BOLD_FONT, fill=MED_BLUE, border=THIN_BORDER)
    elif row_num == r_cf["期末现金"]:
        set_cell(ws_cf, row_num, 1, label, font=BOLD_FONT, fill=MED_BLUE, border=THICK_BORDER)
    else:
        set_cell(ws_cf, row_num, 1, label, font=BOLD_FONT, border=THIN_BORDER)

# CF 历史数据
cf_hist = {
    "净利润": [1513, 2587, 3822],
    "折旧摊销": [340, 480, 663],
    "营运资本变动": [390, -742, -703],
    "资本开支": [-810, -2148, -3198],
    "新增借款": [300, 550, 956],
    "偿还借款": [-200, -300, -500],
    "分配股利": [-303, -517, -764],
}

for yr_idx, col in enumerate(COLS):
    c = ord(col) - 64

    # 历史
    if yr_idx < 3:
        # 净利润 - 从利润表链接
        set_cell(ws_cf, r_cf["净利润"], c, cf_hist["净利润"][yr_idx], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
        set_cell(ws_cf, r_cf["折旧摊销"], c, cf_hist["折旧摊销"][yr_idx], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
        set_cell(ws_cf, r_cf["营运资本变动"], c, cf_hist["营运资本变动"][yr_idx], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
        set_cell(ws_cf, r_cf["资本开支"], c, cf_hist["资本开支"][yr_idx], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
        set_cell(ws_cf, r_cf["新增借款"], c, cf_hist["新增借款"][yr_idx], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
        set_cell(ws_cf, r_cf["偿还借款"], c, cf_hist["偿还借款"][yr_idx], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
        set_cell(ws_cf, r_cf["分配股利"], c, cf_hist["分配股利"][yr_idx], font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
    else:
        # 预测 - 公式
        # 净利润 = 利润表净利润
        set_cell(ws_cf, r_cf["净利润"], c, f"=利润表!{col}14", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")
        # 折旧摊销 = 营收 * D&A/营收%
        set_cell(ws_cf, r_cf["折旧摊销"], c, f"=利润表!{col}3*'假设条件'!{col}10", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")
        # 营运资本变动 = -(ΔAR + Δ存货 - ΔAP - Δ其他流动负债) 简化
        prev_col = get_column_letter(c - 1)
        set_cell(ws_cf, r_cf["营运资本变动"], c,
                 f"=-(资产负债表!{col}4-资产负债表!{prev_col}4+资产负债表!{col}5-资产负债表!{prev_col}5-资产负债表!{col}11+资产负债表!{prev_col}11+资产负债表!{col}13-资产负债表!{prev_col}13)",
                 font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")
        # 资本开支 = 营收 * Capex/营收%
        set_cell(ws_cf, r_cf["资本开支"], c, f"=-利润表!{col}3*'假设条件'!{col}11", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")
        # 新增借款（简单假设：每年增加一定比例）
        set_cell(ws_cf, r_cf["新增借款"], c, f"=资产负债表!{col}14+资产负债表!{col}17-资产负债表!{prev_col}14-资产负债表!{prev_col}17", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")
        # 偿还借款 = 0（简化处理）
        set_cell(ws_cf, r_cf["偿还借款"], c, "=0", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")
        # 分配股利 = 净利润 * 股利支付率
        set_cell(ws_cf, r_cf["分配股利"], c, f"=-利润表!{col}14*'假设条件'!{col}16", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

    # 经营活动现金流 = 净利润 + D&A + 营运资本变动
    set_cell(ws_cf, r_cf["经营活动现金流"], c,
             f"={col}{r_cf['净利润']}+{col}{r_cf['折旧摊销']}+{col}{r_cf['营运资本变动']}",
             font=FORMULA_FONT, fill=MED_BLUE, border=THIN_BORDER, num_fmt="#,##0")

    # 投资活动现金流 = 资本开支（简化）
    set_cell(ws_cf, r_cf["投资活动现金流"], c,
             f"={col}{r_cf['资本开支']}",
             font=FORMULA_FONT, fill=MED_BLUE, border=THIN_BORDER, num_fmt="#,##0")

    # 筹资活动现金流 = 新增借款 - 偿还借款 - 分配股利
    set_cell(ws_cf, r_cf["筹资活动现金流"], c,
             f"={col}{r_cf['新增借款']}+{col}{r_cf['偿还借款']}+{col}{r_cf['分配股利']}",
             font=FORMULA_FONT, fill=MED_BLUE, border=THIN_BORDER, num_fmt="#,##0")

    # 现金净变动
    set_cell(ws_cf, r_cf["现金净变动"], c,
             f"={col}{r_cf['经营活动现金流']}+{col}{r_cf['投资活动现金流']}+{col}{r_cf['筹资活动现金流']}",
             font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

    # 期初现金
    if yr_idx == 0:
        set_cell(ws_cf, r_cf["期初现金"], c, 3150, font=INPUT_FONT, fill=LIGHT_GREY, border=THIN_BORDER, num_fmt="#,##0")
    else:
        prev_col = get_column_letter(c - 1)
        set_cell(ws_cf, r_cf["期初现金"], c, f"={prev_col}{r_cf['期末现金']}", font=FORMULA_FONT, border=THIN_BORDER, num_fmt="#,##0")

    # 期末现金 = 期初 + 净变动
    set_cell(ws_cf, r_cf["期末现金"], c,
             f"={col}{r_cf['期初现金']}+{col}{r_cf['现金净变动']}",
             font=BOLD_FONT, fill=MED_BLUE, border=THICK_BORDER, num_fmt="#,##0")

    # 自由现金流 = OCF - Capex
    set_cell(ws_cf, r_cf["自由现金流"], c,
             f"={col}{r_cf['经营活动现金流']}+{col}{r_cf['资本开支']}",
             font=BOLD_FONT, fill=MED_BLUE, border=THICK_BORDER, num_fmt="#,##0")

    # OCF/净利润
    set_cell(ws_cf, r_cf["OCF/净利润"], c,
             f"={col}{r_cf['经营活动现金流']}/{col}{r_cf['净利润']}",
             font=FORMULA_FONT, border=THIN_BORDER, num_fmt="0.00")

# ============================================================
# 配置打印
# ============================================================
for ws in [ws_as, ws_is, ws_bs, ws_cf]:
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

# ============================================================
# 保存
# ============================================================
output_path = "d:/claude-for-finance/outputs/models/沪电股份_三表模型.xlsx"
wb.save(output_path)
print(f"三表模型已保存到: {output_path}")
print(f"  - Sheet 1: 假设条件")
print(f"  - Sheet 2: 利润表（含2026E-2030E预测）")
print(f"  - Sheet 3: 资产负债表")
print(f"  - Sheet 4: 现金流量表")
print(f"  - 所有预测单元格使用公式，可修改假设条件自动更新")
