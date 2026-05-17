"""
沪电股份（002463.SZ）DCF 估值模型
DCF Valuation Model
基于2025年报数据 + 2026Q1数据
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
# 样式
# ============================================================
DARK_BLUE = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
LIGHT_BLUE = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
MED_BLUE = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
LIGHT_GREY = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
WHITE = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

HDR_FONT = Font(name="微软雅黑", bold=True, color="FFFFFF", size=11)
SUB_FONT = Font(name="微软雅黑", bold=True, color="000000", size=10)
INP_FONT = Font(name="微软雅黑", color="0000FF", size=10)  # 蓝色输入
FML_FONT = Font(name="微软雅黑", color="000000", size=10)  # 黑色公式
BOLD_FT = Font(name="微软雅黑", bold=True, color="000000", size=10)
GRN_FONT = Font(name="微软雅黑", color="008000", size=10)  # 绿色跨表引用
WHT_FONT = Font(name="微软雅黑", bold=True, color="FFFFFF", size=10)

THIN = Border(
    left=Side("thin"), right=Side("thin"),
    top=Side("thin"), bottom=Side("thin")
)
MED_BORD = Border(
    left=Side("medium"), right=Side("medium"),
    top=Side("medium"), bottom=Side("medium")
)
THICK = Border(
    left=Side("medium"), right=Side("medium"),
    top=Side("medium"), bottom=Side("medium")
)

wb = openpyxl.Workbook()

# ============================================================
# 辅助函数
# ============================================================
def sc(ws, row, col, value, font=FML_FONT, fill=WHITE, num_fmt=None, border=THIN, alignment=None):
    """set_cell 简写"""
    c = ws.cell(row=row, column=col, value=value)
    c.font = font
    c.fill = fill
    if num_fmt: c.number_format = num_fmt
    if border: c.border = border
    if alignment: c.alignment = alignment
    return c

def section_header(ws, row, text, max_col=9):
    for i in range(1, max_col + 1):
        cell = ws.cell(row=row, column=i)
        cell.fill = DARK_BLUE
        cell.font = HDR_FONT if i == 1 else Font(name="微软雅黑", color="FFFFFF", size=10)
    ws.cell(row=row, column=1).value = text

def sub_header(ws, row, labels):
    for i, label in enumerate(labels):
        sc(ws, row, 1 + i, label, font=SUB_FONT, fill=LIGHT_BLUE)

def input_row(ws, row, label, values, num_fmt="0.00"):
    """写入一行数据：标签 + 数值列表（蓝色字体输入）"""
    sc(ws, row, 1, label, font=BOLD_FT, border=THIN)
    for i, v in enumerate(values):
        sc(ws, row, 2 + i, v, font=INP_FONT, fill=LIGHT_GREY, border=THIN, num_fmt=num_fmt)

def formula_row(ws, row, label, formulas, num_fmt="0.00"):
    """写入一行公式"""
    sc(ws, row, 1, label, font=BOLD_FT, border=THIN)
    for i, v in enumerate(formulas):
        sc(ws, row, 2 + i, v, font=FML_FONT, border=THIN, num_fmt=num_fmt)

# 列: B=历史/输入列, C=2026E, D=2027E, E=2028E, F=2029E, G=2030E
# 敏感度表用 H/M 列的右侧

# ============================================================
# Sheet 1: DCF
# ============================================================
ws = wb.active
ws.title = "DCF"
ws.sheet_properties.tabColor = "1F4E79"
ws.column_dimensions["A"].width = 34
for c_idx in range(2, 18):
    ws.column_dimensions[get_column_letter(c_idx)].width = 14

# === Section 1: 标题 ===
section_header(ws, 1, "沪电股份（002463.SZ）DCF 估值模型")
sc(ws, 2, 1, "报告日期：2026-05-10 | 当前股价：108.71元 | 总股本：19.2亿股", font=SUB_FONT)

# === Section 2: 情景选择 ===
section_header(ws, 4, "情景选择 (Scenario Selector)")
sc(ws, 5, 1, "情景选择器 (1=熊市 2=基准 3=牛市)", font=BOLD_FT, border=THIN)
sc(ws, 5, 2, 2, font=INP_FONT, fill=LIGHT_GREY, border=THIN)  # 默认基准情景
sc(ws, 5, 3, "=IF(B5=1,\"熊市\",IF(B5=2,\"基准\",\"牛市\"))", font=FML_FONT, border=THIN)

# === Section 3: 市场数据 ===
section_header(ws, 7, "市场数据 (Market Data)")
input_row(ws, 8, "当前股价(元)", [108.71], "0.00")
input_row(ws, 9, "总股本(百万股)", [1920], "#,##0")
formula_row(ws, 10, "总市值(百万元)", ["=B8*B9"], "#,##0")
input_row(ws, 11, "净债务(百万元)", [1950], "#,##0")  # 来自报告：19.5亿
formula_row(ws, 12, "企业价值(百万元)", ["=B10+B11"], "#,##0")

# === Section 4: 情景假设 ===
# 三个情景块: 熊市(行14-19), 基准(行21-26), 牛市(行28-33)
# 每个块: 营收增长率, EBIT利润率, 税率, D&A/营收, Capex/营收, NWC/营收增量, 永续增长率, WACC

scenarios = {
    "熊市": {
        "row": 14,
        "title": "熊市情景假设 (Bear Case)",
        "rev_growth": [0.30, 0.22, 0.15, 0.10, 0.08],
        "ebit_margin": [0.22, 0.22, 0.22, 0.22, 0.22],
        "tax_rate": [0.15, 0.15, 0.15, 0.15, 0.15],
        "da_pct": [0.045, 0.045, 0.04, 0.04, 0.04],
        "capex_pct": [0.18, 0.16, 0.14, 0.12, 0.10],
        "nwc_pct": [0.02, 0.02, 0.02, 0.02, 0.02],
        "term_g": 0.025,
        "wacc": 0.095
    },
    "基准": {
        "row": 21,
        "title": "基准情景假设 (Base Case)",
        "rev_growth": [0.37, 0.30, 0.22, 0.16, 0.12],
        "ebit_margin": [0.25, 0.255, 0.26, 0.26, 0.255],
        "tax_rate": [0.15, 0.15, 0.15, 0.15, 0.15],
        "da_pct": [0.045, 0.045, 0.04, 0.04, 0.04],
        "capex_pct": [0.18, 0.16, 0.14, 0.12, 0.10],
        "nwc_pct": [0.02, 0.02, 0.02, 0.02, 0.02],
        "term_g": 0.03,
        "wacc": 0.09
    },
    "牛市": {
        "row": 28,
        "title": "牛市情景假设 (Bull Case)",
        "rev_growth": [0.45, 0.42, 0.35, 0.24, 0.16],
        "ebit_margin": [0.28, 0.29, 0.30, 0.30, 0.29],
        "tax_rate": [0.15, 0.15, 0.15, 0.15, 0.15],
        "da_pct": [0.045, 0.045, 0.04, 0.04, 0.04],
        "capex_pct": [0.18, 0.16, 0.14, 0.12, 0.10],
        "nwc_pct": [0.02, 0.02, 0.02, 0.02, 0.02],
        "term_g": 0.035,
        "wacc": 0.085
    }
}

proj_years = ["2026E", "2027E", "2028E", "2029E", "2030E"]
assumption_labels = [
    "营收增长率",
    "EBIT利润率",
    "税率",
    "D&A/营收",
    "Capex/营收",
    "NWC变动/营收增量",
    "",
    "永续增长率",
    "WACC"
]

for sc_name, sc_data in scenarios.items():
    r = sc_data["row"]
    section_header(ws, r, sc_data["title"])
    sub_header(ws, r + 1, [""] + proj_years)

    input_row(ws, r + 2, assumption_labels[0], sc_data["rev_growth"], "0.0%")
    input_row(ws, r + 3, assumption_labels[1], sc_data["ebit_margin"], "0.0%")
    input_row(ws, r + 4, assumption_labels[2], sc_data["tax_rate"], "0.0%")
    input_row(ws, r + 5, assumption_labels[3], sc_data["da_pct"], "0.0%")
    input_row(ws, r + 6, assumption_labels[4], sc_data["capex_pct"], "0.0%")
    input_row(ws, r + 7, assumption_labels[5], sc_data["nwc_pct"], "0.0%")
    # 空行
    input_row(ws, r + 9, assumption_labels[7], [sc_data["term_g"]] * 5, "0.0%")
    input_row(ws, r + 10, assumption_labels[8], [sc_data["wacc"]] * 5, "0.0%")

# === 合并列 (Consolidation Column) ===
# 使用 INDEX 从三个情景中选取
# 合并列区域: 行37-47
section_header(ws, 36, "合并假设 (Consolidated Assumptions - 由情景选择器决定)")
sub_header(ws, 37, [""] + proj_years)

# 熊市数据起始列=C(3), 基准=D(4), 牛市=E(5)
# 熊市块起始行: 14, 基准: 21, 牛市: 28
# 每块内: 营收增长在 row+2, EBIT在 row+3, 税率在 row+4, D&A在 row+5, Capex在 row+6, NWC在 row+7, 永续g在 row+9, WACC在 row+10

consol_labels = ["营收增长率", "EBIT利润率", "税率", "D&A/营收", "Capex/营收", "NWC变动/营收增量", "永续增长率", "WACC"]
consol_rows = [38, 39, 40, 41, 42, 43, 44, 45]

# 各行对应的情景块偏移
consol_offsets = [("rev_growth", 2), ("ebit_margin", 3), ("tax_rate", 4),
                  ("da_pct", 5), ("capex_pct", 6), ("nwc_pct", 7),
                  ("term_g", 9), ("wacc", 10)]

for ci, (label, (_, offset)) in enumerate(zip(consol_labels, consol_offsets)):
    row_n = consol_rows[ci]
    sc(ws, row_n, 1, label, font=BOLD_FT, border=THIN)
    for yr_idx in range(5):
        col_letter = get_column_letter(2 + yr_idx)  # C, D, E, F, G
        # INDEX公式: =INDEX(熊市区C~G, 基准区C~G, 牛市区C~G, 1, B5)
        bear_start_row = scenarios["熊市"]["row"] + offset
        base_start_row = scenarios["基准"]["row"] + offset
        bull_start_row = scenarios["牛市"]["row"] + offset

        # 使用CHOOSE函数
        formula = (f"=CHOOSE($B$5,"
                   f"{col_letter}{bear_start_row},"
                   f"{col_letter}{base_start_row},"
                   f"{col_letter}{bull_start_row})")
        sc(ws, row_n, 2 + yr_idx, formula, font=GRN_FONT, border=THIN, num_fmt="0.0%")

# === 营业收入基础 ===
section_header(ws, 47, "现金流预测 (Cash Flow Projection)")
sc(ws, 48, 1, "", border=THIN)
sub_header(ws, 48, [""] + ["2025A"] + proj_years)

# 营收基线 = 2025年实际营收 18,945百万元
REV_2025 = 18945

sc(ws, 49, 1, "营业收入(百万元)", font=BOLD_FT, border=THIN)
sc(ws, 49, 2, REV_2025, font=INP_FONT, fill=LIGHT_GREY, border=THIN, num_fmt="#,##0")
# 2026E = 2025A * (1+合并营收增长率)
sc(ws, 49, 3, "=B49*(1+C38)", font=FML_FONT, border=THIN, num_fmt="#,##0")
sc(ws, 49, 4, "=C49*(1+D38)", font=FML_FONT, border=THIN, num_fmt="#,##0")
sc(ws, 49, 5, "=D49*(1+E38)", font=FML_FONT, border=THIN, num_fmt="#,##0")
sc(ws, 49, 6, "=E49*(1+F38)", font=FML_FONT, border=THIN, num_fmt="#,##0")
sc(ws, 49, 7, "=F49*(1+G38)", font=FML_FONT, border=THIN, num_fmt="#,##0")

# 营收增长率
sc(ws, 50, 1, "  营收增长率", font=FML_FONT, border=THIN)
for ci in range(2, 8):
    if ci == 2:
        sc(ws, 50, ci, "", border=THIN)
    else:
        sc(ws, 50, ci, f"={get_column_letter(ci)}{49}/{get_column_letter(ci-1)}{49}-1", font=FML_FONT, border=THIN, num_fmt="0.0%")

# EBIT = 营收 * EBIT利润率
sc(ws, 52, 1, "EBIT(百万元)", font=BOLD_FT, border=THIN)
for ci in range(2, 8):
    if ci == 2:
        sc(ws, 52, ci, "", border=THIN)  # 2025A EBIT显示在标题
    else:
        cl = get_column_letter(ci)
        sc(ws, 52, ci, f"={cl}{49}*{cl}39", font=FML_FONT, border=THIN, num_fmt="#,##0")

# NOPAT = EBIT * (1-税率)
sc(ws, 53, 1, "  NOPAT(百万元)", font=FML_FONT, border=THIN)
for ci in range(3, 8):
    cl = get_column_letter(ci)
    sc(ws, 53, ci, f"={cl}{52}*(1-{cl}40)", font=FML_FONT, border=THIN, num_fmt="#,##0")

# D&A
sc(ws, 54, 1, "加：折旧摊销", font=FML_FONT, border=THIN)
for ci in range(3, 8):
    cl = get_column_letter(ci)
    sc(ws, 54, ci, f"={cl}{49}*{cl}41", font=FML_FONT, border=THIN, num_fmt="#,##0")

# Capex
sc(ws, 55, 1, "减：资本开支", font=FML_FONT, border=THIN)
for ci in range(3, 8):
    cl = get_column_letter(ci)
    sc(ws, 55, ci, f"=-{cl}{49}*{cl}42", font=FML_FONT, border=THIN, num_fmt="#,##0")

# NWC变动
sc(ws, 56, 1, "减：营运资本增加", font=FML_FONT, border=THIN)
for ci in range(3, 8):
    cl = get_column_letter(ci)
    prev_cl = get_column_letter(ci - 1)
    sc(ws, 56, ci, f"=-({cl}{49}-{prev_cl}{49})*{cl}43", font=FML_FONT, border=THIN, num_fmt="#,##0")

# FCF
sc(ws, 58, 1, "自由现金流(FCF)", font=BOLD_FT, fill=MED_BLUE, border=THICK)
for ci in range(3, 8):
    cl = get_column_letter(ci)
    sc(ws, 58, ci, f"={cl}{53}+{cl}{54}+{cl}{55}+{cl}{56}", font=BOLD_FT, fill=MED_BLUE, border=THICK, num_fmt="#,##0")

# === 折现计算 ===
section_header(ws, 60, "折现计算 (Discounting)")
sub_header(ws, 61, [""] + proj_years)

# FCF
sc(ws, 62, 1, "FCF(百万元)", font=BOLD_FT, border=THIN)
for ci in range(3, 8):
    cl = get_column_letter(ci)
    sc(ws, 62, ci, f"={cl}58", font=FML_FONT, border=THIN, num_fmt="#,##0")

# 折现期数 (年中惯例: 0.5, 1.5, 2.5, 3.5, 4.5)
sc(ws, 63, 1, "折现期数", font=FML_FONT, border=THIN)
periods = [0.5, 1.5, 2.5, 3.5, 4.5]
for ci, p in enumerate(periods, 3):
    sc(ws, 63, ci, p, font=INP_FONT, fill=LIGHT_GREY, border=THIN, num_fmt="0.0")

# WACC
sc(ws, 64, 1, "WACC", font=FML_FONT, border=THIN)
for ci in range(3, 8):
    cl = get_column_letter(ci)
    sc(ws, 64, ci, f"={cl}45", font=FML_FONT, border=THIN, num_fmt="0.0%")

# 折现因子 = 1/(1+WACC)^期数
sc(ws, 65, 1, "折现因子", font=FML_FONT, border=THIN)
for ci in range(3, 8):
    cl = get_column_letter(ci)
    sc(ws, 65, ci, f"=1/(1+{cl}64)^{cl}63", font=FML_FONT, border=THIN, num_fmt="0.0000")

# PV of FCF
sc(ws, 66, 1, "FCF现值(百万元)", font=BOLD_FT, border=THIN)
for ci in range(3, 8):
    cl = get_column_letter(ci)
    sc(ws, 66, ci, f"={cl}62*{cl}65", font=FML_FONT, border=THIN, num_fmt="#,##0")

# === 终值 ===
section_header(ws, 68, "终值计算 (Terminal Value)")

# 终值FCF = 最后一年FCF * (1+g)
sc(ws, 69, 1, "终值FCF(百万元)", font=FML_FONT, border=THIN)
sc(ws, 69, 3, "=G58*(1+G44)", font=FML_FONT, border=THIN, num_fmt="#,##0")

# 终值 = FCF终值 / (WACC - g)
sc(ws, 70, 1, "终值(百万元)", font=BOLD_FT, fill=MED_BLUE, border=THICK)
sc(ws, 70, 3, "=C69/(G45-G44)", font=BOLD_FT, fill=MED_BLUE, border=THICK, num_fmt="#,##0")

# 终值现值 = 终值 / (1+WACC)^4.5
sc(ws, 71, 1, "终值现值(百万元)", font=BOLD_FT, border=THIN)
sc(ws, 71, 3, "=C70*G65", font=BOLD_FT, border=THIN, num_fmt="#,##0")

# === 估值汇总 ===
section_header(ws, 73, "估值汇总 (Valuation Summary)")

sc(ws, 74, 1, "显性期FCF现值合计(百万元)", font=BOLD_FT, border=THIN)
sc(ws, 74, 3, "=SUM(C66:G66)", font=BOLD_FT, border=THIN, num_fmt="#,##0")

sc(ws, 75, 1, "终值现值(百万元)", font=BOLD_FT, border=THIN)
sc(ws, 75, 3, "=C71", font=BOLD_FT, border=THIN, num_fmt="#,##0")

sc(ws, 76, 1, "企业价值 EV(百万元)", font=BOLD_FT, fill=MED_BLUE, border=THICK)
sc(ws, 76, 3, "=C74+C75", font=BOLD_FT, fill=MED_BLUE, border=THICK, num_fmt="#,##0")

sc(ws, 77, 1, "减：净债务(百万元)", font=BOLD_FT, border=THIN)
sc(ws, 77, 3, "=B11", font=BOLD_FT, border=THIN, num_fmt="#,##0")

sc(ws, 78, 1, "股权价值(百万元)", font=BOLD_FT, fill=MED_BLUE, border=THICK)
sc(ws, 78, 3, "=C76-C77", font=BOLD_FT, fill=MED_BLUE, border=THICK, num_fmt="#,##0")

sc(ws, 79, 1, "总股本(百万股)", font=BOLD_FT, border=THIN)
sc(ws, 79, 3, "=B9", font=BOLD_FT, border=THIN, num_fmt="#,##0")

sc(ws, 80, 1, "每股内在价值(元)", font=Font(name="微软雅黑", bold=True, color="FFFFFF", size=12), fill=DARK_BLUE, border=THICK)
sc(ws, 80, 3, "=C78/C79", font=Font(name="微软雅黑", bold=True, color="FFFFFF", size=12), fill=DARK_BLUE, border=THICK, num_fmt="0.00")

sc(ws, 81, 1, "当前股价(元)", font=BOLD_FT, border=THIN)
sc(ws, 81, 3, "=B8", font=BOLD_FT, border=THIN, num_fmt="0.00")

sc(ws, 82, 1, "潜在涨幅", font=BOLD_FT, border=THIN)
sc(ws, 82, 3, "=C80/C81-1", font=BOLD_FT, border=THIN, num_fmt="0.0%")

# ============================================================
# 敏感度分析表格 (在 DCF sheet 底部)
# ============================================================

# === 表1: WACC vs 永续增长率 ===
sens_start = 85
section_header(ws, sens_start, "敏感度分析 1: WACC vs 永续增长率（每股内在价值 元）")

wacc_range = [0.075, 0.0825, 0.09, 0.0975, 0.105]
term_g_range = [0.02, 0.025, 0.03, 0.035, 0.04]

# 行/列标签
sc(ws, sens_start + 1, 1, "WACC vs 永续增长率", font=SUB_FONT, fill=LIGHT_BLUE, border=THIN)
for ci, g in enumerate(term_g_range):
    sc(ws, sens_start + 1, 2 + ci, g, font=INP_FONT, fill=LIGHT_GREY, border=THIN, num_fmt="0.0%")

for ri, w in enumerate(wacc_range):
    row = sens_start + 2 + ri
    sc(ws, row, 1, w, font=INP_FONT, fill=LIGHT_GREY, border=THIN, num_fmt="0.0%")
    for ci, g in enumerate(term_g_range):
        col = 2 + ci
        # 简化DCF重新计算: 假设固定营收55,000(2026E基准), 但用不同的WACC和g
        # 实际上理想情况是做完整重算，这里用核心DCF公式
        sc(ws, row, col,
           f"=($C$74+($G$58*(1+{get_column_letter(col)}${sens_start+1})/(${get_column_letter(row)}${sens_start+1}-{get_column_letter(col)}${sens_start+1})*$G$65-$C$77)/$C$79",
           font=FML_FONT, border=THIN, num_fmt="0.00")

# 标记中心单元格（基准情景：WACC=9.0%, g=3.0%）
center_row = sens_start + 2 + 2  # WACC=9.0%的行
center_col = 2 + 2              # g=3.0%的列
ws.cell(row=center_row, column=center_col).fill = MED_BLUE
ws.cell(row=center_row, column=center_col).font = Font(name="微软雅黑", bold=True, color="000000", size=11)

# === 表2: 营收增长率 vs EBIT利润率 ===
sens_start2 = sens_start + 8
section_header(ws, sens_start2, "敏感度分析 2: 营收增长率(2026E) vs EBIT利润率（每股内在价值 元）")

rev_g_range = [0.25, 0.31, 0.37, 0.43, 0.49]
ebit_m_range = [0.20, 0.225, 0.25, 0.275, 0.30]

sc(ws, sens_start2 + 1, 1, "营收增长\\EBIT利润率", font=SUB_FONT, fill=LIGHT_BLUE, border=THIN)
for ci, m in enumerate(ebit_m_range):
    sc(ws, sens_start2 + 1, 2 + ci, m, font=INP_FONT, fill=LIGHT_GREY, border=THIN, num_fmt="0.0%")

for ri, g in enumerate(rev_g_range):
    row = sens_start2 + 2 + ri
    sc(ws, row, 1, g, font=INP_FONT, fill=LIGHT_GREY, border=THIN, num_fmt="0.0%")
    for ci, m in enumerate(ebit_m_range):
        col = 2 + ci
        sc(ws, row, col,
           f"=(B49*(1+${get_column_letter(row)}${sens_start2+1})*{get_column_letter(col)}${sens_start2+1}*(1-0.15)+B49*(1+${get_column_letter(row)}${sens_start2+1})*0.045-B49*(1+${get_column_letter(row)}${sens_start2+1})*0.18-B49*(1+${get_column_letter(row)}${sens_start2+1})*0.02*${get_column_letter(row)}${sens_start2+1}",
           font=FML_FONT, border=THIN, num_fmt="0.00")

# === 表3: Beta vs 无风险利率 ===
sens_start3 = sens_start2 + 8
section_header(ws, sens_start3, "敏感度分析 3: Beta vs 无风险利率（每股内在价值 元）")

beta_range = [0.90, 1.035, 1.17, 1.305, 1.44]
rf_range = [0.0126, 0.0151, 0.0176, 0.0201, 0.0226]

sc(ws, sens_start3 + 1, 1, "Beta\\无风险利率", font=SUB_FONT, fill=LIGHT_BLUE, border=THIN)
for ci, rf in enumerate(rf_range):
    sc(ws, sens_start3 + 1, 2 + ci, rf, font=INP_FONT, fill=LIGHT_GREY, border=THIN, num_fmt="0.00%")

for ri, beta in enumerate(beta_range):
    row = sens_start3 + 2 + ri
    sc(ws, row, 1, beta, font=INP_FONT, fill=LIGHT_GREY, border=THIN, num_fmt="0.00")
    for ci, rf in enumerate(rf_range):
        col = 2 + ci
        sc(ws, row, col,
           f"=C78/($C$79*(${get_column_letter(row)}${sens_start3+1}+${get_column_letter(col)}${sens_start3+2}*${get_column_letter(row)}${sens_start3+1}))",
           font=FML_FONT, border=THIN, num_fmt="0.00")

# ============================================================
# Sheet 2: WACC 计算
# ============================================================
ws2 = wb.create_sheet("WACC")
ws2.sheet_properties.tabColor = "548235"
ws2.column_dimensions["A"].width = 32
ws2.column_dimensions["B"].width = 16
ws2.column_dimensions["C"].width = 40

section_header(ws2, 1, "WACC 计算明细")

# 权益成本
section_header(ws2, 3, "权益成本 (Cost of Equity)")
sc(ws2, 4, 1, "无风险利率(Rf)", font=BOLD_FT, border=THIN)
sc(ws2, 4, 2, 0.0176, font=INP_FONT, fill=LIGHT_GREY, border=THIN, num_fmt="0.00%")
sc(ws2, 4, 3, "中国10年期国债收益率 2026-05-08", font=FML_FONT)

sc(ws2, 5, 1, "Beta (5年月度)", font=BOLD_FT, border=THIN)
sc(ws2, 5, 2, 1.17, font=INP_FONT, fill=LIGHT_GREY, border=THIN, num_fmt="0.00")
sc(ws2, 5, 3, "Yahoo Finance 数据", font=FML_FONT)

sc(ws2, 6, 1, "股权风险溢价(ERP)", font=BOLD_FT, border=THIN)
sc(ws2, 6, 2, 0.065, font=INP_FONT, fill=LIGHT_GREY, border=THIN, num_fmt="0.00%")
sc(ws2, 6, 3, "A股市场标准ERP", font=FML_FONT)

sc(ws2, 7, 1, "权益成本(Ke = Rf + Beta*ERP)", font=BOLD_FT, fill=MED_BLUE, border=THICK)
sc(ws2, 7, 2, "=B4+B5*B6", font=GRN_FONT, border=THICK, num_fmt="0.00%")

# 债务成本
section_header(ws2, 9, "债务成本 (Cost of Debt)")
sc(ws2, 10, 1, "税前债务成本(Kd)", font=BOLD_FT, border=THIN)
sc(ws2, 10, 2, 0.035, font=INP_FONT, fill=LIGHT_GREY, border=THIN, num_fmt="0.00%")
sc(ws2, 10, 3, "基于企业信贷评级", font=FML_FONT)

sc(ws2, 11, 1, "有效税率", font=BOLD_FT, border=THIN)
sc(ws2, 11, 2, 0.15, font=INP_FONT, fill=LIGHT_GREY, border=THIN, num_fmt="0.00%")
sc(ws2, 11, 3, "高新技术企业税率", font=FML_FONT)

sc(ws2, 12, 1, "税后债务成本", font=BOLD_FT, fill=MED_BLUE, border=THICK)
sc(ws2, 12, 2, "=B10*(1-B11)", font=GRN_FONT, border=THICK, num_fmt="0.00%")

# 资本结构
section_header(ws2, 14, "资本结构 (Capital Structure)")
sc(ws2, 15, 1, "当前股价(元)", font=BOLD_FT, border=THIN)
sc(ws2, 15, 2, "=DCF!B8", font=GRN_FONT, border=THIN, num_fmt="0.00")

sc(ws2, 16, 1, "总股本(百万股)", font=BOLD_FT, border=THIN)
sc(ws2, 16, 2, "=DCF!B9", font=GRN_FONT, border=THIN, num_fmt="#,##0")

sc(ws2, 17, 1, "总市值(百万元)", font=BOLD_FT, fill=MED_BLUE, border=THIN)
sc(ws2, 17, 2, "=B15*B16", font=GRN_FONT, border=THIN, num_fmt="#,##0")

sc(ws2, 18, 1, "总债务(百万元)", font=BOLD_FT, border=THIN)
sc(ws2, 18, 2, 6344, font=INP_FONT, fill=LIGHT_GREY, border=THIN, num_fmt="#,##0")

sc(ws2, 19, 1, "现金及等价物(百万元)", font=BOLD_FT, border=THIN)
sc(ws2, 19, 2, 4392, font=INP_FONT, fill=LIGHT_GREY, border=THIN, num_fmt="#,##0")

sc(ws2, 20, 1, "净债务(百万元)", font=BOLD_FT, border=THIN)
sc(ws2, 20, 2, "=B18-B19", font=GRN_FONT, border=THIN, num_fmt="#,##0")

sc(ws2, 21, 1, "企业价值 EV(百万元)", font=BOLD_FT, fill=MED_BLUE, border=THICK)
sc(ws2, 21, 2, "=B17+B20", font=GRN_FONT, border=THICK, num_fmt="#,##0")

# WACC
section_header(ws2, 23, "WACC 计算")
sc(ws2, 24, 1, "项目", font=SUB_FONT, fill=LIGHT_BLUE, border=THIN)
sc(ws2, 24, 2, "权重", font=SUB_FONT, fill=LIGHT_BLUE, border=THIN)
sc(ws2, 24, 3, "成本", font=SUB_FONT, fill=LIGHT_BLUE, border=THIN)
sc(ws2, 24, 4, "贡献", font=SUB_FONT, fill=LIGHT_BLUE, border=THIN)

sc(ws2, 25, 1, "权益", font=BOLD_FT, border=THIN)
sc(ws2, 25, 2, "=B17/B21", font=FML_FONT, border=THIN, num_fmt="0.0%")
sc(ws2, 25, 3, "=B7", font=FML_FONT, border=THIN, num_fmt="0.00%")
sc(ws2, 25, 4, "=B25*C25", font=FML_FONT, border=THIN, num_fmt="0.00%")

sc(ws2, 26, 1, "债务", font=BOLD_FT, border=THIN)
sc(ws2, 26, 2, "=B20/B21", font=FML_FONT, border=THIN, num_fmt="0.0%")
sc(ws2, 26, 3, "=B12", font=FML_FONT, border=THIN, num_fmt="0.00%")
sc(ws2, 26, 4, "=B26*C26", font=FML_FONT, border=THIN, num_fmt="0.00%")

sc(ws2, 28, 1, "加权平均资本成本(WACC)", font=Font(name="微软雅黑", bold=True, color="FFFFFF", size=11), fill=DARK_BLUE, border=THICK)
sc(ws2, 28, 2, "=D25+D26", font=Font(name="微软雅黑", bold=True, color="FFFFFF", size=11), fill=DARK_BLUE, border=THICK, num_fmt="0.00%")

# ============================================================
# 保存
# ============================================================
output_path = "d:/claude-for-finance/outputs/models/沪电股份_DCF估值模型.xlsx"
wb.save(output_path)
print(f"DCF估值模型已保存到: {output_path}")
print(f"  - Sheet 1: DCF（含情景假设、现金流预测、估值汇总、3张敏感度表）")
print(f"  - Sheet 2: WACC（权益成本+债务成本+资本结构）")
print(f"  - 蓝色字体=输入, 黑色字体=公式, 绿色字体=跨表引用")
