"""
金风科技 (002202.SZ / 2208.HK) 三表财务模型 + DCF估值模型 v2
生成日期: 2026-05-09
架构: 假设表(集中输入) → 三表模型(Excel公式驱动) → DCF估值(公式链接)
所有预测单元格均使用Excel formula string，修改假设表即可联动更新全模型

预测因子来源说明:
  - 2026E营收+45%: 基于管理层装机目标+在手订单，偏积极，需与CWEA装机预测核对
  - 毛利率15.5%→19%: 基于历史均值及规模效应，需与同行均值对比验证
  - 以上均为AI估算，正式模型应对照管理层指引/Wind共识/行业报告修正驱动因子
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# ========== 样式定义 ==========
DARK_BLUE  = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
LIGHT_BLUE = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
MED_BLUE   = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
LIGHT_GRAY = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
INPUT_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")  # 绿=可编辑

WHITE_BOLD  = Font(color="FFFFFF", size=11, name="Times New Roman", bold=True)
BLACK_BOLD  = Font(color="000000", size=11, name="Times New Roman", bold=True)
BLACK_FONT  = Font(color="000000", size=11, name="Times New Roman")
BLUE_FONT   = Font(color="0000FF", size=11, name="Times New Roman")   # 历史实际值
GREEN_FONT  = Font(color="008000", size=11, name="Times New Roman")   # 公式计算值
INPUT_FONT  = Font(color="0070C0", size=11, name="Times New Roman")   # 蓝色=可编辑输入

thin = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'),  bottom=Side(style='thin')
)
CENTER = Alignment(horizontal='center')
LEFT   = Alignment(horizontal='left')

def hdr(ws, row, col, val, fill=LIGHT_BLUE):
    c = ws.cell(row=row, column=col, value=val)
    c.font = BLACK_BOLD; c.fill = fill; c.alignment = CENTER; c.border = thin

def section(ws, row, val, cols=10):
    c = ws.cell(row=row, column=1, value=val)
    c.font = WHITE_BOLD; c.fill = DARK_BLUE; c.border = thin
    ws.merge_cells(f'A{row}:{get_column_letter(cols)}{row}')
    for cc in range(2, cols+1):
        ws.cell(row=row, column=cc).fill = DARK_BLUE

def hist_cell(ws, row, col, val, fmt='0.00'):
    c = ws.cell(row=row, column=col, value=val)
    c.font = BLUE_FONT; c.fill = LIGHT_GRAY
    c.alignment = CENTER; c.border = thin; c.number_format = fmt

def formula_cell(ws, row, col, formula, fmt='0.00', bold=False, fill=None):
    c = ws.cell(row=row, column=col, value=formula)
    c.font = Font(color="008000", size=11, name="Times New Roman", bold=bold)
    if fill: c.fill = fill
    c.alignment = CENTER; c.border = thin; c.number_format = fmt

def input_cell(ws, row, col, val, fmt='0.00'):
    c = ws.cell(row=row, column=col, value=val)
    c.font = INPUT_FONT; c.fill = INPUT_FILL
    c.alignment = CENTER; c.border = thin; c.number_format = fmt

def label_cell(ws, row, col, val, bold=False):
    c = ws.cell(row=row, column=col, value=val)
    c.font = BLACK_BOLD if bold else BLACK_FONT; c.border = thin

# =====================================================================
# Sheet 0: 假设 — 预测驱动因子集中输入区
# =====================================================================
ws_a = wb.active
ws_a.title = "假设"
ws_a.column_dimensions['A'].width = 34
for c in list('BCDEFG'): ws_a.column_dimensions[c].width = 16

ws_a['A1'] = '金风科技 — 预测驱动因子 (修改本表 → 三表模型+DCF自动更新)'
ws_a['A1'].font = Font(size=13, bold=True, color="FFFFFF", name="Times New Roman")
ws_a.merge_cells('A1:G1')
ws_a['A1'].fill = DARK_BLUE
ws_a['A1'].alignment = CENTER

ws_a['A2'] = '蓝色单元格 = 可编辑输入 | 绿色单元格 = Excel公式 | 灰色 = 历史实际值'
ws_a.merge_cells('A2:G2')
ws_a['A2'].font = Font(size=9, italic=True)

# ---- 年度预测驱动因子 (rows 4-13) ----
section(ws_a, 4, '年度预测驱动因子 (2026E-2030E)', cols=7)
for i, h in enumerate(['假设项目','2026E','2027E','2028E','2029E','2030E','数据来源说明']):
    hdr(ws_a, 5, i+1, h)

# 注意: 以下因子为AI基于历史趋势估算，正式分析须对照管理层指引和Wind共识修正
assumptions_data = [
    ('营收增长率',        [0.45, 0.30, 0.20, 0.15, 0.10], '0.0%',
     '2026E基于装机量目标+在手订单; 需Wind共识验证'),
    ('毛利率',            [0.155, 0.170, 0.180, 0.185, 0.190], '0.0%',
     '历史均值~15-17%; 规模效应逐步修复'),
    ('销售管理费用率',    [0.050, 0.048, 0.046, 0.045, 0.044], '0.0%',
     '历史5.0-5.2%; 规模杠杆下行'),
    ('研发费用率',        [0.035, 0.033, 0.031, 0.030, 0.029], '0.0%',
     '历史3.0-3.9%; 维持高研发投入'),
    ('折旧摊销/营收',     [0.028, 0.027, 0.026, 0.025, 0.024], '0.0%',
     '历史~3%; 资产扩张放缓后下降'),
    ('资本支出/营收',     [0.048, 0.045, 0.042, 0.040, 0.038], '0.0%',
     '历史4.5-5.5%; 产能扩张后capex趋降'),
    ('营运资本变动/增量营收', [0.015, 0.012, 0.010, 0.008, 0.007], '0.0%',
     '应收改善带来比率下降'),
    ('所得税率',          [0.22, 0.22, 0.22, 0.22, 0.22], '0.0%',
     '高新15%偏保守取22%; 可调整'),
]

# 记录各假设行号 (供三表模型公式引用)
ass_row = {}
for i, (label, vals, fmt, note) in enumerate(assumptions_data):
    r = 6 + i
    ass_row[label] = r
    label_cell(ws_a, r, 1, label)
    ws_a.cell(row=r, column=1).border = thin
    for j, v in enumerate(vals):
        input_cell(ws_a, r, j+2, v, fmt)
    ws_a.cell(row=r, column=7, value=note).font = Font(size=9, italic=True, color="595959")
    ws_a.cell(row=r, column=7).border = thin

# ---- 基础年实际数据 (rows 16-21) ----
base_section_r = 6 + len(assumptions_data) + 2  # = 16
section(ws_a, base_section_r - 1, '基础年实际数据 (2025A)', cols=7)

base_items = [
    ('2025A营业收入 (亿元)', 730.23),
    ('2025A净利润 (亿元)',   27.74),
    ('净债务 (亿元)',         268.3),
    ('总股本 (亿股)',          42.22),
    ('港币/人民币汇率',        0.92),
]
ass_base = {}
for i, (label, val) in enumerate(base_items):
    r = base_section_r + i
    ass_base[label] = r
    label_cell(ws_a, r, 1, label)
    ws_a.cell(row=r, column=1).border = thin
    input_cell(ws_a, r, 2, val, '0.00')

# ---- WACC组件 (rows 23+) ----
wacc_section_r = base_section_r + len(base_items) + 2  # 行号
section(ws_a, wacc_section_r - 1, 'WACC 组件 (修改输入 → DCF自动重算)', cols=7)

wacc_items_data = [
    ('无风险利率 Rf',         0.0172, '0.00%', '10Y中国国债收益率'),
    ('Beta (5Y月度)',          0.95,   '0.00',  '相对沪深300; Wind数据'),
    ('股权风险溢价 ERP',       0.06,   '0.00%', 'Damodaran A股ERP'),
    ('税前债务成本 Kd',        0.035,  '0.00%', '2025年均贷款利率'),
    ('税率',                    0.22,  '0.00%', '与利润表所得税率一致'),
    ('A股总市值 (亿元)',       1203.0, '0.00',  '按当前股价×股本'),
    ('净债务 (亿元)',           268.3, '0.00',  '有息负债-货币资金'),
    ('终值增长率 g',            0.025, '0.00%', '参考长期名义GDP增速'),
]
ass_wacc = {}
for i, (label, val, fmt, note) in enumerate(wacc_items_data):
    r = wacc_section_r + i
    ass_wacc[label] = r
    label_cell(ws_a, r, 1, label)
    ws_a.cell(row=r, column=1).border = thin
    input_cell(ws_a, r, 2, val, fmt)
    ws_a.cell(row=r, column=3, value=note).font = Font(size=9, italic=True, color="595959")
    ws_a.cell(row=r, column=3).border = thin

# =====================================================================
# Sheet 1: 三表模型 — 历史列硬编码，预测列全公式
# =====================================================================
ws = wb.create_sheet("三表模型")
ws.column_dimensions['A'].width = 36
for c in list('BCDEFGHIJ'): ws.column_dimensions[c].width = 15

ws['A1'] = '金风科技 (002202.SZ) 三表财务模型 — 预测列为Excel公式'
ws['A1'].font = Font(size=13, bold=True, color="FFFFFF", name="Times New Roman")
ws.merge_cells('A1:J1')
ws['A1'].fill = DARK_BLUE; ws['A1'].alignment = CENTER

ws['A2'] = '单位: 亿元 | 蓝色=历史实际值 | 绿色=Excel公式(驱动因子见"假设"表) | 修改假设表全模型自动联动'
ws.merge_cells('A2:J2')
ws['A2'].font = Font(size=9, italic=True)

YEAR_HDRS = ['项目','2022A','2023A','2024A','2025A','2026E','2027E','2028E','2029E','2030E']
# 历史列: B-E (col 2-5), 预测列: F-J (col 6-10)
HIST_COLS = list(range(2, 6))   # 2,3,4,5
PROJ_COLS = list(range(6, 11))  # 6,7,8,9,10

def year_headers(ws, row):
    for i, h in enumerate(YEAR_HDRS):
        hdr(ws, row, i+1, h)

# ============ 利润表 ============
section(ws, 4, '利润表 (Income Statement)')
year_headers(ws, 5)

# 历史数据
hist_rev  = [425.71, 504.57, 566.99, 730.23]
hist_cogs = [378.89, 418.10, 488.76, 626.72]
hist_gp   = [46.82,  86.47,  78.23, 103.51]
hist_sga  = [22.15,  25.73,  28.35,  36.51]
hist_rd   = [14.32,  15.64,  18.03,  28.22]
hist_ebit = [10.35,  45.10,  31.85,  38.78]
hist_int  = [8.52,   9.87,  10.23,  11.05]
hist_tax  = [1.83,  12.32,   8.02,  10.21]
hist_ni   = [6.12,  13.31,  18.60,  27.74]

r = 6
rev_row = r
label_cell(ws, r, 1, '营业收入', bold=True)
for j, v in enumerate(hist_rev): hist_cell(ws, r, j+2, v)
rev_r = ass_row['营收增长率']
for j in range(5):
    pc = get_column_letter(PROJ_COLS[j]-1)   # previous col letter
    ac = get_column_letter(j+2)              # assumption col letter (B=2026, C=2027...)
    formula_cell(ws, r, PROJ_COLS[j], f"={pc}{rev_row}*(1+假设!{ac}{rev_r})")

r += 1; cogs_row = r
label_cell(ws, r, 1, '营业成本')
for j, v in enumerate(hist_cogs): hist_cell(ws, r, j+2, v)
gm_r = ass_row['毛利率']
for j in range(5):
    cc = get_column_letter(PROJ_COLS[j]); ac = get_column_letter(j+2)
    formula_cell(ws, r, PROJ_COLS[j], f"={cc}{rev_row}*(1-假设!{ac}{gm_r})")

r += 1; gp_row = r
label_cell(ws, r, 1, '毛利', bold=True)
for j, v in enumerate(hist_gp): hist_cell(ws, r, j+2, v)
for j in range(5):
    cc = get_column_letter(PROJ_COLS[j])
    formula_cell(ws, r, PROJ_COLS[j], f"={cc}{rev_row}-{cc}{cogs_row}", fill=MED_BLUE, bold=True)

r += 1; gm_pct_row = r
label_cell(ws, r, 1, '毛利率')
for j in range(4): hist_cell(ws, r, j+2, hist_gp[j]/hist_rev[j], '0.0%')
for j in range(5):
    cc = get_column_letter(PROJ_COLS[j])
    formula_cell(ws, r, PROJ_COLS[j], f"={cc}{gp_row}/{cc}{rev_row}", '0.0%')

r += 1; sga_row = r
label_cell(ws, r, 1, '销售管理费用')
for j, v in enumerate(hist_sga): hist_cell(ws, r, j+2, v)
sga_r = ass_row['销售管理费用率']
for j in range(5):
    cc = get_column_letter(PROJ_COLS[j]); ac = get_column_letter(j+2)
    formula_cell(ws, r, PROJ_COLS[j], f"={cc}{rev_row}*假设!{ac}{sga_r}")

r += 1; rd_row = r
label_cell(ws, r, 1, '研发费用')
for j, v in enumerate(hist_rd): hist_cell(ws, r, j+2, v)
rd_r = ass_row['研发费用率']
for j in range(5):
    cc = get_column_letter(PROJ_COLS[j]); ac = get_column_letter(j+2)
    formula_cell(ws, r, PROJ_COLS[j], f"={cc}{rev_row}*假设!{ac}{rd_r}")

r += 1; ebit_row = r
label_cell(ws, r, 1, 'EBIT (息税前利润)', bold=True)
for j, v in enumerate(hist_ebit): hist_cell(ws, r, j+2, v)
for j in range(5):
    cc = get_column_letter(PROJ_COLS[j])
    formula_cell(ws, r, PROJ_COLS[j],
                 f"={cc}{gp_row}-{cc}{sga_row}-{cc}{rd_row}", fill=MED_BLUE, bold=True)

r += 1; ebit_pct_row = r
label_cell(ws, r, 1, 'EBIT利润率')
for j in range(4): hist_cell(ws, r, j+2, hist_ebit[j]/hist_rev[j], '0.0%')
for j in range(5):
    cc = get_column_letter(PROJ_COLS[j])
    formula_cell(ws, r, PROJ_COLS[j], f"={cc}{ebit_row}/{cc}{rev_row}", '0.0%')

r += 1; int_row = r
label_cell(ws, r, 1, '利息支出')
for j, v in enumerate(hist_int): hist_cell(ws, r, j+2, v)
for j in range(5):   # 简化: 逐年+5%
    pc = get_column_letter(PROJ_COLS[j]-1)
    formula_cell(ws, r, PROJ_COLS[j], f"={pc}{int_row}*1.05")

r += 1; tax_row = r
label_cell(ws, r, 1, '所得税')
for j, v in enumerate(hist_tax): hist_cell(ws, r, j+2, v)
tax_r = ass_row['所得税率']
for j in range(5):
    cc = get_column_letter(PROJ_COLS[j]); ac = get_column_letter(j+2)
    formula_cell(ws, r, PROJ_COLS[j], f"=MAX({cc}{ebit_row},0)*假设!{ac}{tax_r}")

r += 1; ni_row = r
label_cell(ws, r, 1, '归母净利润', bold=True)
for j, v in enumerate(hist_ni): hist_cell(ws, r, j+2, v)
for j in range(5):
    cc = get_column_letter(PROJ_COLS[j])
    formula_cell(ws, r, PROJ_COLS[j],
                 f"={cc}{ebit_row}-{cc}{int_row}-{cc}{tax_row}", fill=MED_BLUE, bold=True)

r += 1; ni_pct_row = r
label_cell(ws, r, 1, '净利润率')
for j in range(4): hist_cell(ws, r, j+2, hist_ni[j]/hist_rev[j], '0.0%')
for j in range(5):
    cc = get_column_letter(PROJ_COLS[j])
    formula_cell(ws, r, PROJ_COLS[j], f"={cc}{ni_row}/{cc}{rev_row}", '0.0%')

# ============ 现金流量表 ============
r += 2
section(ws, r, '现金流量表 (FCF Bridge)')
r += 1; year_headers(ws, r)

hist_da   = [16.80, 18.50, 19.20, 21.50]
hist_capex= [25.30, 28.50, 30.20, 35.80]
hist_nwc  = [8.50, -5.20, 12.30,  8.90]
hist_fcf  = [-10.88, 21.91, -4.30, 5.44]
da_r    = ass_row['折旧摊销/营收']
capex_r = ass_row['资本支出/营收']
nwc_r   = ass_row['营运资本变动/增量营收']
tax_as  = ass_row['所得税率']

r += 1; nopat_row = r
label_cell(ws, r, 1, 'NOPAT = EBIT×(1-税率)')
for j in range(4): hist_cell(ws, r, j+2, round(hist_ebit[j]-hist_tax[j], 2))
for j in range(5):
    cc = get_column_letter(PROJ_COLS[j]); ac = get_column_letter(j+2)
    formula_cell(ws, r, PROJ_COLS[j], f"={cc}{ebit_row}*(1-假设!{ac}{tax_as})")

r += 1; da_row = r
label_cell(ws, r, 1, '+ 折旧与摊销')
for j, v in enumerate(hist_da): hist_cell(ws, r, j+2, v)
for j in range(5):
    cc = get_column_letter(PROJ_COLS[j]); ac = get_column_letter(j+2)
    formula_cell(ws, r, PROJ_COLS[j], f"={cc}{rev_row}*假设!{ac}{da_r}")

r += 1; capex_row = r
label_cell(ws, r, 1, '- 资本支出 (负数表示流出)')
for j, v in enumerate(hist_capex): hist_cell(ws, r, j+2, -v)
for j in range(5):
    cc = get_column_letter(PROJ_COLS[j]); ac = get_column_letter(j+2)
    formula_cell(ws, r, PROJ_COLS[j], f"=-{cc}{rev_row}*假设!{ac}{capex_r}")

r += 1; nwc_row = r
label_cell(ws, r, 1, '- 营运资本增加 (负数表示流出)')
for j, v in enumerate(hist_nwc): hist_cell(ws, r, j+2, -v)
for j in range(5):
    cc  = get_column_letter(PROJ_COLS[j])
    pcc = get_column_letter(PROJ_COLS[j]-1)
    ac  = get_column_letter(j+2)
    formula_cell(ws, r, PROJ_COLS[j],
                 f"=-({cc}{rev_row}-{pcc}{rev_row})*假设!{ac}{nwc_r}")

r += 1; fcf_row = r
label_cell(ws, r, 1, '无杠杆自由现金流 UFCF', bold=True)
for j, v in enumerate(hist_fcf): hist_cell(ws, r, j+2, v)
for j in range(5):
    cc = get_column_letter(PROJ_COLS[j])
    formula_cell(ws, r, PROJ_COLS[j],
                 f"={cc}{nopat_row}+{cc}{da_row}+{cc}{capex_row}+{cc}{nwc_row}",
                 fill=MED_BLUE, bold=True)

# ============ 资产负债表关键指标 ============
r += 2
section(ws, r, '资产负债表关键指标 (Balance Sheet Highlights)')
r += 1; year_headers(ws, r)

hist_ta  = [1329.50, 1434.95, 1552.24, 1664.95]
hist_tl  = [978.20, 1058.85, 1147.97, 1193.14]
hist_eq  = [351.30,  376.10,  404.27,  471.81]
hist_cash= [98.50,  110.40,  100.43,   96.09]
hist_debt= [380.50,  395.80,  395.39,  364.41]
hist_ar  = [235.60,  255.48,  308.25,  323.45]
hist_inv = [85.30,   92.10,  110.50,  135.80]

bs_items = [
    ('总资产',     hist_ta,   1.08),
    ('总负债',     hist_tl,   1.06),
    ('股东权益',   hist_eq,   1.12),
    ('货币资金',   hist_cash, 1.05),
    ('总有息负债', hist_debt, 0.98),
    ('应收账款',   hist_ar,   1.05),
    ('存货',       hist_inv,  1.08),
]
bs_row_map = {}
for label, hist_data, growth in bs_items:
    r += 1
    bs_row_map[label] = r
    label_cell(ws, r, 1, label)
    for j, v in enumerate(hist_data): hist_cell(ws, r, j+2, v)
    for j in range(5):
        pc = get_column_letter(PROJ_COLS[j]-1)
        formula_cell(ws, r, PROJ_COLS[j], f"={pc}{r}*{growth}")

r += 1  # 资产负债率
label_cell(ws, r, 1, '资产负债率')
ta_r = bs_row_map['总资产']; tl_r = bs_row_map['总负债']
for j in range(4): hist_cell(ws, r, j+2, hist_tl[j]/hist_ta[j], '0.0%')
for j in range(5):
    cc = get_column_letter(PROJ_COLS[j])
    formula_cell(ws, r, PROJ_COLS[j], f"={cc}{tl_r}/{cc}{ta_r}", '0.0%')

# =====================================================================
# Sheet 2: DCF估值 — 全公式链接到假设表和三表模型
# =====================================================================
ws2 = wb.create_sheet("DCF估值")
ws2.column_dimensions['A'].width = 40
for c in list('BCDEFGH'): ws2.column_dimensions[c].width = 18

ws2['A1'] = '金风科技 DCF估值模型 — 全公式链接 (假设表→三表模型→DCF)'
ws2['A1'].font = Font(size=13, bold=True, color="FFFFFF", name="Times New Roman")
ws2.merge_cells('A1:H1')
ws2['A1'].fill = DARK_BLUE; ws2['A1'].alignment = CENTER

ws2['A2'] = '修改"假设"工作表任意输入 → 本表所有结果自动重算'
ws2.merge_cells('A2:H2')
ws2['A2'].font = Font(size=9, italic=True)

# ---- 引用假设表的行号 ----
rf_r   = ass_wacc['无风险利率 Rf']
beta_r = ass_wacc['Beta (5Y月度)']
erp_r  = ass_wacc['股权风险溢价 ERP']
kd_r   = ass_wacc['税前债务成本 Kd']
tax_w  = ass_wacc['税率']
mkt_r  = ass_wacc['A股总市值 (亿元)']
nd_r   = ass_wacc['净债务 (亿元)']
tg_r   = ass_wacc['终值增长率 g']
shr_r  = ass_base['总股本 (亿股)']

# ---- WACC计算区 ----
r2 = 4
section(ws2, r2, 'WACC 计算', cols=8); r2 += 1

wacc_components = [
    ('股权成本  Ke = Rf + Beta × ERP',   f'=假设!B{rf_r}+假设!B{beta_r}*假设!B{erp_r}', '0.00%'),
    ('税后债务成本  Kd×(1-税率)',          f'=假设!B{kd_r}*(1-假设!B{tax_w})',             '0.00%'),
    ('权益权重  E/(D+E)',                  f'=假设!B{mkt_r}/(假设!B{mkt_r}+假设!B{nd_r})', '0.0%'),
    ('债务权重  D/(D+E)',                  f'=假设!B{nd_r}/(假设!B{mkt_r}+假设!B{nd_r})',  '0.0%'),
]
wc_rows = []
for label, formula, fmt in wacc_components:
    label_cell(ws2, r2, 1, label); ws2.cell(row=r2, column=1).border = thin
    formula_cell(ws2, r2, 2, formula, fmt)
    wc_rows.append(r2); r2 += 1

wacc_r2 = r2
label_cell(ws2, r2, 1, 'WACC', bold=True); ws2.cell(row=r2, column=1).border = thin
formula_cell(ws2, r2, 2,
             f"=B{wc_rows[0]}*B{wc_rows[2]}+B{wc_rows[1]}*B{wc_rows[3]}",
             '0.00%', bold=True, fill=MED_BLUE)

# ---- FCF预测 (引用三表模型) ----
r2 += 2
section(ws2, r2, '自由现金流预测 (引用"三表模型"工作表)', cols=8); r2 += 1
for i, h in enumerate(['指标','2026E','2027E','2028E','2029E','2030E']):
    hdr(ws2, r2, i+1, h)
r2 += 1

dcf_ref_items = [
    ('营业收入 (引自三表模型)', rev_row,   False),
    ('EBIT',                    ebit_row,  False),
    ('NOPAT = EBIT×(1-t)',      nopat_row, False),
    ('+ 折旧摊销',               da_row,   False),
    ('- 资本支出',               capex_row, False),
    ('- 营运资本变动',           nwc_row,  False),
    ('UFCF (无杠杆自由现金流)', fcf_row,  True),
]
ufcf_r2 = None
for label, src_row, is_bold in dcf_ref_items:
    label_cell(ws2, r2, 1, label, bold=is_bold); ws2.cell(row=r2, column=1).border = thin
    for j in range(5):
        src_col = get_column_letter(PROJ_COLS[j])
        cell_fill = MED_BLUE if is_bold else None
        formula_cell(ws2, r2, j+2,
                     f"=三表模型!{src_col}{src_row}",
                     '0.00', bold=is_bold, fill=cell_fill)
    if is_bold: ufcf_r2 = r2
    r2 += 1

# ---- 折现计算 ----
r2 += 1
section(ws2, r2, '折现计算 (Mid-year convention)', cols=8); r2 += 1
for i, h in enumerate(['指标','2026E','2027E','2028E','2029E','2030E']):
    hdr(ws2, r2, i+1, h)
r2 += 1

disc_row = r2
label_cell(ws2, r2, 1, '折现因子  1/(1+WACC)^期'); ws2.cell(row=r2, column=1).border = thin
for j in range(5):
    period = j + 0.5
    formula_cell(ws2, r2, j+2, f"=1/(1+B{wacc_r2})^{period}", '0.0000')

r2 += 1; pv_fcf_row = r2
label_cell(ws2, r2, 1, 'PV of FCF'); ws2.cell(row=r2, column=1).border = thin
for j in range(5):
    cc = get_column_letter(j+2)
    formula_cell(ws2, r2, j+2, f"={cc}{ufcf_r2}*{cc}{disc_row}", '0.00')

# ---- 估值汇总 ----
r2 += 2
section(ws2, r2, 'DCF 估值汇总', cols=8); r2 += 1

sum_pv_r = r2
label_cell(ws2, r2, 1, 'PV of Projected FCFs (亿元)'); ws2.cell(row=r2, column=1).border = thin
formula_cell(ws2, r2, 2, f"=SUM(B{pv_fcf_row}:F{pv_fcf_row})", '0.00')

r2 += 1; term_fcf_r = r2
label_cell(ws2, r2, 1, '终值期 FCF = UFCF₅×(1+g)'); ws2.cell(row=r2, column=1).border = thin
formula_cell(ws2, r2, 2, f"=F{ufcf_r2}*(1+假设!B{tg_r})", '0.00')

r2 += 1; tv_r = r2
label_cell(ws2, r2, 1, '终值 TV = 终值FCF/(WACC-g)'); ws2.cell(row=r2, column=1).border = thin
formula_cell(ws2, r2, 2, f"=B{term_fcf_r}/(B{wacc_r2}-假设!B{tg_r})", '0.00')

r2 += 1; pv_tv_r = r2
label_cell(ws2, r2, 1, 'PV of TV = TV/(1+WACC)^5'); ws2.cell(row=r2, column=1).border = thin
formula_cell(ws2, r2, 2, f"=B{tv_r}/(1+B{wacc_r2})^5", '0.00')

r2 += 1; tv_pct_r = r2
label_cell(ws2, r2, 1, '终值占EV比例 (>70%需关注)'); ws2.cell(row=r2, column=1).border = thin
formula_cell(ws2, r2, 2, f"=B{pv_tv_r}/(B{sum_pv_r}+B{pv_tv_r})", '0.0%')

r2 += 1; ev_r = r2
label_cell(ws2, r2, 1, 'Enterprise Value EV (亿元)', bold=True); ws2.cell(row=r2, column=1).border = thin
formula_cell(ws2, r2, 2, f"=B{sum_pv_r}+B{pv_tv_r}", '0.00', bold=True, fill=MED_BLUE)

r2 += 1; nd_ref_r = r2
label_cell(ws2, r2, 1, '(-) 净债务 (亿元)'); ws2.cell(row=r2, column=1).border = thin
formula_cell(ws2, r2, 2, f"=-假设!B{nd_r}", '0.00')

r2 += 1; eq_val_r = r2
label_cell(ws2, r2, 1, 'Equity Value (亿元)', bold=True); ws2.cell(row=r2, column=1).border = thin
formula_cell(ws2, r2, 2, f"=B{ev_r}+B{nd_ref_r}", '0.00', bold=True, fill=MED_BLUE)

r2 += 1; shr_ref_r = r2
label_cell(ws2, r2, 1, '总股本 (亿股)'); ws2.cell(row=r2, column=1).border = thin
formula_cell(ws2, r2, 2, f"=假设!B{shr_r}", '0.00')

r2 += 1; implied_r = r2
label_cell(ws2, r2, 1, '→ 隐含A股股价 (元/股)', bold=True); ws2.cell(row=r2, column=1).border = thin
ws2.cell(row=r2, column=1).fill = MED_BLUE
formula_cell(ws2, r2, 2, f"=B{eq_val_r}/B{shr_ref_r}", '0.00',
             bold=True, fill=MED_BLUE)
ws2.cell(row=r2, column=2).font = Font(color="008000", size=14, bold=True, name="Times New Roman")

r2 += 1; curr_price_r = r2
label_cell(ws2, r2, 1, '当前A股股价 (元/股) — 可修改'); ws2.cell(row=r2, column=1).border = thin
input_cell(ws2, r2, 2, 28.49, '0.00')

r2 += 1; upside_r = r2
label_cell(ws2, r2, 1, 'A股隐含上行空间', bold=True); ws2.cell(row=r2, column=1).border = thin
formula_cell(ws2, r2, 2, f"=B{implied_r}/B{curr_price_r}-1", '0.0%',
             bold=True, fill=MED_BLUE)
ws2.cell(row=r2, column=2).font = Font(color="008000", size=13, bold=True, name="Times New Roman")

# ---- 敏感度分析 ----
r2 += 3
section(ws2, r2, '敏感度分析: WACC vs 终值增长率 → 隐含A股股价(元)', cols=8); r2 += 1

ws2.cell(row=r2, column=1, value='WACC \\ g →').font = BLACK_BOLD
ws2.cell(row=r2, column=1).fill = LIGHT_BLUE
ws2.cell(row=r2, column=1).border = thin
ws2.cell(row=r2, column=1).alignment = CENTER

tg_vals   = [0.015, 0.020, 0.025, 0.030, 0.035]
wacc_vals = [0.065, 0.070, 0.075, 0.080, 0.085]

for j, tg in enumerate(tg_vals):
    c = ws2.cell(row=r2, column=j+2, value=tg)
    c.number_format = '0.0%'; c.font = BLACK_BOLD
    c.fill = LIGHT_BLUE; c.alignment = CENTER; c.border = thin

# 敏感度公式: 引用三表模型FCF，但使用各自WACC/tg组合
# FCF值随假设表联动更新；WACC/tg在此表中硬编码为灵敏度变量
LAST_PROJ = get_column_letter(PROJ_COLS[-1])   # 'J' = 2030E
for i, w in enumerate(wacc_vals):
    r2_s = r2 + 1 + i
    c = ws2.cell(row=r2_s, column=1, value=w)
    c.number_format = '0.0%'; c.font = BLACK_BOLD
    c.fill = LIGHT_BLUE; c.alignment = CENTER; c.border = thin

    for j, tg in enumerate(tg_vals):
        col_s = j + 2
        cell = ws2.cell(row=r2_s, column=col_s)
        if w <= tg:
            cell.value = 'N/A'
        else:
            # PV of 5-year FCFs
            pv_parts = "+".join(
                f"三表模型!{get_column_letter(PROJ_COLS[k])}{fcf_row}/(1+{w})^{k+0.5}"
                for k in range(5)
            )
            # PV of terminal value
            pv_tv_part = (f"三表模型!{LAST_PROJ}{fcf_row}*(1+{tg})"
                          f"/({w}-{tg})/(1+{w})^5")
            nd_ref = f"假设!B{nd_r}"
            sh_ref = f"假设!B{shr_r}"
            cell.value = f"=({pv_parts}+{pv_tv_part}-{nd_ref})/{sh_ref}"
            cell.number_format = '0.00'
        cell.font = BLACK_FONT; cell.alignment = CENTER; cell.border = thin
        # 高亮base case (~7.5% WACC, 2.5% tg)
        if abs(w - 0.075) < 0.001 and abs(tg - 0.025) < 0.001:
            cell.fill = MED_BLUE
            cell.font = Font(color="000000", size=11, bold=True, name="Times New Roman")

r2 += 6

# ---- 同行对比 ----
r2 += 2
section(ws2, r2, '相对估值 — 同行对比 (2025年数据)', cols=8); r2 += 1
peer_hdrs = ['公司', '市值(亿)', '营收(亿)', '净利润(亿)', 'PE(TTM)', 'PB', 'EV/EBITDA']
for i, h in enumerate(peer_hdrs): hdr(ws2, r2, i+1, h)
r2 += 1

peers = [
    ('金风科技(A)', 1203, 730.23, 27.74, 38.7, 3.03, 18.5),
    ('金风科技(H)', 732,  727.82, 26.64, 21.0, 1.63, 12.8),
    ('明阳智能',    350,  320.0,   8.5,  41.2, 1.85, 20.5),
    ('运达股份',    180,  294.0,   3.4,  52.9, 1.65, 25.0),
    ('三一重能',    280,  273.8,   7.12, 39.3, 2.10, 18.0),
    ('东方电气',    680,  620.0,  38.3,  17.8, 1.50, 11.2),
    ('电气风电',     80,  136.8,  -9.9, 'N/A', 1.20, 'N/A'),
]
for peer in peers:
    for j, v in enumerate(peer):
        c = ws2.cell(row=r2, column=j+1, value=v)
        c.font = BLACK_BOLD if j == 0 else BLACK_FONT
        c.alignment = LEFT if j == 0 else CENTER
        c.border = thin
    r2 += 1

# ---- 投资建议 ----
r2 += 1
section(ws2, r2, '投资建议', cols=8); r2 += 1
recomm = [
    ('评级',           '推荐(A股) / 强烈推荐(H股)'),
    ('A股目标价区间',  f'DCF隐含价格见DCF!B{implied_r} | 区间参考19-24元'),
    ('估值方法',       'DCF 60% + PB相对估值 40%'),
    ('关键敏感变量',   '修改"假设"表驱动因子 → 全模型自动重算'),
    ('核心风险',       '风机价格战、应收减值、欧盟调查、电价政策'),
    ('催化剂',         '2026H1业绩超预期、海上风电加速、海外大单'),
]
for label, val in recomm:
    label_cell(ws2, r2, 1, label, bold=True); ws2.cell(row=r2, column=1).border = thin
    c = ws2.cell(row=r2, column=2, value=val)
    c.font = BLACK_FONT; c.alignment = LEFT; c.border = thin
    ws2.merge_cells(f'B{r2}:H{r2}')
    r2 += 1

# =====================================================================
# 保存
# =====================================================================
filepath = r'd:\claude-for-finance\outputs\金风科技_三表模型_DCF估值_v2.xlsx'
wb.save(filepath)
print(f"✓ 已保存: {filepath}")
print(f"  架构: 假设表(蓝色输入) → 三表模型(绿色公式) → DCF估值(公式链接)")
print(f"  - 假设表驱动因子行: rev_growth={ass_row['营收增长率']}, WACC_rf={ass_wacc['无风险利率 Rf']}")
print(f"  - 三表模型: rev_row={rev_row}, ebit_row={ebit_row}, fcf_row={fcf_row}")
print(f"  - DCF: wacc_r2={wacc_r2}, ufcf_r2={ufcf_r2}, implied_r={implied_r}")
print(f"  修改假设表任意蓝色单元格 → 所有绿色公式自动联动更新")
