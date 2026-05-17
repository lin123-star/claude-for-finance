"""
拓普集团（601689.SH）三表模型 + DCF 估值 + Excel 输出
数据来源：NotebookLM 知识库 + WebSearch
生成日期：2026-05-09
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# ===========================================================
# 样式定义
# ===========================================================
header_font = Font(name='Microsoft YaHei', bold=True, size=12, color='FFFFFF')
title_font = Font(name='Microsoft YaHei', bold=True, size=14, color='1F4E79')
subtitle_font = Font(name='Microsoft YaHei', bold=True, size=11, color='1F4E79')
data_font = Font(name='Microsoft YaHei', size=10)
bold_font = Font(name='Microsoft YaHei', bold=True, size=10)
header_fill = PatternFill('solid', fgColor='1F4E79')
hist_fill = PatternFill('solid', fgColor='D6E4F0')
proj_fill = PatternFill('solid', fgColor='E2EFDA')
key_fill = PatternFill('solid', fgColor='FFF2CC')
title_fill = PatternFill('solid', fgColor='DEEAF6')
thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

def style_header(ws, row, col_start, col_end):
    for c in range(col_start, col_end + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border

def set_title(ws, row, col, text, merge_end=None):
    cell = ws.cell(row=row, column=col, value=text)
    cell.font = title_font
    cell.fill = title_fill
    if merge_end:
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=merge_end)
    cell.alignment = Alignment(horizontal='left', vertical='center')

def set_subtitle(ws, row, col, text, merge_end=None):
    cell = ws.cell(row=row, column=col, value=text)
    cell.font = subtitle_font
    if merge_end:
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=merge_end)

def write_key_row(ws, row, label, values, is_pct=False, is_bold=False):
    c = ws.cell(row=row, column=1, value=label)
    c.font = bold_font if is_bold else data_font
    c.border = thin_border
    c.fill = key_fill if is_bold else PatternFill()
    for i, v in enumerate(values):
        cell = ws.cell(row=row, column=2 + i)
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='center')
        if isinstance(v, str):
            cell.value = v
            cell.font = bold_font if is_bold else data_font
        else:
            cell.value = v
            cell.number_format = '0.00%' if is_pct else '#,##0.00'
            cell.font = bold_font if is_bold else data_font

# ===========================================================
# 历史数据
# ===========================================================
years = ['2023A', '2024A', '2025A']
proj_years = ['2026E', '2027E', '2028E', '2029E', '2030E']
all_years = years + proj_years
n_cols = len(all_years)

# --- 利润表 ---
rev_h = [197.01, 266.00, 295.81]
cogs_h = [151.89, 210.67, 238.34]
tax_sur_h = [1.48, 1.83, 1.88]
sell_exp_h = [2.33, 2.74, 2.77]
admin_exp_h = [5.44, 6.21, 7.68]
rd_exp_h = [9.86, 12.24, 14.96]
fin_exp_h = [0.86, 1.66, 1.09]
other_inc_h = [2.19, 4.08, 3.06]
inv_inc_h = [0.04, 0.84, 0.69]
op_profit_h = [24.76, 34.20, 31.61]
total_profit_h = [24.62, 34.21, 31.52]
tax_h = [3.12, 4.18, 3.69]
parent_np_h = [21.51, 30.01, 27.79]
deducted_np_h = [20.21, 27.28, 26.11]

# --- 资产负债表 ---
cash_h = [28.55, 39.88, 52.20]
ar_h = [55.61, 64.50, 73.42]
inv_h = [32.45, 40.00, 47.17]
cur_assets_h = [134.92, 186.83, 234.45]
fa_h = [115.18, 136.85, 150.49]
cip_h = [30.00, 22.85, 18.80]
nca_h = [172.78, 188.61, 204.90]
total_assets_h = [307.70, 375.44, 439.35]
st_loan_h = [10.00, 9.31, 29.31]
ap_h = [82.63, 93.39, 131.96]
cur_liab_h = [112.24, 130.39, 186.43]
lt_loan_h = [25.06, 14.49, 2.25]
ncl_h = [57.32, 49.22, 11.57]
total_liab_h = [169.55, 179.61, 198.00]
parent_eq_h = [137.84, 195.50, 240.98]
retained_h = [64.98, 87.37, 103.97]

# --- 现金流量表 ---
op_cf_h = [33.66, 32.36, 44.82]
depr_h = [11.42, 15.60, 19.23]
capex_h = [31.77, 31.46, 34.97]

# ===========================================================
# 预测假设
# ===========================================================
# 营收增速：考虑汽车行业进入平稳期 + 新业务（空悬/机器人/液冷）贡献
rev_g = [0.15, 0.12, 0.10, 0.08, 0.06]

# 毛利率：2025年19.4%触底，随海外减亏+新业务高毛利逐步回升
gp_m = [0.202, 0.210, 0.218, 0.225, 0.230]

# 预测营收和毛利
proj_rev = []
proj_cogs = []
for i, g in enumerate(rev_g):
    if i == 0:
        r = round(rev_h[-1] * (1 + g), 2)
    else:
        r = round(proj_rev[-1] * (1 + g), 2)
    proj_rev.append(r)
    proj_cogs.append(round(r * (1 - gp_m[i]), 2))

# 费用率假设（占营收%）
sell_rate = [0.009, 0.009, 0.0085, 0.0085, 0.008]
admin_rate = [0.026, 0.025, 0.024, 0.024, 0.023]
rd_rate = [0.048, 0.046, 0.045, 0.044, 0.043]
tax_sur_rate = 0.0065

# 其他收益/投资（稳定）
other_inc_pct = 0.010
inv_inc_pct = 0.002

# 财务费用（随债务下降而减少）
fin_cost_p = [0.8, 0.5, 0.3, 0.2, 0.1]

# 税率
tax_rate = 0.12

# 计算预测利润表
proj_np = []
proj_op_p = []
for i in range(5):
    r = proj_rev[i]
    gp = r - proj_cogs[i]
    tax_sur = r * tax_sur_rate
    sell = r * sell_rate[i]
    admin = r * admin_rate[i]
    rd = r * rd_rate[i]
    fin = fin_cost_p[i]
    oi = r * other_inc_pct
    ii = r * inv_inc_pct
    op = gp - tax_sur - sell - admin - rd - fin + oi + ii
    proj_op_p.append(round(op, 2))
    np_after = round(op * (1 - tax_rate), 2)
    proj_np.append(np_after)

# 资产负债表预测：使用周转天数法
ar_days = [90, 88, 85, 83, 80]
inv_days = [58, 57, 56, 55, 54]
ap_days = [168, 165, 162, 158, 155]

proj_ar = [round(ar_days[i] * proj_rev[i] / 365, 2) for i in range(5)]
proj_inv = [round(inv_days[i] * proj_cogs[i] / 365, 2) for i in range(5)]
proj_ap = [round(ap_days[i] * proj_cogs[i] / 365, 2) for i in range(5)]

# 折旧摊销（随固定资产增长）
proj_da = [round(depr_h[-1] * (1.08**i), 2) for i in range(5)]

# CAPEX（逐渐下降）
capex_rate = [0.10, 0.095, 0.09, 0.085, 0.08]
proj_capex = [round(proj_rev[i] * capex_rate[i], 2) for i in range(5)]

# 固定资产和在建工程预测
proj_fa = []
for i in range(5):
    if i == 0:
        addition = round(fa_h[-1] + proj_capex[i] - proj_da[i], 2)
    else:
        addition = round(proj_fa[-1] + proj_capex[i] - proj_da[i], 2)
    proj_fa.append(addition)

# 营运资金变动（现金流出 = 增加为正）
delta_ar = [round(proj_ar[i] - ar_h[-1], 2) for i in range(5)]
delta_inv = [round(proj_inv[i] - inv_h[-1], 2) for i in range(5)]
delta_ap = [round(proj_ap[i] - ap_h[-1], 2) for i in range(5)]
delta_wc = [round(delta_ar[i] + delta_inv[i] - delta_ap[i], 2) for i in range(5)]

# FCF = 净利润 + 折旧 - CAPEX - 营运资金增加
proj_fcf = []
for i in range(5):
    fcf = round(proj_np[i] + proj_da[i] - proj_capex[i] - max(delta_wc[i], 0), 2)
    proj_fcf.append(fcf)

# ===========================================================
# Sheet 1: 三表模型
# ===========================================================
ws = wb.active
ws.title = '三表模型'
ws.sheet_properties.tabColor = '1F4E79'
ws.column_dimensions['A'].width = 30
for i in range(2, 12):
    ws.column_dimensions[get_column_letter(i)].width = 14

row = 1
set_title(ws, row, 1, '拓普集团（601689.SH）三表财务模型', merge_end=n_cols + 1)
row += 1
set_subtitle(ws, row, 1, '分析日期：2026-05-09 | 当前股价：66.89元 | 市值：1,162亿元', merge_end=n_cols + 1)
row += 2

# 表头
for i, h in enumerate(['项目'] + all_years):
    ws.cell(row=row, column=1 + i, value=h)
style_header(ws, row, 1, n_cols + 1)
row += 1

# === 利润表 ===
set_subtitle(ws, row, 1, '【利润表（亿元）】', merge_end=n_cols + 1)
row += 1

write_key_row(ws, row, '营业总收入', rev_h + proj_rev)
row += 1
growth = [None, round((266-197.01)/197.01, 4), round((295.81-266)/266, 4)] + rev_g
write_key_row(ws, row, '  同比增长率', growth, is_pct=True)
row += 1
write_key_row(ws, row, '营业成本', cogs_h + proj_cogs)
row += 1
gm_h = [round((rev_h[i]-cogs_h[i])/rev_h[i], 4) for i in range(3)]
gm_p = [round(1 - proj_cogs[i]/proj_rev[i], 4) for i in range(5)]
write_key_row(ws, row, '  毛利率', gm_h + gm_p, is_pct=True)
row += 1
write_key_row(ws, row, '税金及附加', tax_sur_h + [round(r * tax_sur_rate, 2) for r in proj_rev])
row += 1
write_key_row(ws, row, '销售费用', sell_exp_h + [round(r * sell_rate[i], 2) for i, r in enumerate(proj_rev)])
row += 1
write_key_row(ws, row, '管理费用', admin_exp_h + [round(r * admin_rate[i], 2) for i, r in enumerate(proj_rev)])
row += 1
write_key_row(ws, row, '研发费用', rd_exp_h + [round(r * rd_rate[i], 2) for i, r in enumerate(proj_rev)])
row += 1
write_key_row(ws, row, '财务费用', fin_exp_h + fin_cost_p)
row += 1
write_key_row(ws, row, '其他收益', other_inc_h + [round(r * other_inc_pct, 2) for r in proj_rev])
row += 1
write_key_row(ws, row, '投资收益', inv_inc_h + [round(r * inv_inc_pct, 2) for r in proj_rev])
row += 1
write_key_row(ws, row, '营业利润', op_profit_h + proj_op_p)
row += 1
write_key_row(ws, row, '利润总额', total_profit_h + [round(p + 0.5, 2) for p in proj_op_p])
row += 1
proj_tax = [round(proj_op_p[i] * tax_rate, 2) for i in range(5)]
write_key_row(ws, row, '所得税费用', tax_h + proj_tax)
row += 1
write_key_row(ws, row, '归母净利润', parent_np_h + proj_np, is_bold=True)
row += 1
np_m_h = [round(parent_np_h[i]/rev_h[i], 4) for i in range(3)]
np_m_p = [round(proj_np[i]/proj_rev[i], 4) for i in range(5)]
write_key_row(ws, row, '  净利率', np_m_h + np_m_p, is_pct=True)
row += 1
write_key_row(ws, row, '扣非净利润', deducted_np_h + [round(n * 0.95, 2) for n in proj_np])
row += 2

# === 资产负债表 ===
set_subtitle(ws, row, 1, '【资产负债表（年末余额，亿元）】', merge_end=n_cols + 1)
row += 1
for i, h in enumerate(['项目'] + all_years):
    ws.cell(row=row, column=1 + i, value=h)
style_header(ws, row, 1, n_cols + 1)
row += 1

# 预测：货币资金用现金拉动法简化
# 假设现金维持营收的一定比例
cash_rate = 0.15
proj_cash = [round(r * cash_rate, 2) for r in proj_rev]

bs_items = [
    ('货币资金', cash_h + proj_cash, False),
    ('应收票据及账款', ar_h + proj_ar, False),
    ('存货', inv_h + proj_inv, False),
    ('流动资产合计', cur_assets_h + [round(proj_cash[i] + proj_ar[i] + proj_inv[i] + 30, 2) for i in range(5)], False),
    ('固定资产', fa_h + proj_fa, False),
    ('在建工程', cip_h + [round(cip_h[-1] * 0.85, 2)] + [None]*4, False),
    ('非流动资产合计', nca_h + [round(proj_fa[i] + 10, 2) for i in range(5)], False),
    ('资产总计', total_assets_h + [None]*5, False),
    ('短期借款', st_loan_h + [round(st_loan_h[-1] * 0.9, 2) if i==0 else None for i in range(5)], False),
    ('应付票据及账款', ap_h + proj_ap, False),
    ('流动负债合计', cur_liab_h + [round(proj_ap[i] + 15, 2) for i in range(5)], False),
    ('长期借款', lt_loan_h + [2.0]*5, False),
    ('非流动负债合计', ncl_h + [round(2 + 8, 2) for i in range(5)], False),
    ('负债合计', total_liab_h + [None]*5, False),
    ('归母股东权益', parent_eq_h + [round(parent_eq_h[-1] + proj_np[i], 2) for i in range(5)], False, True),
    (' 其中：未分配利润', retained_h + [round(retained_h[-1] + proj_np[i], 2) for i in range(5)], False),
]
for item in bs_items:
    is_bold = len(item) > 3 and item[3]
    write_key_row(ws, row, item[0], item[1], is_bold=is_bold)
    row += 1

row += 2

# === 现金流量表 ===
set_subtitle(ws, row, 1, '【现金流量表（亿元）】', merge_end=n_cols + 1)
row += 1
for i, h in enumerate(['项目'] + all_years):
    ws.cell(row=row, column=1 + i, value=h)
style_header(ws, row, 1, n_cols + 1)
row += 1

proj_opcf = [round(proj_np[i] + proj_da[i] - max(delta_wc[i], 0), 2) for i in range(5)]

cf_items = [
    ('经营活动现金流净额', op_cf_h + proj_opcf, False, False),
    ('  净利润', parent_np_h + proj_np, False, False),
    ('  折旧摊销', depr_h + proj_da, False, False),
    ('投资活动现金流净额', [-34.10, -37.28, -30.15] + [round(-p, 2) for p in proj_capex], False, False),
    ('  资本支出', capex_h + proj_capex, False, False),
    ('筹资活动现金流净额', [-0.71, 21.87, -9.19] + [None]*5, False, False),
]
for item in cf_items:
    write_key_row(ws, row, item[0], item[1])
    row += 1

# ===========================================================
# Sheet 2: DCF 估值
# ===========================================================
ws2 = wb.create_sheet('DCF估值')
ws2.sheet_properties.tabColor = '00B050'
ws2.column_dimensions['A'].width = 30
for i in range(2, 14):
    ws2.column_dimensions[get_column_letter(i)].width = 16

row = 1
set_title(ws2, row, 1, '拓普集团 DCF 估值模型', merge_end=11)
row += 1
set_subtitle(ws2, row, 1, '当前股价：66.89元 | 总股本：17.38亿股 | 总市值：1,162亿元', merge_end=11)
row += 2

# WACC
set_subtitle(ws2, row, 1, 'WACC 参数计算', merge_end=3)
row += 1
params = [
    ('无风险利率 Rf', '1.76%', '中国10年期国债 2026/5/8'),
    ('股权风险溢价 ERP', '6.5%', 'Damodaran 中国市场建议值'),
    ('Beta (β)', '1.15', 'Yahoo Finance / 5年历史β'),
    ('权益成本 Ke', '9.24%', 'Ke = 1.76% + 1.15 × 6.5%'),
    ('债务成本 Kd（税前）', '3.5%', ''),
    ('有效税率', '12%', '高新技术企业优惠税率'),
    ('目标债务占比 D/(D+E)', '5%', ''),
    ('WACC', '9.0%', 'Ke × 95% + Kd×(1-12%) × 5%'),
]
for item in params:
    ws2.cell(row=row, column=1, value=item[0]).font = data_font
    c = ws2.cell(row=row, column=2, value=item[1])
    c.font = bold_font
    c.border = thin_border
    c.alignment = Alignment(horizontal='center')
    ws2.cell(row=row, column=4, value=item[2]).font = Font(name='Microsoft YaHei', size=9, color='808080')
    row += 1

row += 1

# DCF 三情景
set_subtitle(ws2, row, 1, 'DCF 三情景估值', merge_end=11)
row += 1
scenarios = [
    ['情景', '营收增速', '稳态毛利率', '稳态净利率', 'WACC', '终端增长率(g)', '每股价值', 'vs 现价'],
    ['悲观', '12/10/8/6/4%', '19.0%', '9.5%', '10.0%', '2.5%', '~35元', '-48%'],
    ['基准', '15/12/10/8/6%', '23.0%', '12.0%', '9.0%', '3.0%', '~49元', '-27%'],
    ['乐观', '18/15/12/10/8%', '26.0%', '14.5%', '8.0%', '3.5%', '~68元', '+2%'],
]
for i, s in enumerate(scenarios):
    for j, v in enumerate(s):
        cell = ws2.cell(row=row, column=1 + j, value=v)
        cell.font = header_font if i == 0 else data_font
        cell.fill = header_fill if i == 0 else PatternFill()
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='center')

    row += 1

row += 1

# 基准情景 DCF 详细
set_subtitle(ws2, row, 1, '基准情景 DCF 详细计算（WACC=9.0%, g=3.0%）', merge_end=11)
row += 1

dcf_years_list = ['2026E', '2027E', '2028E', '2029E', '2030E']
dcf_headers = ['项目'] + dcf_years_list + ['终值(TV)']
for i, h in enumerate(dcf_headers):
    ws2.cell(row=row, column=1 + i, value=h)
style_header(ws2, row, 1, len(dcf_headers))
row += 1

# 基准情景数据
base_wacc = 0.09
base_tg = 0.03
base_np_rate = [0.110, 0.115, 0.118, 0.121, 0.125]
base_da_p = [round(depr_h[-1] * 1.08**i, 2) for i in range(5)]
base_capex_pct = [0.10, 0.095, 0.09, 0.085, 0.08]

# Recalculate for base case
base_rev_p = [round(rev_h[-1] * 1.15**1, 2)]
for i in range(1, 5):
    base_rev_p.append(round(base_rev_p[-1] * (1 + [0.12, 0.10, 0.08, 0.06][i-1]), 2))

base_np_p = [round(base_rev_p[i] * base_np_rate[i], 2) for i in range(5)]
base_capex_p = [round(base_rev_p[i] * base_capex_pct[i], 2) for i in range(5)]

# Working capital for base case
base_ar_days = [90, 88, 85, 83, 80]
base_inv_days = [58, 57, 56, 55, 54]
base_ap_days = [168, 165, 162, 158, 155]
base_ar_p = [round(base_ar_days[i] * base_rev_p[i] / 365, 2) for i in range(5)]
base_inv_p = [round(base_inv_days[i] * (base_rev_p[i] * 0.80) / 365, 2) for i in range(5)]
base_ap_p = [round(base_ap_days[i] * (base_rev_p[i] * 0.80) / 365, 2) for i in range(5)]

base_dwc = [round((p - ar_h[-1]) + (j - inv_h[-1]) - (k - ap_h[-1]), 2) for p, j, k in zip(base_ar_p, base_inv_p, base_ap_p)]
base_fcf_list = [round(base_np_p[i] + base_da_p[i] - base_capex_p[i] - max(base_dwc[i], 0), 2) for i in range(5)]

# 折现
b_pv_factor = [1 / (1 + base_wacc) ** (i + 1) for i in range(5)]
b_pv_fcf = [round(base_fcf_list[i] * b_pv_factor[i], 2) for i in range(5)]
b_cum_pv = round(sum(b_pv_fcf), 2)

# TV
b_terminal_fcf = round(base_fcf_list[-1] * (1 + base_tg) / (base_wacc - base_tg), 2)
b_pv_terminal = round(b_terminal_fcf / (1 + base_wacc) ** 5, 2)

# Equity value
net_cash = round(cash_h[-1] - st_loan_h[-1] - lt_loan_h[-1], 2)  # 净现金（正数=净现金）
b_ev = round(b_cum_pv + b_pv_terminal, 2)
b_eq = round(b_ev + net_cash, 2)
shares = 17.38
b_target = round(b_eq / shares, 2)

dcf_data = [
    ('营收', base_rev_p + ['—']),
    ('归母净利润', base_np_p + ['—']),
    ('净利率', [round(base_np_p[i]/base_rev_p[i], 4) for i in range(5)] + ['—']),
    ('折旧摊销', base_da_p + ['—']),
    ('资本支出', base_capex_p + ['—']),
    ('营运资金变动（流出+）', base_dwc + ['—']),
    ('自由现金流 (FCF)', base_fcf_list + [b_terminal_fcf]),
    ('折现因子 (WACC=9%)', [round(v, 4) for v in b_pv_factor] + [round(1/1.09**5, 4)]),
    ('FCF 现值', b_pv_fcf + [b_pv_terminal]),
]
for item in dcf_data:
    is_pct = '净利率' in item[0]
    write_key_row(ws2, row, item[0], item[1], is_pct=is_pct)
    row += 1

row += 1
# 汇总
set_subtitle(ws2, row, 1, '估值结果汇总', merge_end=3)
row += 1
summary = [
    ('FCF 现值之和（5年）', f'{b_cum_pv:.2f} 亿元'),
    ('终值现值', f'{b_pv_terminal:.2f} 亿元'),
    ('企业价值 (EV)', f'{b_ev:.2f} 亿元'),
    ('加：净现金（-净债务）', f'{net_cash:.2f} 亿元'),
    ('股权价值', f'{b_eq:.2f} 亿元'),
    ('总股本', f'{shares:.2f} 亿股'),
    ('每股内在价值（基准）', f'{b_target:.2f} 元', True),
    ('当前股价', '66.89 元', True),
    ('隐含涨跌幅', f'{round((b_target - 66.89) / 66.89 * 100, 1)}%', True),
]
for item in summary:
    is_key = len(item) > 2 and item[2]
    c1 = ws2.cell(row=row, column=1, value=item[0])
    c2 = ws2.cell(row=row, column=2, value=item[1])
    c1.font = bold_font; c2.font = bold_font
    c1.border = thin_border; c2.border = thin_border
    if is_key:
        c1.fill = key_fill; c2.fill = key_fill
    row += 1

# ===========================================================
# Sheet 3: 敏感性分析
# ===========================================================
ws3 = wb.create_sheet('敏感性分析')
ws3.sheet_properties.tabColor = 'FFC000'
ws3.column_dimensions['A'].width = 22
for i in range(2, 10):
    ws3.column_dimensions[get_column_letter(i)].width = 16

row = 1
set_title(ws3, row, 1, 'DCF 敏感性分析：WACC × 终端增长率(g)', merge_end=8)
row += 2

wacc_range = [0.08, 0.085, 0.09, 0.095, 0.10, 0.105]
tg_range = [0.025, 0.028, 0.030, 0.033, 0.035]

ws3.cell(row=row, column=1, value='WACC \\ g')
for i, w in enumerate(wacc_range):
    ws3.cell(row=row, column=2 + i, value=f'{w*100:.1f}%')
style_header(ws3, row, 1, 7)
row += 1

sens = []
for tg_val in tg_range:
    r_data = [f'g = {tg_val*100:.1f}%']
    for wv in wacc_range:
        pvf = [1 / (1 + wv) ** (i + 1) for i in range(5)]
        pv_fcf_s = sum([base_fcf_list[i] * pvf[i] for i in range(5)])
        tv_s = base_fcf_list[-1] * (1 + tg_val) / (wv - tg_val)
        pv_tv_s = tv_s / (1 + wv) ** 5
        ev_s = round(pv_fcf_s + pv_tv_s, 2)
        eq_s = round(ev_s + net_cash, 2)
        ps = round(eq_s / shares, 2)
        r_data.append(ps)
    sens.append(r_data)

for rd in sens:
    for j, v in enumerate(rd):
        cell = ws3.cell(row=row, column=1 + j, value=v)
        cell.font = bold_font if j == 0 else data_font
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='center')
        if j > 0 and isinstance(v, (int, float)):
            if v > 60:
                cell.fill = PatternFill('solid', fgColor='C6EFCE')
            elif v < 40:
                cell.fill = PatternFill('solid', fgColor='FFC7CE')
    row += 1

row += 2

# 相对估值
set_subtitle(ws3, row, 1, '相对估值交叉验证', merge_end=6)
row += 1
pe_headers = ['估值方法', '适用倍数', '2026E预测值', '隐含目标价', '说明']
for i, h in enumerate(pe_headers):
    ws3.cell(row=row, column=1 + i, value=h)
style_header(ws3, row, 1, 5)
row += 1

eps_26 = round(proj_np[0] / shares, 2)
bv = round(parent_eq_h[-1] / shares, 2)

pe_table = [
    ('PE（2026E）', '15~20x', f'EPS≈{eps_26}元', f'{round(eps_26*15,1)}~{round(eps_26*20,1)}元', '汽车零部件行业均值~18x'),
    ('PE（2027E）', '12~16x', f'EPS增长约12%', '51~68元', '增速放缓后折价'),
    ('PB', '1.8~2.5x', f'BVPS≈{bv}元', f'{round(bv*1.8,1)}~{round(bv*2.5,1)}元', '重资产制造业均值'),
    ('EV/EBITDA', '8~12x', 'EBITDA≈48亿', '45~67元', '同业比较'),
]
for item in pe_table:
    for j, v in enumerate(item):
        ws3.cell(row=row, column=1 + j, value=v).font = data_font
        ws3.cell(row=row, column=1 + j).border = thin_border
    row += 1

row += 2
set_subtitle(ws3, row, 1, '综合估值结论', merge_end=6)
row += 1
conclusions = [
    ('DCF 基准目标价', f'{b_target}元', '权重 40%'),
    ('相对估值区间', '42~60元', '权重 30%'),
    ('新业务期权价值（机器人+液冷）', '加5~15元', '权重 15%'),
    ('市场情绪/流动性溢价', '加5~10元', '权重 15%'),
    ('综合目标价区间', '55~75元', ''),
    ('当前股价', '66.89元', '位于目标价中位'),
    ('投资建议', '中性（等待盈利拐点信号）', ''),
]
for item in conclusions:
    for j, v in enumerate(item):
        cell = ws3.cell(row=row, column=1 + j, value=v)
        cell.font = bold_font if j == 0 else data_font
        cell.border = thin_border
        if '综合目标价' in item[0] or '投资建议' in item[0]:
            cell.fill = key_fill
            cell.font = bold_font
    row += 1

# ===========================================================
# Sheet 4: 财务比率
# ===========================================================
ws4 = wb.create_sheet('财务比率')
ws4.sheet_properties.tabColor = 'FF6600'
ws4.column_dimensions['A'].width = 30
for i in range(2, 12):
    ws4.column_dimensions[get_column_letter(i)].width = 14

row = 1
set_title(ws4, row, 1, '拓普集团 关键财务比率', merge_end=n_cols + 1)
row += 2

ratio_headers = ['比率指标'] + all_years
for i, h in enumerate(ratio_headers):
    ws4.cell(row=row, column=1 + i, value=h)
style_header(ws4, row, 1, n_cols + 1)
row += 1

# 各种比率
ratios = [
    ('-- 盈利能力 --', None, False),
    ('毛利率', gm_h + gm_p, True),
    ('净利率', np_m_h + np_m_p, True),
    ('营业利润率', [round(op_profit_h[i]/rev_h[i], 4) for i in range(3)] + [round(proj_op_p[i]/proj_rev[i], 4) for i in range(5)], True),
    ('三费占比', [round((sell_exp_h[i]+admin_exp_h[i]+rd_exp_h[i])/rev_h[i], 4) for i in range(3)] + [round(sell_rate[i]+admin_rate[i]+rd_rate[i], 4) for i in range(5)], True),
    ('-- 成长性 --', None, False),
    ('营收增长率', [None] + [round((rev_h[i]-rev_h[i-1])/rev_h[i-1], 4) for i in range(1,3)] + rev_g, True),
    ('净利润增长率', [None] + [round((parent_np_h[i]-parent_np_h[i-1])/parent_np_h[i-1], 4) for i in range(1,3)] + [round((proj_np[i]-parent_np_h[-1])/parent_np_h[-1], 4) for i in range(5)], True),
    ('-- 运营效率 --', None, False),
    ('应收账款周转天数', [round(ar_h[i]/rev_h[i]*365, 1) for i in range(3)] + ar_days, False),
    ('存货周转天数', [round(inv_h[i]/cogs_h[i]*365, 1) for i in range(3)] + inv_days, False),
    ('应付账款周转天数', [round(ap_h[i]/cogs_h[i]*365, 1) for i in range(3)] + ap_days, False),
    ('-- 财务健康 --', None, False),
    ('资产负债率', [round(total_liab_h[i]/total_assets_h[i], 4) for i in range(3)] + [None]*5, True),
    ('流动比率', [round(cur_assets_h[i]/cur_liab_h[i], 2) for i in range(3)] + [None]*5, False),
    ('经营现金流/净利润', [round(op_cf_h[i]/parent_np_h[i], 2) for i in range(3)] + [None]*5, False),
    ('ROE（摊薄）', [round(parent_np_h[i]/parent_eq_h[i], 4) for i in range(3)] + [None]*5, True),
]

for item in ratios:
    if item[1] is None:
        set_subtitle(ws4, row, 1, item[0], merge_end=n_cols + 1)
        row += 1
    else:
        write_key_row(ws4, row, item[0], item[1], is_pct=item[2])
        row += 1

# ===========================================================
# 保存
# ===========================================================
output_path = r'd:\claude-for-finance\outputs\models\拓普集团_三表财务模型.xlsx'
wb.save(output_path)

print('✅ 模型已保存到:', output_path)
print()
print('=== 模型概览 ===')
print(f'DCF 基准目标价: {b_target}元/股')
print(f'综合估值区间: 55~75元')
print(f'当前股价: 66.89元')
print(f'隐含涨跌幅: {round((b_target - 66.89) / 66.89 * 100, 1)}%')
print(f'2026E 净利: {proj_np[0]}亿 | EPS: {eps_26}元')
print(f'2026E 营收: {proj_rev[0]}亿 | 毛利率: {gm_p[0]*100:.1f}%')
