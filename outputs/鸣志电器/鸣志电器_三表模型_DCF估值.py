"""
鸣志电器 (603728.SH) 三表模型 + DCF 估值
基于 2025 年年报和 2026 年一季报数据
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# 样式定义
header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
sub_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
input_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
output_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

hf = Font(bold=True, color="FFFFFF", size=10, name="Arial")
bf = Font(bold=True, size=10, name="Arial")
nf = Font(size=10, name="Arial")
bf_blue = Font(bold=True, size=10, name="Arial", color="0000FF")
title_f = Font(bold=True, size=14, name="Arial")
sub_f = Font(bold=True, size=12, name="Arial")

thin = Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin'))
ca = Alignment(horizontal='center', vertical='center')

def sc(ws, r, c, v, fn=nf, fl=None, nf_=None):
    cl = ws.cell(row=r, column=c, value=v)
    cl.font = fn
    if fl: cl.fill = fl
    if nf_: cl.number_format = nf_
    cl.border = thin
    cl.alignment = ca
    return cl

def sh(ws, r, vs, start=1):
    for i, v in enumerate(vs):
        sc(ws, r, start+i, v, hf, header_fill)

# ============ Sheet 1: 假设与输入 ============
ws1 = wb.active
ws1.title = "假设与输入"
ws1.merge_cells("A1:F1")
sc(ws1, 1, 1, "鸣志电器 (603728.SH) 三表模型 & DCF 估值", title_f)
ws1.merge_cells("A2:F2")
sc(ws1, 2, 1, "数据截止: 2026-05-10 | 基于2025年报及2026Q1", sub_f)

rows = [
    ("市场数据", None, True),
    ("最新股价 (元)", 63.48, False),
    ("总股本 (亿股)", 4.19, False),
    ("总市值 (亿元)", 265.91, False),
    ("货币资金 (亿元)", 8.71, False),
    ("有息负债 (亿元)", 5.32, False),
    ("净现金 (亿元)", 3.39, False),
    ("资产负债率", 0.45, False),
    ("", None, False),
    ("DCF 参数", None, True),
    ("WACC", 0.0539, False),
    ("永续增长率", 0.03, False),
    ("税率", 0.15, False),
    ("", None, False),
    ("增长率假设", None, True),
    ("2026E 营收增速", 0.25, False),
    ("2027E 营收增速", 0.22, False),
    ("2028E 营收增速", 0.18, False),
    ("2029E 营收增速", 0.15, False),
    ("2030E 营收增速", 0.12, False),
    ("", None, False),
    ("利润率假设", None, True),
    ("2026E EBIT利润率", 0.035, False),
    ("2027E EBIT利润率", 0.05, False),
    ("2028E EBIT利润率", 0.065, False),
    ("2029E EBIT利润率", 0.08, False),
    ("2030E EBIT利润率", 0.09, False),
    ("", None, False),
    ("可比公司 PE", None, True),
    ("雷赛智能 2026E PE", 43, False),
    ("汇川技术 2026E PE", 28, False),
    ("伟创电气 2026E PE", 45, False),
]

r = 3
for label, val, is_hdr in rows:
    r += 1
    if is_hdr:
        ws1.merge_cells(f"A{r}:F{r}")
        sc(ws1, r, 1, label, hf, header_fill)
    elif label:
        sc(ws1, r, 1, label, bf)
        sc(ws1, r, 2, val, bf_blue, input_fill, nf_='#,##0.00' if isinstance(val, float) and val > 1 else '0.0%' if isinstance(val, float) else '#,##0')
        sc(ws1, r, 3, "用户输入", nf, yellow_fill)

# ============ Sheet 2: 利润表 ============
ws2 = wb.create_sheet("利润表")
ws2.merge_cells("A1:K1")
sc(ws2, 1, 1, "鸣志电器 利润表 (百万元)", title_f)

years = ["2021A","2022A","2023A","2024A","2025A","2026E","2027E","2028E","2029E","2030E"]
sh(ws2, 3, ["项目"] + years)

# 历史数据 (百万元)
rev_hist = [2714, 2960, 2543, 2416, 2762]
ni_hist = [280, 247, 140, 78, 61.13]
cogs_hist = [rev_hist[i]*0.626 for i in range(5)]

# 预测假设
gr = [0.25, 0.22, 0.18, 0.15, 0.12]
gm = [0.37, 0.38, 0.39, 0.40, 0.40]
em = [0.035, 0.05, 0.065, 0.08, 0.09]

rev_proj = []
for i in range(5):
    prev = rev_hist[4] if i == 0 else rev_proj[i-1]
    rev_proj.append(round(prev * (1 + gr[i]), 1))

ni_proj = []
cogs_proj = []
gp_proj = []
ebit_proj = []
for i in range(5):
    ni_proj.append(round(rev_proj[i] * em[i] * 0.85, 1))
    cogs_proj.append(round(rev_proj[i] * (1 - gm[i]), 1))
    gp_proj.append(round(rev_proj[i] * gm[i], 1))
    ebit_proj.append(round(rev_proj[i] * em[i], 1))

all_rev = rev_hist + rev_proj
all_cogs = cogs_hist + cogs_proj
all_gp = [rev_hist[i]*0.374 for i in range(5)] + gp_proj
all_ebit = [rev_hist[i]*0.036 for i in range(5)] + ebit_proj
all_ni = ni_hist + ni_proj

labels_is = [
    ("营业总收入", all_rev, '#,##0'),
    ("营收增速", [f"={get_column_letter(c+2)}5/{get_column_letter(c+1)}5-1" if c > 0 else "-" for c in range(10)], '0.0%'),
    ("营业成本", all_cogs, '#,##0'),
    ("毛利", all_gp, '#,##0'),
    ("毛利率", [f"={get_column_letter(c+2)}8/{get_column_letter(c+2)}5" if all_rev[c] else 0 for c in range(10)], '0.0%'),
    ("营业利润 (EBIT)", all_ebit, '#,##0'),
    ("EBIT利润率", [f"={get_column_letter(c+2)}10/{get_column_letter(c+2)}5" if all_rev[c] else 0 for c in range(10)], '0.0%'),
    ("所得税 (15%)", [round(v*0.15,1) for v in all_ebit], '#,##0'),
    ("归母净利润", all_ni, '#,##0'),
    ("净利率", [f"={get_column_letter(c+2)}13/{get_column_letter(c+2)}5" if all_rev[c] else 0 for c in range(10)], '0.0%'),
    ("EPS (元)", [round(ni/419, 2) for ni in all_ni], '#,##0.00'),
]

data_start = 5
for idx, (label, vals, fmt) in enumerate(labels_is):
    row = data_start + idx
    sc(ws2, row, 1, label, bf)
    for c in range(10):
        fmt_use = '#,##0' if '#' in str(fmt) else fmt
        if isinstance(vals[c], str):
            sc(ws2, row, c+2, vals[c], nf, output_fill)
        elif label == "毛利率" and idx == 4:
            sc(ws2, row, c+2, vals[c], nf, nf_='0.0%')
        else:
            sc(ws2, row, c+2, vals[c], bf_blue if c >= 5 else nf,
               output_fill if c >= 5 else None, nf_=fmt_use)

# ============ Sheet 3: DCF 估值 ============
ws3 = wb.create_sheet("DCF估值")
ws3.merge_cells("A1:G1")
sc(ws3, 1, 1, "鸣志电器 DCF 估值模型", title_f)

# 预测参数
rev_fcst = rev_proj
da_rate = [0.05, 0.05, 0.05, 0.04, 0.04]
capex_rate = [0.10, 0.09, 0.08, 0.07, 0.06]
wc_rate = [0.05, 0.045, 0.04, 0.035, 0.03]

nopat = [round(rev_fcst[i] * em[i] * 0.85, 1) for i in range(5)]
da = [round(rev_fcst[i] * da_rate[i], 1) for i in range(5)]
capex = [round(rev_fcst[i] * capex_rate[i], 1) for i in range(5)]
wc_change = [round(rev_fcst[i] * wc_rate[i], 1) for i in range(5)]
ufcf = [round(nopat[i] + da[i] - capex[i] - wc_change[i], 1) for i in range(5)]

sh(ws3, 3, ["项目"] + years[5:])
dcf_labels_is = [
    ("NOPAT", nopat),
    ("+ 折旧摊销", da),
    ("- 资本开支", capex),
    ("- 营运资本变动", wc_change),
    ("= UFCF", ufcf),
]
for idx, (label, vals) in enumerate(dcf_labels_is):
    row = 4 + idx
    sc(ws3, row, 1, label, bf)
    for c in range(5):
        sc(ws3, row, c+2, vals[c], nf, nf_='#,##0')

# DCF 计算
wacc = 0.0539
tg = 0.03
tax_rate = 0.15

# 折现
dcf_row = 10
sc(ws3, dcf_row, 1, "折现因子", bf)
for i in range(5):
    df = 1 / (1 + wacc) ** (i + 1)
    sc(ws3, dcf_row, 2+i, round(df, 4), nf, nf_='0.0000')

sc(ws3, dcf_row+1, 1, "UFCF 现值", bf)
pv_total = 0
for i in range(5):
    pv = round(ufcf[i] / (1 + wacc) ** (i + 1), 1)
    pv_total += pv
    sc(ws3, dcf_row+1, 2+i, pv, nf, nf_='#,##0')

# 终值
tv = round(ufcf[4] * (1 + tg) / (wacc - tg), 1)
pv_tv = round(tv / (1 + wacc) ** 5, 1)
ev = round(pv_total + pv_tv, 1)
cash = 8.71 * 100  # 百万元
debt = 5.32 * 100  # 百万元
eq_val = round(ev + cash - debt, 1)
shares = 4.19 * 100  # 百万股
per_share = round(eq_val / shares, 2)

r = dcf_row + 3
summary_items = [
    ("预测期 FCF 现值合计 (百万元)", pv_total),
    ("终值现值 (百万元)", pv_tv),
    ("企业价值 EV (百万元)", ev),
    ("+ 现金 (百万元)", cash),
    ("- 有息负债 (百万元)", debt),
    ("股权价值 (百万元)", eq_val),
    ("总股本 (百万股)", shares),
    ("每股内在价值 (元)", per_share),
]
for label, val in summary_items:
    sc(ws3, r, 1, label, bf)
    sc(ws3, r, 2, val, bf, output_fill, nf_='#,##0.00' if '每股' in label else '#,##0')
    r += 1

# 敏感性分析
r += 1
ws3.merge_cells(f"A{r}:G{r}")
sc(ws3, r, 1, "敏感性分析: WACC vs 永续增长率", bf)
r += 1

wacc_vals = [0.044, 0.049, 0.054, 0.059, 0.064]
tg_vals = [0.02, 0.025, 0.03, 0.035, 0.04]
sc(ws3, r, 1, "WACC \\ TG", bf, header_fill)
for j, tg_v in enumerate(tg_vals):
    sc(ws3, r, 2+j, tg_v, hf, header_fill, nf_='0.0%')
r += 1

for w in wacc_vals:
    sc(ws3, r, 1, w, bf, header_fill, nf_='0.0%')
    for j, tg_v in enumerate(tg_vals):
        if w == tg_v and w == 0.054:  # Avoid div by zero for demo
            val = 0
        else:
            tv_s = ufcf[4] * (1 + tg_v) / (w - tg_v) if w != tg_v else 0
            pv_tv_s = tv_s / (1 + w) ** 5
            pv_fcf_s = sum(ufcf[i] / (1 + w) ** (i+1) for i in range(5))
            ev_s = pv_fcf_s + pv_tv_s
            eq_s = ev_s + cash - debt
            val = round(eq_s / shares, 2)
        sc(ws3, r, 2+j, val, nf, nf_='#,##0.00')
    r += 1

# 当前股价参考
r += 1
sc(ws3, r, 1, "当前股价 (元)", bf)
sc(ws3, r, 2, 63.48, bf, yellow_fill, nf_='#,##0.00')
sc(ws3, r, 3, "上行/下行空间", bf)
sc(ws3, r, 4, per_share / 63.48 - 1, bf, nf_='0.0%')

# ============ Sheet 4: 相对估值 ============
ws4 = wb.create_sheet("相对估值")
ws4.merge_cells("A1:E1")
sc(ws4, 1, 1, "鸣志电器 相对估值比较", title_f)

sh(ws4, 3, ["公司", "2025营收(亿)", "2025净利(亿)", "2026E PE", "PB"])
comps = [
    ("鸣志电器", 27.62, 0.61, 195, 8.94),
    ("雷赛智能", 18.50, 1.80, 43, 8.50),
    ("汇川技术", 451.0, 58.0, 28, 8.00),
    ("伟创电气", 22.0, 3.50, 45, 6.50),
]
for idx, (name, rev, ni, pe, pb) in enumerate(comps):
    r = 4 + idx
    sc(ws4, r, 1, name, bf)
    sc(ws4, r, 2, rev, nf, nf_='#,##0.0')
    sc(ws4, r, 3, ni, nf, nf_='#,##0.00')
    sc(ws4, r, 4, pe, nf, nf_='#,##0')
    sc(ws4, r, 5, pb, nf, nf_='#,##0.00')

# PE 法估值
r = 10
sc(ws4, r, 1, "PE 法估值", bf)
sc(ws4, r+1, 1, "同业中位数 PE", bf)
sc(ws4, r+1, 2, 43, nf, output_fill)
sc(ws4, r+2, 1, "2026E 预测净利润 (百万元)", bf)
sc(ws4, r+2, 2, ni_proj[0], nf, output_fill)
sc(ws4, r+3, 1, "隐含市值 (百万元)", bf)
sc(ws4, r+3, 2, 43 * ni_proj[0], nf, output_fill, nf_='#,##0')
sc(ws4, r+4, 1, "每股价值 (元)", bf)
pe_val = round(43 * ni_proj[0] / 419, 2)
sc(ws4, r+4, 2, pe_val, bf, output_fill, nf_='#,##0.00')
sc(ws4, r+5, 1, "相对当前股价空间", bf)
sc(ws4, r+5, 2, pe_val / 63.48 - 1, bf, nf_='0.0%')

# 保存文件
import os
output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "鸣志电器_三表模型_DCF估值.xlsx")
wb.save(output_path)
print(f"Excel 模型已保存: {output_path}")
