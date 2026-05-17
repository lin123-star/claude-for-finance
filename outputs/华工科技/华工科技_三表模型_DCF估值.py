"""
华工科技 (000988.SZ) 三表模型 + DCF 估值
基于 2025 年年报和 2026 年一季报数据
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# 样式
HDR = PatternFill("solid", fgColor="1F4E79")
SUB = PatternFill("solid", fgColor="D9E1F2")
OUT = PatternFill("solid", fgColor="BDD7EE")
INP = PatternFill("solid", fgColor="F2F2F2")
YEL = PatternFill("solid", fgColor="FFF2CC")

HF = Font(bold=True, color="FFFFFF", size=10, name="Arial")
BF = Font(bold=True, size=10, name="Arial")
NF = Font(size=10, name="Arial")
BL = Font(bold=True, size=10, name="Arial", color="0000FF")
TF = Font(bold=True, size=14, name="Arial")

thin = Border(*(Side('thin'),)*4)
ca = Alignment(horizontal='center', vertical='center')

def sc(ws, r, c, v, fn=NF, fl=None, nf_=None):
    cl = ws.cell(row=r, column=c, value=v)
    cl.font = fn
    if fl: cl.fill = fl
    if nf_: cl.number_format = nf_
    cl.border = thin
    cl.alignment = ca

def hdr(ws, r, vs):
    for i, v in enumerate(vs):
        sc(ws, r, 1+i, v, HF, HDR)

# ============ Sheet 1: 假设 ============
ws1 = wb.active
ws1.title = "假设与输入"
sc(ws1, 1, 1, "华工科技 (000988.SZ) 三表模型 & DCF 估值", TF)
sc(ws1, 2, 1, "数据截止: 2026-05-10 | 基于2025年报及2026Q1", BF)

data = [
    ("市场数据", 1),
    ("最新股价 (元)", 136.69),
    ("总股本 (亿股)", 10.05),
    ("总市值 (亿元)", 1374),
    ("WACC", 0.09),
    ("永续增长率", 0.03),
    ("税率", 0.15),
    ("", None),
    ("增长率假设", 1),
    ("2026E 营收增速", 0.30),
    ("2027E 营收增速", 0.25),
    ("2028E 营收增速", 0.20),
    ("2029E 营收增速", 0.15),
    ("2030E 营收增速", 0.10),
    ("", None),
    ("EBIT利润率假设", 1),
    ("2026E", 0.09),
    ("2027E", 0.10),
    ("2028E", 0.105),
    ("2029E", 0.11),
    ("2030E", 0.11),
    ("", None),
    ("光模块同业 PE", 1),
    ("中际旭创 2026E PE", 38),
    ("新易盛 2026E PE", 28),
    ("天孚通信 2026E PE", 35),
]

r = 3
for label, val in data:
    r += 1
    if val == 1:
        sc(ws1, r, 1, label, HF, HDR)
    elif label:
        sc(ws1, r, 1, label, BF)
        nf = '0.0%' if isinstance(val, float) and val < 1 else '#,##0.00'
        sc(ws1, r, 2, val, BL, INP, nf)

# ============ Sheet 2: 利润表 ============
ws2 = wb.create_sheet("利润表")
sc(ws2, 1, 1, "华工科技 利润表 (百万元)", TF)

rev_hist = [10309, 11705, 14360]
ni_hist = [1007, 1217, 1470]
years = ["2023A", "2024A", "2025A", "2026E", "2027E", "2028E", "2029E", "2030E"]

hdr(ws2, 3, ["项目"] + years)

# 预测
gr = [0.30, 0.25, 0.20, 0.15, 0.10]
em = [0.09, 0.10, 0.105, 0.11, 0.11]
rev = rev_hist + [round(rev_hist[2]*1.3, 1)]
for i in range(1,5):
    rev.append(round(rev[-1]*(1+gr[i]), 1))
ebit = [round(rev_hist[i]*0.085, 1) for i in range(3)] + [round(rev[3+i]*em[i], 1) for i in range(5)]
ni = [round(rev_hist[i]*0.10, 1) for i in range(3)] + [round(ebit[3+i]*0.85, 1) for i in range(5)]

rows = [
    ("营业总收入", rev, '#,##0'),
    ("营收增速", [f"={get_column_letter(c+2)}5/{get_column_letter(c+1)}5-1" if c > 0 else "-" for c in range(8)], '0.0%'),
    ("营业利润 (EBIT)", ebit, '#,##0'),
    ("EBIT利润率", [f"={get_column_letter(c+2)}7/{get_column_letter(c+2)}5" for c in range(8)], '0.0%'),
    ("所得税 (15%)", [round(v*0.15,1) for v in ebit], '#,##0'),
    ("归母净利润", ni, '#,##0'),
]

for idx, (label, vals, fmt) in enumerate(rows):
    row = 5 + idx
    sc(ws2, row, 1, label, BF)
    for c in range(8):
        if isinstance(vals[c], str):
            sc(ws2, row, 2+c, vals[c], NF, OUT)
        else:
            fl = OUT if c >= 3 else None
            sc(ws2, row, 2+c, vals[c], (BL if c >= 3 else NF), fl, fmt)

# ============ Sheet 3: DCF ============
ws3 = wb.create_sheet("DCF估值")
sc(ws3, 1, 1, "华工科技 DCF 估值模型", TF)

fcst_rev = rev[3:]
fcst_ebit = ebit[3:]
fcst_ni = ni[3:]
da_rate = [0.035, 0.035, 0.03, 0.03, 0.03]
capex_rate = [0.08, 0.07, 0.06, 0.06, 0.05]
wc_rate = [0.05, 0.04, 0.03, 0.03, 0.02]

nopat = [round(fcst_ebit[i]*0.85, 1) for i in range(5)]
da = [round(fcst_rev[i]*da_rate[i], 1) for i in range(5)]
capex = [round(fcst_rev[i]*capex_rate[i], 1) for i in range(5)]
wc = [round(fcst_rev[i]*wc_rate[i], 1) for i in range(5)]
ufcf = [round(nopat[i]+da[i]-capex[i]-wc[i], 1) for i in range(5)]

hdr(ws3, 3, ["项目"] + years[3:])
dcf_rows = [
    ("NOPAT", nopat),
    ("+ 折旧摊销", da),
    ("- 资本开支", capex),
    ("- 营运资本变动", wc),
    ("= UFCF", ufcf),
]
for idx, (label, vals) in enumerate(dcf_rows):
    row = 4+idx
    sc(ws3, row, 1, label, BF)
    for c in range(5):
        sc(ws3, row, 2+c, vals[c], NF, nf_='#,##0')

wacc = 0.09; tg = 0.03
sc(ws3, 10, 1, "折现因子", BF)
for i in range(5):
    sc(ws3, 10, 2+i, round(1/(1+wacc)**(i+1),4), NF, nf_='0.0000')
sc(ws3, 11, 1, "UFCF 现值", BF)
pv = sum(ufcf[i]/(1+wacc)**(i+1) for i in range(5))
for i in range(5):
    sc(ws3, 11, 2+i, round(ufcf[i]/(1+wacc)**(i+1),1), NF, nf_='#,##0')

tv = round(ufcf[4]*(1+tg)/(wacc-tg), 1)
pv_tv = round(tv/(1+wacc)**5, 1)
ev = round(pv + pv_tv, 1)
cash = 45*100; debt = 30*100
eq = round(ev + cash - debt, 1)
ps = round(eq/1005, 2)

r = 14
for lbl, val in [
    ("预测期 FCF 现值 (百万)", round(pv,1)),
    ("终值现值 (百万)", pv_tv),
    ("企业价值 EV (百万)", ev),
    ("+ 现金 (百万)", cash),
    ("- 有息负债 (百万)", debt),
    ("股权价值 (百万)", eq),
    ("每股内在价值 (元)", ps),
    ("当前股价 (元)", 136.69),
    ("上行空间", round(ps/136.69-1, 3)),
]:
    sc(ws3, r, 1, lbl, BF)
    sc(ws3, r, 2, val, BL, OUT, ('#,##0' if abs(val) > 99 else '#,##0.00') if not isinstance(val, float) or val > 10 else '0.0%')
    r += 1

# 敏感性分析
r += 1
sc(ws3, r, 1, "敏感性分析: WACC vs 永续增长率", BF)
r += 1
wv = [0.08, 0.085, 0.09, 0.095, 0.10]
gv = [0.02, 0.025, 0.03, 0.035, 0.04]
sc(ws3, r, 1, "WACC \\ TG", BF, HDR)
for j, g in enumerate(gv):
    sc(ws3, r, 2+j, g, HF, HDR, '0.0%')
r += 1
for w in wv:
    sc(ws3, r, 1, w, BF, HDR, '0.0%')
    for g in gv:
        tv_s = ufcf[4]*(1+g)/(w-g)
        pv_tv_s = tv_s/(1+w)**5
        pv_s = sum(ufcf[i]/(1+w)**(i+1) for i in range(5))
        ev_s = pv_s + pv_tv_s + cash - debt
        sc(ws3, r, 2+gv.index(g)+1, round(ev_s/1005, 2), NF, nf_='#,##0')
    r += 1

# ============ Sheet 4: 相对估值 ============
ws4 = wb.create_sheet("相对估值")
sc(ws4, 1, 1, "华工科技 相对估值", TF)

hdr(ws4, 3, ["公司", "2025营收(亿)", "2025净利(亿)", "毛利率", "PE"])
comps = [
    ("华工科技", 143.6, 14.7, 0.20, 50.5),
    ("中际旭创", 328.0, 82.0, 0.426, 38),
    ("新易盛", 147.0, 46.0, 0.478, 28),
    ("天孚通信", 52.0, 17.0, 0.52, 35),
    ("大族激光", 180.0, 18.0, 0.32, 25),
]
for idx, (n, rv, ni, gm, pe) in enumerate(comps):
    row = 4+idx
    sc(ws4, row, 1, n, BF)
    sc(ws4, row, 2, rv, NF, nf_='#,##0.0')
    sc(ws4, row, 3, ni, NF, nf_='#,##0.0')
    sc(ws4, row, 4, gm, NF, nf_='0.0%')
    sc(ws4, row, 5, pe, NF, nf_='#,##0')

# 保存
import os
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "华工科技_三表模型_DCF估值.xlsx")
wb.save(out)
print(f"Excel saved: {out}")
