#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
东山精密 (002384.SZ) 三表模型 + DCF 估值
基于2025年年报及2026年一季报数据
"""

import os

# ========== 1. 参数设定 ==========
COMPANY = "东山精密"
OUTPUT_DIR = r"D:\claude-for-finance\outputs\东山精密"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, f"{COMPANY}_三表模型_DCF估值.xlsx")

# 确保目录存在
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ========== 2. 历史财务数据 ==========
# 利润表核心数据 (亿元)
revenue = {
    2020: 280.93, 2021: 317.93, 2022: 315.80,
    2023: 336.51, 2024: 367.70, 2025: 401.25
}
net_profit = {
    2020: 15.30, 2021: 18.62, 2022: 23.68,
    2023: 19.65, 2024: 10.86, 2025: 13.86
}
operating_profit = {
    2020: 17.23, 2021: 20.81, 2022: 26.31,
    2023: 21.89, 2024: 12.45, 2025: 15.98
}
gross_profit = {
    2020: 48.24, 2021: 52.87, 2022: 50.97,
    2023: 47.72, 2024: 46.70, 2025: 56.54
}
gross_margin = {k: v / revenue[k] * 100 for k, v in gross_profit.items()}
net_margin = {k: v / revenue[k] * 100 for k, v in net_profit.items()}

# 资产负债表 (2025年末, 亿元)
total_assets_2025 = 602.51
total_liabilities_2025 = 385.43  # 推算: assets * 63.98%
equity_2025 = 217.08  # approx
cash_2025 = 45.23  # 估算
debt_2025 = 178.75  # 有息负债
receivables_2025 = 97.93
inventory_2025 = 89.29
goodwill_2025 = 47.69

# 现金流 (2025年, 亿元)
operating_cf_2025 = 53.07
investing_cf_2025 = -82.83
financing_cf_2025 = 37.49
capex_2025 = 35.0  # 估算

# ========== 3. 计算关键财务比率 ==========
ratios = {
    "毛利率_2025": gross_margin[2025],
    "净利率_2025": net_margin[2025],
    "资产负债率": 63.98,
    "经营现金流净额": operating_cf_2025,
    "净现比": operating_cf_2025 / net_profit[2025] * 100 if net_profit[2025] > 0 else 0,
    "营收CAGR_3Y": (revenue[2025] / revenue[2022]) ** (1/3) - 1,
    "营收CAGR_5Y": (revenue[2025] / revenue[2020]) ** (1/5) - 1,
}

ratios_avg_revenue_growth = (revenue[2025] / revenue[2024] - 1) * 100

# ========== 4. DCF 估值假设 ==========
# 基于2025年并表后数据 + 2026Q1爆发式增长
# 预测期间: 2026E-2030E (5年显式预测期)
# 终值增长率: 3.0%
# WACC: 9.0% (考虑到较高负债率和行业风险)

# 2026年营收预测基准: 2026Q1实际131.38亿 * 4 = 525.52亿, 但考虑季节性
# 券商一致预期2026营收约550-650亿
# 取保守中性预测

assumptions = {
    "WACC": 0.09,
    "terminal_growth": 0.03,
    "tax_rate": 0.15,
    "debt_total": debt_2025,
    "cash_total": cash_2025,
    "shares_outstanding": 18.32,  # 亿股
    "current_stock_price": 211.39,
}

# DCF预测 (单位: 亿元)
dcf_years = [2026, 2027, 2028, 2029, 2030]

# 情景1: 保守情景 (基于PCB主业稳健增长, 光模块逐步放量)
conservative = {
    "revenue_growth": [0.35, 0.25, 0.18, 0.12, 0.10],
    "ebitda_margin": [0.09, 0.11, 0.13, 0.14, 0.15],
    "capex_pct_revenue": [0.10, 0.09, 0.08, 0.07, 0.07],
    "depreciation_pct_revenue": [0.055, 0.055, 0.05, 0.05, 0.045],
}

# 情景2: 基准情景 (光模块爆发 + AI PCB放量)
base = {
    "revenue_growth": [0.45, 0.30, 0.22, 0.15, 0.12],
    "ebitda_margin": [0.10, 0.12, 0.14, 0.16, 0.17],
    "capex_pct_revenue": [0.12, 0.10, 0.08, 0.07, 0.06],
    "depreciation_pct_revenue": [0.05, 0.05, 0.045, 0.04, 0.04],
}

# 情景3: 乐观情景 (AI算力全面爆发 + 光芯片外售)
optimistic = {
    "revenue_growth": [0.55, 0.38, 0.28, 0.20, 0.15],
    "ebitda_margin": [0.12, 0.14, 0.16, 0.18, 0.19],
    "capex_pct_revenue": [0.14, 0.12, 0.09, 0.07, 0.06],
    "depreciation_pct_revenue": [0.05, 0.045, 0.04, 0.035, 0.03],
}


def compute_dcf(scenario, label):
    """计算DCF估值"""
    results = {}
    prev_revenue = revenue[2025]
    pv_fcf = 0
    discount_factor = 1.0

    for i, yr in enumerate(dcf_years):
        rev = prev_revenue * (1 + scenario["revenue_growth"][i])
        ebitda = rev * scenario["ebitda_margin"][i]
        depr = rev * scenario["depreciation_pct_revenue"][i]
        ebit = ebitda - depr
        tax = ebit * assumptions["tax_rate"]
        nopat = ebit - tax
        capex = rev * scenario["capex_pct_revenue"][i]
        delta_wc = rev * 0.02  # 营运资本增加简化假设

        fcf = nopat + depr - capex - delta_wc

        discount_factor *= (1 + assumptions["WACC"])
        pv_fcf_i = fcf / discount_factor

        results[yr] = {
            "revenue": round(rev, 2),
            "ebitda": round(ebitda, 2),
            "ebitda_margin": round(scenario["ebitda_margin"][i] * 100, 1),
            "depreciation": round(depr, 2),
            "ebit": round(ebit, 2),
            "nopat": round(nopat, 2),
            "capex": round(capex, 2),
            "fcf": round(fcf, 2),
            "pv_fcf": round(pv_fcf_i, 2),
        }
        pv_fcf += pv_fcf_i
        prev_revenue = rev

    # 终值计算
    terminal_fcf = results[dcf_years[-1]]["fcf"] * (1 + assumptions["terminal_growth"])
    terminal_value = terminal_fcf / (assumptions["WACC"] - assumptions["terminal_growth"])
    pv_terminal = terminal_value / discount_factor

    # 企业价值
    ev = pv_fcf + pv_terminal
    equity_value = ev - assumptions["debt_total"] + assumptions["cash_total"]
    per_share = equity_value / assumptions["shares_outstanding"]

    return {
        "label": label,
        "pv_fcf_sum": round(pv_fcf, 2),
        "pv_terminal": round(pv_terminal, 2),
        "terminal_value": round(terminal_value, 2),
        "ev": round(ev, 2),
        "equity_value": round(equity_value, 2),
        "per_share": round(per_share, 2),
        "yearly": results,
    }


dcf_conservative = compute_dcf(conservative, "保守情景")
dcf_base = compute_dcf(base, "基准情景")
dcf_optimistic = compute_dcf(optimistic, "乐观情景")

# ========== 5. 相对估值 ==========
# 可比公司 (2026E基于一致预期)
comps = {
    "鹏鼎控股": {"pe_2025": 15.0, "pe_2026e": 18.0, "peg": 1.2, "ev_ebitda": 10.0},
    "胜宏科技": {"pe_2025": 25.0, "pe_2026e": 22.0, "peg": 0.8, "ev_ebitda": 15.0},
    "沪电股份": {"pe_2025": 28.0, "pe_2026e": 25.0, "peg": 0.9, "ev_ebitda": 18.0},
    "深南电路": {"pe_2025": 30.0, "pe_2026e": 28.0, "peg": 1.0, "ev_ebitda": 20.0},
}

# 东山精密当前估值
current_pe_ttm = 189.79
current_pb = 17.10
current_ev_ebitda = 35.0  # 估算

# ========== 6. 生成Excel ==========
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, numbers
    from openpyxl.utils import get_column_letter

    wb = Workbook()

    # ---- Sheet 1: 三表模型 ----
    ws1 = wb.active
    ws1.title = "三表模型"

    # 样式定义
    header_font = Font(name='微软雅黑', bold=True, size=11)
    title_font = Font(name='微软雅黑', bold=True, size=14)
    data_font = Font(name='微软雅黑', size=10)
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    header_font_white = Font(name='微软雅黑', bold=True, size=10, color='FFFFFF')
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    row = 1
    ws1.cell(row=row, column=1, value=f"{COMPANY} 三表模型").font = title_font
    row = 2
    ws1.cell(row=row, column=1, value=f"数据截至: 2025年年报").font = Font(name='微软雅黑', size=9, italic=True)

    # 利润表
    row = 4
    headers = ["项目", "2020", "2021", "2022", "2023", "2024", "2025"]
    for col, h in enumerate(headers, 1):
        cell = ws1.cell(row=row, column=col, value=h)
        cell.font = header_font_white
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='center')

    pl_items = [
        ("营业总收入(亿元)", list(revenue.values())),
        ("营业成本(亿元)", [revenue[k] - gross_profit[k] for k in revenue]),
        ("毛利润(亿元)", list(gross_profit.values())),
        ("毛利率(%)", [round(gross_margin[k], 2) for k in gross_margin]),
        ("营业利润(亿元)", list(operating_profit.values())),
        ("归母净利润(亿元)", list(net_profit.values())),
        ("净利率(%)", [round(net_margin[k], 2) for k in net_margin]),
        ("营收增速(%)", [0] + [round((revenue[k] - revenue[k-1]) / revenue[k-1] * 100, 2) for k in list(revenue.keys())[1:]]),
        ("净利润增速(%)", [0] + [round((net_profit[k] - net_profit[k-1]) / net_profit[k-1] * 100, 2) for k in list(net_profit.keys())[1:]]),
    ]

    years_list = list(revenue.keys())
    for item_name, item_data in pl_items:
        row += 1
        cell = ws1.cell(row=row, column=1, value=item_name)
        cell.font = Font(name='微软雅黑', bold=True, size=10)
        cell.border = thin_border
        for j, val in enumerate(item_data):
            cell = ws1.cell(row=row, column=j+2, value=val)
            cell.font = data_font
            cell.border = thin_border
            if isinstance(val, float):
                cell.number_format = '#,##0.00'

    # 杜邦分解
    row += 2
    ws1.cell(row=row, column=1, value="杜邦分解 (2025年)").font = Font(name='微软雅黑', bold=True, size=11)
    row += 1
    dupont_items = [
        ("净利率(%)", round(net_margin[2025], 2)),
        ("资产周转率(次)", round(revenue[2025] / total_assets_2025, 4)),
        ("权益乘数", round(total_assets_2025 / equity_2025, 4)),
        ("ROE(%)", round(net_margin[2025]/100 * (revenue[2025]/total_assets_2025) * (total_assets_2025/equity_2025) * 100, 2)),
    ]
    for item_name, val in dupont_items:
        ws1.cell(row=row, column=1, value=item_name).font = data_font
        ws1.cell(row=row, column=2, value=val).font = data_font
        row += 1

    # 资产负债表概要
    row += 2
    ws1.cell(row=row, column=1, value="资产负债表概要 (2025年末, 亿元)").font = Font(name='微软雅黑', bold=True, size=11)
    row += 1
    bs_items = [
        ("总资产", total_assets_2025),
        ("总负债", total_liabilities_2025),
        ("归母净资产", equity_2025),
        ("货币资金", cash_2025),
        ("有息负债", debt_2025),
        ("应收账款", receivables_2025),
        ("存货", inventory_2025),
        ("商誉", goodwill_2025),
        ("资产负债率(%)", 63.98),
    ]
    for item_name, val in bs_items:
        ws1.cell(row=row, column=1, value=item_name).font = data_font
        ws1.cell(row=row, column=2, value=val).font = data_font
        row += 1

    # 现金流量表
    row += 2
    ws1.cell(row=row, column=1, value="现金流量表 (2025年, 亿元)").font = Font(name='微软雅黑', bold=True, size=11)
    row += 1
    cf_items = [
        ("经营活动现金流", operating_cf_2025),
        ("投资活动现金流", investing_cf_2025),
        ("筹资活动现金流", financing_cf_2025),
        ("净现比(%)", round(operating_cf_2025 / net_profit[2025] * 100, 2)),
    ]
    for item_name, val in cf_items:
        ws1.cell(row=row, column=1, value=item_name).font = data_font
        ws1.cell(row=row, column=2, value=val).font = data_font
        row += 1

    # ---- Sheet 2: DCF估值 ----
    ws2 = wb.create_sheet("DCF估值")

    ws2.cell(row=1, column=1, value=f"{COMPANY} DCF估值模型").font = title_font
    ws2.cell(row=2, column=1, value="核心假设").font = Font(name='微软雅黑', bold=True, size=12)

    row = 3
    key_assumptions = [
        ("WACC", f"{assumptions['WACC']*100:.1f}%"),
        ("终值增长率", f"{assumptions['terminal_growth']*100:.1f}%"),
        ("税率", f"{assumptions['tax_rate']*100:.0f}%"),
        ("总股本(亿股)", assumptions["shares_outstanding"]),
        ("净债务(亿元)", round(assumptions["debt_total"] - assumptions["cash_total"], 2)),
        ("当前股价(元)", assumptions["current_stock_price"]),
    ]
    for item_name, val in key_assumptions:
        ws2.cell(row=row, column=1, value=item_name).font = data_font
        ws2.cell(row=row, column=2, value=val).font = data_font
        row += 1

    # 三个情景
    for scenario_label, dcf_result in [("保守情景", dcf_conservative), ("基准情景", dcf_base), ("乐观情景", dcf_optimistic)]:
        row += 2
        ws2.cell(row=row, column=1, value=f"{scenario_label} DCF明细").font = Font(name='微软雅黑', bold=True, size=11, color='4472C4')
        row += 1

        dcf_headers = ["年份", "营收(亿)", "EBITDA(亿)", "EBITDA%", "折旧(亿)", "EBIT(亿)", "NOPAT(亿)", "CAPEX(亿)", "FCF(亿)", "PV of FCF(亿)"]
        for col, h in enumerate(dcf_headers, 1):
            cell = ws2.cell(row=row, column=col, value=h)
            cell.font = header_font_white
            cell.fill = PatternFill(start_color='5B9BD5', end_color='5B9BD5', fill_type='solid')
            cell.border = thin_border

        for yr in dcf_years:
            row += 1
            d = dcf_result["yearly"][yr]
            vals = [yr, d["revenue"], d["ebitda"], d["ebitda_margin"], d["depreciation"],
                    d["ebit"], d["nopat"], d["capex"], d["fcf"], d["pv_fcf"]]
            for col, val in enumerate(vals, 1):
                cell = ws2.cell(row=row, column=col, value=val)
                cell.font = data_font
                cell.border = thin_border

        row += 1
        summary_items = [
            ("显式期PV of FCF合计", dcf_result["pv_fcf_sum"]),
            ("终值", dcf_result["terminal_value"]),
            ("终值PV", dcf_result["pv_terminal"]),
            ("企业价值(EV)", dcf_result["ev"]),
            ("(+)现金", assumptions["cash_total"]),
            ("(-)债务", assumptions["debt_total"]),
            ("股权价值(亿元)", dcf_result["equity_value"]),
            ("每股价值(元)", dcf_result["per_share"]),
            ("当前股价(元)", assumptions["current_stock_price"]),
            ("隐含涨跌幅(%)", round((dcf_result["per_share"] / assumptions["current_stock_price"] - 1) * 100, 2)),
        ]
        for item_name, val in summary_items:
            row += 1
            cell = ws2.cell(row=row, column=1, value=item_name)
            cell.font = Font(name='微软雅黑', bold=True, size=10)
            cell = ws2.cell(row=row, column=2, value=val)
            cell.font = data_font

    # ---- Sheet 3: 相对估值 ----
    ws3 = wb.create_sheet("相对估值")

    ws3.cell(row=1, column=1, value="相对估值 - 可比公司分析").font = title_font

    row = 3
    comp_headers = ["公司", "2025PE", "2026E PE", "PEG", "EV/EBITDA", "2025营收增速", "2025净利增速"]
    for col, h in enumerate(comp_headers, 1):
        cell = ws3.cell(row=row, column=col, value=h)
        cell.font = header_font_white
        cell.fill = header_fill
        cell.border = thin_border

    comp_data = [
        ["鹏鼎控股", 15.0, 18.0, 1.2, 10.0, 11.40, 3.25],
        ["胜宏科技", 25.0, 22.0, 0.8, 15.0, 79.77, 273.52],
        ["沪电股份", 28.0, 25.0, 0.9, 18.0, 42.00, 47.74],
        ["深南电路", 30.0, 28.0, 1.0, 20.0, 32.05, 74.47],
    ]

    for d in comp_data:
        row += 1
        for col, val in enumerate(d, 1):
            cell = ws3.cell(row=row, column=col, value=val)
            cell.font = data_font
            cell.border = thin_border

    # 东山精密对比
    row += 2
    ws3.cell(row=row, column=1, value="东山精密当前估值").font = Font(name='微软雅黑', bold=True, size=11, color='4472C4')
    row += 1
    ds_metrics = [
        ("PE(TTM)", current_pe_ttm),
        ("PB(市净率)", current_pb),
        ("PS(市销率TTM)", round(assumptions["current_stock_price"] * assumptions["shares_outstanding"] / revenue[2025], 2)),
        ("EV/EBITDA(估算)", current_ev_ebitda),
        ("总市值(亿元)", round(assumptions["current_stock_price"] * assumptions["shares_outstanding"], 0)),
        ("2026Q1净利润同比", "143.47%"),
    ]
    for item_name, val in ds_metrics:
        ws3.cell(row=row, column=1, value=item_name).font = data_font
        ws3.cell(row=row, column=2, value=val).font = data_font
        row += 1

    # 调整列宽
    for ws in [ws1, ws2, ws3]:
        for col in range(1, 15):
            ws.column_dimensions[get_column_letter(col)].width = 16

    wb.save(OUTPUT_FILE)
    print(f"Excel模型已保存至: {OUTPUT_FILE}")
    print(f"  - 三表模型 (利润表+杜邦+资产负债表+现金流量表)")
    print(f"  - DCF估值 (三情景: 保守/基准/乐观)")
    print(f"  - 相对估值 (可比公司分析)")

    # 输出DCF结果摘要
    print("\n========== DCF估值结果 ==========")
    for result in [dcf_conservative, dcf_base, dcf_optimistic]:
        print(f"\n{result['label']}:")
        print(f"  显式期PV of FCF: {result['pv_fcf_sum']}亿元")
        print(f"  终值PV: {result['pv_terminal']}亿元")
        print(f"  企业价值(EV): {result['ev']}亿元")
        print(f"  股权价值: {result['equity_value']}亿元")
        print(f"  每股价值: {result['per_share']}元")

except ImportError:
    print("警告: openpyxl未安装, 跳过Excel生成")
    print("尝试自动安装...")
    import subprocess
    subprocess.run(["python", "-m", "pip", "install", "openpyxl"], shell=True)
    print("请重新运行此脚本")
