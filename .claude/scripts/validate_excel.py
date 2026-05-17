# -*- coding: utf-8 -*-
"""
Excel 财务模型验证脚本
用途：检查 AI 生成的三表模型和 DCF 估值 Excel 的公式错误、数据缺失、逻辑问题
用法：python validate_excel.py <excel_path> [--sheet <sheet_name>]
"""

import sys
import argparse
import openpyxl

# Excel 错误值
EXCEL_ERRORS = {"#DIV/0!", "#REF!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!"}

def scan_errors(ws, sheet_name):
    """扫描所有单元格中的 Excel 错误值"""
    errors = []
    for row in ws.iter_rows():
        for cell in row:
            if cell.value in EXCEL_ERRORS:
                errors.append({
                    "sheet": sheet_name,
                    "cell": cell.coordinate,
                    "error": cell.value,
                    "row_label": ws.cell(row=cell.row, column=1).value or ""
                })
    return errors


def scan_hardcoded_in_formula_area(ws, formula_start_row, formula_end_row, value_cols):
    """
    检查比率/派生区（非原始数据输入区）是否存在硬编码数值
    formula_start_row: 比率区起始行（含）
    formula_end_row: 比率区结束行（含），None 表示到末尾
    value_cols: 数据列范围，如 range(2, 8) 表示 B-G 列
    """
    hardcoded = []
    end = formula_end_row or ws.max_row
    for row in range(formula_start_row, end + 1):
        label = ws.cell(row=row, column=1).value or ""
        for col in value_cols:
            cell = ws.cell(row=row, column=col)
            v = cell.value
            if v is None or v == "":
                continue
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                hardcoded.append({
                    "cell": cell.coordinate,
                    "label": label,
                    "value": v,
                    "hint": "派生指标区发现数值，应为公式引用"
                })
    return hardcoded


def check_formula_references(ws, sheet_name):
    """检查公式中是否引用了不存在的行号（常见错误：行号偏移）"""
    suspicious = []
    max_row = ws.max_row
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                formula = cell.value
                # 检查引用是否超出实际行数
                import re
                refs = re.findall(r'[A-Z]+(\d+)', formula)
                for ref_row in refs:
                    if int(ref_row) > max_row + 5:  # 允许少量超出（预留行）
                        suspicious.append({
                            "cell": cell.coordinate,
                            "formula": formula,
                            "bad_ref_row": ref_row,
                            "max_row": max_row
                        })
    return suspicious


def check_dcf_logic(ws, sheet_name):
    """验证 DCF 特有的财务逻辑"""
    issues = []

    # 找 WACC 和终值增长率所在单元格 (通常在假设区)
    wacc_val = None
    g_val = None
    wacc_cell = None
    g_cell = None

    for row in ws.iter_rows(max_col=4):
        for cell in row:
            label = ws.cell(row=cell.row, column=1).value or ""
            if isinstance(label, str) and "WACC" in label.upper() and "/" not in label and wacc_val is None:
                val_cell = ws.cell(row=cell.row, column=2)
                if isinstance(val_cell.value, (int, float)):
                    wacc_val = val_cell.value
                    wacc_cell = val_cell.coordinate
            if isinstance(label, str) and ("终值增长" in label or "永续增长" in label or "terminal growth" in label.lower()) and "/" not in label and g_val is None:
                val_cell = ws.cell(row=cell.row, column=2)
                if isinstance(val_cell.value, (int, float)):
                    g_val = val_cell.value
                    g_cell = val_cell.coordinate

    if wacc_val is not None and g_val is not None:
        if g_val >= wacc_val:
            issues.append({
                "type": "DCF数学约束违反",
                "detail": f"终值增长率 g={g_val:.2%} ({g_cell}) >= WACC={wacc_val:.2%} ({wacc_cell})",
                "severity": "CRITICAL",
                "fix": "终值增长率必须 < WACC，否则 DCF 分母为负，终值趋向无穷大"
            })
        elif g_val > 0.05:
            issues.append({
                "type": "DCF假设偏高",
                "detail": f"终值增长率 g={g_val:.2%} > 5%，A股公司长期增速建议≤3%",
                "severity": "WARNING",
                "fix": "建议将永续增长率设为 2-3%"
            })

    # 检查 WACC 合理范围
    if wacc_val is not None:
        if wacc_val < 0.06 or wacc_val > 0.18:
            issues.append({
                "type": "WACC范围异常",
                "detail": f"WACC={wacc_val:.2%}，A股合理范围通常 6%-18%",
                "severity": "WARNING",
                "fix": "请核查 WACC 计算，检查无风险利率和 Beta 输入"
            })

    return issues, wacc_cell, g_cell


def check_balance_sheet(ws, sheet_name):
    """检查资产负债表是否平衡（资产 = 负债 + 股东权益）"""
    issues = []
    assets_row = None
    liab_row = None
    equity_row = None

    for row in ws.iter_rows(max_col=2):
        label = ws.cell(row=row[0].row, column=1).value or ""
        if isinstance(label, str):
            if "资产总计" in label or "总资产" in label:
                assets_row = row[0].row
            elif "负债合计" in label or "负债总计" in label:
                liab_row = row[0].row
            elif "股东权益" in label or "净资产" in label or "所有者权益" in label:
                equity_row = row[0].row

    if assets_row and liab_row and equity_row:
        for col in range(2, min(ws.max_column + 1, 8)):
            a = ws.cell(row=assets_row, column=col).value
            l = ws.cell(row=liab_row, column=col).value
            e = ws.cell(row=equity_row, column=col).value
            col_label = ws.cell(row=3, column=col).value or ws.cell(row=4, column=col).value or f"Col{col}"
            if isinstance(a, (int, float)) and isinstance(l, (int, float)) and isinstance(e, (int, float)):
                diff = abs(a - l - e)
                if diff > 0.1:  # 允许 0.1 亿元误差
                    issues.append({
                        "type": "资产负债表不平衡",
                        "detail": f"{col_label}: 资产={a:.2f} ≠ 负债{l:.2f} + 权益{e:.2f}，差额={diff:.2f}亿",
                        "severity": "CRITICAL",
                        "fix": "检查负债或股东权益的汇总公式"
                    })

    return issues


def print_report(excel_path, all_results):
    """输出验证报告"""
    print(f"\n{'='*60}")
    print(f"Excel 验证报告: {excel_path}")
    print(f"{'='*60}")

    total_critical = sum(1 for r in all_results if r.get("severity") == "CRITICAL")
    total_warning = sum(1 for r in all_results if r.get("severity") == "WARNING")
    total_errors = sum(1 for r in all_results if "error" in r)
    total_hardcoded = sum(1 for r in all_results if "hint" in r)

    if not all_results:
        print("✅ 所有检查通过，未发现问题\n")
        return 0

    print(f"\n发现 {len(all_results)} 个问题：{total_critical} 个严重 / {total_warning} 个警告 / {total_errors} 个公式错误 / {total_hardcoded} 个硬编码数值\n")

    if total_errors:
        print("【公式错误（#DIV/0! / #REF! 等）】")
        for r in all_results:
            if "error" in r:
                print(f"  ❌ {r['sheet']}!{r['cell']} [{r['row_label']}]: {r['error']}")
        print()

    if total_hardcoded:
        print("【硬编码数值（应改为公式引用）】")
        for r in all_results:
            if "hint" in r:
                print(f"  ⚠️  {r['cell']} [{r['label']}]: {r['value']} — {r['hint']}")
        print()

    critical_issues = [r for r in all_results if r.get("severity") == "CRITICAL"]
    warning_issues = [r for r in all_results if r.get("severity") == "WARNING"]

    if critical_issues:
        print("【严重逻辑错误（CRITICAL）】")
        for r in critical_issues:
            print(f"  ❌ [{r['type']}] {r['detail']}")
            print(f"     修复建议: {r['fix']}")
        print()

    if warning_issues:
        print("【警告（WARNING）】")
        for r in warning_issues:
            print(f"  ⚠️  [{r['type']}] {r['detail']}")
            print(f"     修复建议: {r['fix']}")
        print()

    return total_critical + total_errors  # 返回严重问题数（0=通过）


def main():
    parser = argparse.ArgumentParser(description="Excel 财务模型验证工具")
    parser.add_argument("excel_path", help="Excel 文件路径")
    parser.add_argument("--sheet", default=None, help="指定 Sheet 名称（默认扫描全部）")
    parser.add_argument("--ratio-start", type=int, default=None, help="比率区起始行（用于硬编码检测）")
    args = parser.parse_args()

    try:
        wb = openpyxl.load_workbook(args.excel_path, data_only=True)
    except FileNotFoundError:
        print(f"❌ 文件不存在: {args.excel_path}")
        sys.exit(1)
    except Exception as e:
        print(f"❌ 无法打开文件: {e}")
        sys.exit(1)

    sheets = [args.sheet] if args.sheet and args.sheet in wb.sheetnames else wb.sheetnames
    print(f"\n检查 Sheet: {sheets}")

    all_results = []

    for sheet_name in sheets:
        ws = wb[sheet_name]
        print(f"\n--- {sheet_name} ---")

        # 1. 扫描公式错误
        errors = scan_errors(ws, sheet_name)
        if errors:
            all_results.extend(errors)
            print(f"  发现 {len(errors)} 个公式错误")
        else:
            print("  ✅ 无公式错误")

        # 2. DCF 逻辑检查
        dcf_issues, _, _ = check_dcf_logic(ws, sheet_name)
        if dcf_issues:
            all_results.extend(dcf_issues)
        else:
            print("  ✅ DCF 约束检查通过（或非 DCF Sheet）")

        # 3. 资产负债表平衡检查
        bs_issues = check_balance_sheet(ws, sheet_name)
        if bs_issues:
            all_results.extend(bs_issues)
        else:
            print("  ✅ 资产负债表平衡检查通过（或无资产负债表）")

    exit_code = print_report(args.excel_path, all_results)
    sys.exit(0 if exit_code == 0 else 1)


if __name__ == "__main__":
    main()
