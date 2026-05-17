---
description: 深度解析财报附注，挖掘关联交易、会计政策变更、或有负债、商誉减值、受限资产等隐藏风险与机会
argument-hint: "[公司名称] [NotebookLM链接，可选]"
---

# 财报附注深度分析

## 使用示例

```
/footnote-analysis 鸣志电器
/footnote-analysis 鸣志电器 https://notebooklm.google.com/notebook/xxx
/footnote-analysis 宁德时代 https://notebooklm.google.com/notebook/xxx 2024
```

## 执行流程

### Step 1：解析参数

从参数中提取：
- **公司名称**（必填）
- **NotebookLM URL**（可选，若未提供则询问）
- **年份**（可选，默认最新年报）

若未提供 URL，询问：
> "请提供包含 [公司名] 年报的 NotebookLM 链接，或直接粘贴附注文字内容。"

### Step 2：加载 Skill

Read: c:\Users\Pulin\.claude\skills\footnote-analyst\SKILL.md

### Step 3：执行附注分析

按 `footnote-analyst` skill 定义的五大维度依次分析：

1. 关联交易
2. 会计政策与估计变更
3. 或有负债与表外风险
4. 商誉与无形资产减值测试
5. 受限资产与流动性隐患

每个维度完成后，基于发现决定是否触发追问。

### Step 4：联动现有研究

检查 `d:\claude-for-finance\outputs\[公司名]\` 目录：
- 若存在投研报告 → 说明附注发现对原有假设的修正方向
- 若存在假设追踪文件（`outputs\假设库\[公司名]_假设追踪.md`）→ 更新相关假设

### Step 5：输出报告

按 skill 定义的格式输出报告，保存至：
```
d:\claude-for-finance\outputs\[公司名]\附注分析-[YYYYMMDD].md
```
