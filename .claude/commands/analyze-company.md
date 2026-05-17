---
description: 对一家 A 股上市公司执行完整投研分析，输出专业报告
argument-hint: "[公司名称，如：宁德时代、比亚迪]"
---

# 完整投研分析命令

## 使用示例

```
/analyze-company 宁德时代
/analyze-company 比亚迪
/analyze-company 天赐材料
```

## 执行流程

### Step 1：解析输入

从参数中提取公司名称。若未提供，询问：
> "请输入要分析的公司名称（如：宁德时代）"

### Step 2：确认分析深度

询问用户：
> "需要什么深度的分析？"
> - **完整分析**（默认）：竞争 + 财务 + 估值，约需 5-10 分钟
> - **快速分析**：竞争 + 财务，不做深度估值，约需 2-3 分钟

### Step 3：启动 Master Orchestrator

将以下信息传递给 `master-orchestrator` Agent：
- 公司名称
- 分析深度
- 用户特殊关注点（如有）

Master Orchestrator 将自动执行：
1. 调用 `financial-data-fetcher` 获取最新数据
2. 并行调用 `competitor-analyst`、`financial-analyst`、`valuation-expert`
3. 调用 `quality-checker` 审核报告
4. 调用 `research-report-summarizer` 输出最终报告

### Step 4：输出确认

报告生成后告知用户：
> "✅ 分析完成！报告已保存至 `outputs/[公司名]-研究报告-[日期].md`"
> 并在对话窗口显示执行摘要。
