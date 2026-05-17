---
description: 持续监测仓位假设是否成立，发现变化时告警。用法：/monitor 单公司监测；/monitor [公司名] 指定公司；/monitor-add [公司名] 将公司加入监测清单
---

请调用 `monitor-agent` Agent 执行监测任务。

## 判断用法

**`/monitor`（无参数）：批量监测**

读取 `d:\claude-for-finance\watchlist.json` 中所有 `active: true` 的公司，依次执行假设核验，最后输出汇总面板。

---

**`/monitor [公司名]`：单公司快速核验**

只对指定公司执行监测流程。若该公司不在 watchlist.json 中，提示：
> "[公司名]" 不在监测清单中。是否要先添加？输入 `/monitor-add [公司名]` 初始化监测条目。

---

**`/monitor-add [公司名]`：将公司加入监测清单**

执行以下步骤：
1. 在 `outputs/[公司名]/` 目录查找最新研究报告，读取其中的"核心假设"和"投资逻辑"章节
2. 若无报告，询问用户：是否先运行 `/analyze-company [公司名]` 生成初始报告？
3. 提取 3-5 条关键假设，提示用户确认（可修改）
4. 询问告警阈值（若用户无意见，使用默认：毛利率下滑 2pct、市占率下滑 3pct、订单增速转负）
5. 将新条目写入 watchlist.json，`active: true`，`last_checked: null`

**新条目模板：**
```json
{
  "name": "[公司名]",
  "active": true,
  "initial_report": "d:/claude-for-finance/outputs/[公司名]/[公司名]-研究报告-[日期].md",
  "sector": "[所属板块]",
  "last_checked": null,
  "alert_level": null,
  "key_assumptions": [
    "[假设1]",
    "[假设2]",
    "[假设3]"
  ],
  "thresholds": {
    "gross_margin_decline": "毛利率连续两季度下滑超 2pct",
    "market_share_decline": "市场份额下滑超 3pct",
    "order_growth_negative": "订单同比增速转负"
  }
}
```
