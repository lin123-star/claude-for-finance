---
description: 生成持仓监测仪表板，可视化全部 watchlist 公司状态。用法：/dashboard（文字面板）；/dashboard --export（同时导出 HTML 文件）
---

请调用 `dashboard-agent` Agent 执行仪表板生成任务。

## 判断用法

**`/dashboard`（无参数）：文字面板**

生成完整的 Markdown 监测仪表板，直接在对话中显示。包含：
- 组合风险概览（告警分布统计）
- 持仓明细表（所有 watchlist 公司状态）
- 触发告警详情
- 假设库覆盖率
- 宏观背景摘要（取最新宏观面板）
- 待办提醒（超期未检查、无假设追踪、待重新分析）

数据来源全部来自本地文件，**不联网**，速度快。若告警状态过期，建议先运行 `/monitor` 更新后再查看仪表板。

---

**`/dashboard --export`：文字面板 + 导出 HTML**

在输出文字面板的同时，生成一个可在浏览器中直接打开的 HTML 仪表板文件，保存到：
```
outputs/监测仪表板/dashboard-[YYYY-MM-DD].html
```

HTML 文件包含完整的视觉样式（颜色区分告警级别），无外部依赖，双击即可查看。

---

**注意事项**
- 仪表板数据以 `watchlist.json` 和 `outputs/假设库/` 为准
- 未运行过 `/monitor` 的公司告警级别显示为"未检查"
- 若需要最新状态，请先 `/monitor`，再 `/dashboard`
