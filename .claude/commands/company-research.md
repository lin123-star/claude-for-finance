---
description: 对一家公司做洞察导向研究。先建事实底账，再挖掘值得深挖的角度，然后只围绕这些角度做专项分析，最后产出可证伪的洞察。用法：/company-research [公司名]
argument-hint: "[公司名称，如：宁德时代、比亚迪]"
---

# 洞察导向公司研究

## 用法

```
/company-research 宁德时代
/company-research 比亚迪
```

与 `/analyze-company` 的差别：本命令先决定「该问什么」，再调用专项分析。角度打分后如果没有重要性足够的问题，就停在角度文件，不继续堆报告。

## 执行

公司名缺失时，只问一次公司名称，然后开始。不要再询问分析深度。

把下面的顺序交给 `master-orchestrator`，并声明这是「洞察导向公司研究」：

1. `source-curator` 写事实底账。身份对不上或标准清单大面积缺失时，先把缺口给用户，不往下编。
2. `angle-miner` 写分析角度。若结论是「当前没有值得深挖的角度」，把角度文件交还用户并停止。
3. 按角度文件的任务单调用专项代理。每个代理只回答任务单上的问题，必须引用事实 ID。
   - 财务问题 → `financial-analyst`
   - 竞争问题 → `competitor-analyst`
   - 估值问题 → `valuation-expert`（财务底稿未完成时不做）
   - 任务单出现周期或政策时 → `macro-researcher`
4. `signal-integrator` 对专项结论做矛盾仲裁。
5. `insight-engine` 写洞察文件。
6. `quality-checker` 做洞察审计。未过的洞察退回 `insight-engine` 降级为观察，不直接放进报告结论。
7. `research-report-summarizer` 写最终报告。报告第一节必须是洞察，不是公司简介。买卖评级只在这一步出现，且每一条评级理由要引用洞察 ID。

## 产物

```
outputs/[公司名]/[公司名]-事实底账-[日期].md
outputs/[公司名]/[公司名]-分析角度-[日期].md
outputs/[公司名]/[公司名]-洞察-[日期].md
outputs/[公司名]/[公司名]-研究报告-[日期].md
```

完成后在对话里先给三块：入选角度、洞察卡片、仍需人判断的问题。再给出报告路径。
