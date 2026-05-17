# 金融分析 Agentic 系统

## 项目定位

基于 Claude Code 的 A 股上市公司投研自动化系统。目标是输入公司名，自动完成竞争分析、财务分析、估值建模，输出专业投研报告。

## 目录结构

```
d:\claude-for-finance\
├── .claude\
│   ├── agents\          # 子 Agent 定义
│   ├── commands\        # 斜杠命令
│   └── skills\          # 知识模板（agents 通过 Read 工具加载）
├── 宁德时代\             # 宁德时代专项研究
│   ├── .claude\         # 专项 Agent（覆盖根目录通用版）
│   ├── data\            # 原始财报、行业数据
│   └── outputs\         # 分析报告、Excel 模型
└── financial-services-plugins\  # Anthropic 官方插件库（参考用）
```

## 全局行为规范

- 语言：**中文**输出（除 Excel 公式、Python 代码外）
- 数据来源优先级：① 用户上传文件 ② 同花顺（10jqka.com.cn） ③ 东方财富 ④ SEC EDGAR（境外公司）
- 报告格式：结构化 Markdown，输出到对应项目的 `outputs/` 目录
- Excel 模型：输出到 `outputs/models/`，所有单元格必须用公式，禁止 hardcode 数值
- **数据时效警告**：训练数据存在滞后，凡涉及财务数据必须通过网络搜索获取最新版本

## Skills 加载方式

官方 skills 位于 `financial-services-plugins/financial-analysis/skills/`，**不会自动触发**。
Agent 需要显式加载对应官方路径：

```
# DCF 估值
Read: d:\claude-for-finance\financial-services-plugins\financial-analysis\skills\dcf-model\SKILL.md

# 竞争分析
Read: d:\claude-for-finance\financial-services-plugins\financial-analysis\skills\competitive-analysis\SKILL.md

# 财务分析（三表模型 + Comps）
Read: d:\claude-for-finance\financial-services-plugins\financial-analysis\skills\3-statement-model\SKILL.md
Read: d:\claude-for-finance\financial-services-plugins\financial-analysis\skills\comps-analysis\SKILL.md

# 高端制造业专项分析（供应链 + 产能 + 景气周期 + 访调工作包）
# 触发方式：/manufacturing-analysis [公司名]
Read: d:\claude-for-finance\financial-services-plugins\financial-analysis\skills\high-end-manufacturing-screener\SKILL.md

# 财报附注深度分析（关联交易、会计政策、或有负债、商誉减值、受限资产）
# 触发方式：/footnote-analysis [公司名] [NotebookLM URL]
Read: c:\Users\Pulin\.claude\skills\footnote-analyst\SKILL.md
```

> ⚠️ `.claude/skills/` 下的旧文件已停用，请勿引用。

## 权限声明

当前开放工具：WebSearch、WebFetch（同花顺、investing.com）、Bash（python *）

## Python 环境（重要）

**唯一可用的 Python 解释器**：`C:\Python314\python.exe`（命令行直接用 `python`）

Bash 中运行 Python 脚本的标准写法：
```bash
python d:\claude-for-finance\outputs\[公司名]_三表模型_DCF估值.py
```

**禁止行为**：
- ❌ 不得因为脚本运行失败就创建 run_*.py / run_*.bat / go.py / make_*.py 等 wrapper 脚本
- ❌ 不得尝试 python3、py、python3.14 等变体（系统只有 python 有效）
- ✅ 运行失败时：直接修复原始 .py 文件中的 bug，不要包装

## Agent 调用入口

- 完整投研分析 → `master-orchestrator` agent
- 斜杠命令 → `/analyze-company [公司名]`
- 财报附注分析 → `/footnote-analysis [公司名] [NotebookLM URL]`
