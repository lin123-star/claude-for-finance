---
description: 用麦肯锡结构化思维深度拆解一篇行业报告或公司研报。支持 PDF 上传、粘贴文本、URL 链接。输出：金字塔结构还原 + 五层Why So追问 + So What行动转化 + MECE完整性审查 + 质量评分 + 更优研究框架建议。用法：/analyze-report [附上内容或URL]
---

请先读取并加载 skill 文件：

```
Read: c:\Users\Pulin\.claude\skills\research-report-analyst\SKILL.md
```

加载完成后，按照 skill 中定义的六步框架（Step 0-6）对用户提供的研报内容进行完整分析。

## 输入方式

用户可以通过以下任意方式提供研报内容：

1. **PDF 附件**：直接读取文件内容
2. **粘贴文本**：`/analyze-report` 后直接粘贴报告段落
3. **URL 链接**：使用 WebFetch 抓取网页正文

## 执行要求

- 六步必须完整执行，不能因内容长而省略任一步骤
- Why So 五层追问是核心，每层必须区分"报告原文依据"和"分析师自行推断"
- 质量评分必须给出具体数值，不能只写定性评级
- 最后的"更优研究框架建议"必须具体可操作，不能泛泛而谈

## 保存输出（可选）

若用户要求保存分析结果，写入：
```
d:\claude-for-finance\outputs\研报分析\[标的名]_研报分析_[日期].md
```
用 Python `pathlib.write_text(encoding='utf-8')` 写入，确保中文不乱码。
