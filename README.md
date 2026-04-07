# thesis-docx

[![中文](https://img.shields.io/badge/中文-说明-1677ff)](#中文)
[![English](https://img.shields.io/badge/English-Docs-111827)](#english)
[![License](https://img.shields.io/badge/license-MIT-16a34a)](./LICENSE)
[![Word](https://img.shields.io/badge/Microsoft_Word-Recommended-185ABD)](#运行前提)
[![Mermaid](https://img.shields.io/badge/Mermaid-Supported-0ea5e9)](#功能)

`thesis-docx` is an audit-first skill for thesis and dissertation Word workflows.
It helps AI agents revise `.docx` documents under strict academic formatting
constraints without blindly normalizing the whole paper.

---

## 中文

### 这是什么

`thesis-docx` 是一个面向毕业论文 / 学位论文场景的通用 skill。

它不是简单的“帮 AI 改 Word 格式”，而是一套更稳的论文工作流约束：

- **先审计，再修改**
- **学校明确规定的格式严格执行**
- **学校没规定的格式默认保留现状**
- **优先通过 Word + PDF 逐页复核，而不是只看结构**

它的目标不是“改得快”，而是**减少论文定稿阶段最容易出现的格式返工**。

### 适合什么场景

这个 skill 适合下面这些论文任务：

- 按学校模板或格式规范修订 Word 文档
- 统一正文、标题、图题注、表题注、参考文献等样式
- 修复目录、页码、分节、交叉引用、图表编号
- 生成基于真实资料的 Mermaid 架构图、流程图、E-R 图
- 在论文里放入 LaTeX 风格的代码片段或伪代码
- 对已改过多轮的论文做最终格式审计

### 为什么这个 skill 和普通“文档助手”不一样

论文文档和一般 Word 文档不一样，很多问题并不是表面样式能看出来的。

例如：

- 样式名看对了，但 `styleId` 实际错了
- 标题看起来没缩进，但底层有 `firstLineChars`
- 题注已经改成 `图6-1`，正文交叉引用还显示旧值
- WPS 看起来正常，Word 打开代码框标题只显示半截
- 分节首页设置会悄悄吃掉页眉或页码

`thesis-docx` 把这些坑当成默认风险来处理，而不是等用户一点点指出来。

### 核心原则

- **Audit First**：先审计，再决定修不修、怎么修
- **Minimal Justified Fixes**：只修有依据的问题，不做全局乱改
- **Style-Driven, Not Chaos-Driven**：优先修样式系统，不堆散乱直接格式
- **PDF-Level Verification**：对论文来说，结构正确不等于页面正确
- **No Fabrication**：图、代码、架构、流程都必须基于真实材料

### 功能

- 生成、补充、改写论文内容
- 按用户给定模板或格式要求修订 Word 文档
- 统一正文、标题、图题注、表题注等样式
- 在论文中生成基于真实材料的 Mermaid 架构图、流程图、E-R 图等
- 在论文中使用 LaTeX 风格的代码排版
- 通过 Word 导出 PDF 并做逐页格式审计
- 审计 OOXML 隐藏问题，例如：
  - `styleId`
  - `firstLineChars`
  - `titlePg`
  - REF 域显示值
  - section 级页眉页脚引用

### 运行前提

- Windows 环境下建议安装桌面版 Microsoft Word，用于高保真 DOCX 编辑与 PDF 导出
- 运行 Python 脚本时建议具备：
  - `python-docx`
  - `lxml`
- 渲染 Mermaid 图时建议具备：
  - Node.js
  - `mmdc` 或可用的 `npx`

### 快速开始

如果你已经把 skill 放到本地并准备让 AI 调用，推荐按这个顺序验证：

```powershell
# 1) 检查 Word COM/DOM 是否可用
powershell -ExecutionPolicy Bypass -File scripts/check_word_com.ps1 -Json

# 2) 先对论文 docx 做 OOXML 审计
python scripts/audit_docx_ooxml.py .\draft.docx --output_json .\draft.audit.json --output_txt .\draft.audit.txt

# 3) 如有需要，先 dry-run 审计批量样式归一化，不直接写入
powershell -ExecutionPolicy Bypass -File scripts/normalize_word_styles.ps1 -InputPath .\draft.docx -AuditOnly

# 4) 真正修完后，再导出 PDF 做逐页复核
powershell -ExecutionPolicy Bypass -File scripts/export_word_pdf.ps1 -DocPath .\draft.docx -PdfPath .\draft.audit.pdf
```

### 推荐工作顺序

1. 检查 Word 与 COM/DOM 是否可用
2. 先读学校规范、模板、截图、已有样式
3. 先跑审计，再决定是否批量改样式
4. 只修学校明确要求或用户明确要求的项目
5. 导出 PDF，逐页复核后再判断是否完成

### 仓库结构

```text
.
├── agents/
├── SKILL.md
├── scripts/
├── references/
├── examples/
├── README.md
└── LICENSE
```

### 目录说明

- `SKILL.md`：skill 主说明
- `agents/`：界面集成元数据
- `scripts/`：供 AI 调用的辅助脚本
- `references/`：补充规则与工作流说明
- `examples/`：示例输入、配置与最终检查清单

推荐优先阅读：

- `references/paper-format-workflow.md`
- `references/failure-patterns-and-quality-gates.md`
- `references/script-usage.md`

### 适用对象

适用于支持以下能力的 AI 工具：

- 读取仓库文件
- 读取 `SKILL.md`
- 在本地执行脚本或命令

### 安装

1. 下载或 clone 此仓库
2. 保持整个仓库文件夹完整
3. 将整个文件夹复制到你希望 AI 使用的位置

示例：

```powershell
git clone https://github.com/the-shy123456/thesis-docx.git
```

之后直接复制整个 `thesis-docx` 文件夹即可，不需要拆文件。

### Star 趋势

[![Star History Chart](https://api.star-history.com/svg?repos=the-shy123456/thesis-docx&type=Date)](https://www.star-history.com/#the-shy123456/thesis-docx&Date)

### License

MIT

---

## English

### What It Is

`thesis-docx` is a general-purpose skill for thesis and dissertation Word
workflows.

It is not just a "format my Word file" helper. It is designed around a safer
workflow for high-stakes academic documents:

- **audit first**
- **strictly enforce school-defined rules**
- **preserve unspecified formatting by default**
- **verify through Word + PDF review, not structure alone**

The goal is not merely to edit fast, but to reduce the format regressions that
usually appear near final submission.

### Good Fit For

This skill is suitable for:

- revising Word thesis documents under school-specific rules
- normalizing body text, headings, figure captions, table captions, and references
- fixing TOC, page numbers, sections, cross-references, and caption numbering
- generating Mermaid-based architecture, flowchart, and E-R figures from real materials
- presenting code snippets or pseudocode in thesis-friendly form
- running final-format audits on heavily revised theses

### Why It Is Different From Generic DOCX Helpers

Thesis documents often fail for reasons that are invisible at the surface level.

Examples:

- a style name looks correct, but the effective `styleId` is wrong
- a heading seems unindented, but still carries `firstLineChars`
- a caption says `Fig 6-1`, but the body still renders an old REF value
- WPS looks fine, while Word clips a code-box title row
- section first-page settings silently remove page headers or page numbers

`thesis-docx` treats these as default risks, not rare edge cases.

### Core Principles

- **Audit First**: inspect before editing
- **Minimal Justified Fixes**: change only what is supported by requirements
- **Style-Driven Workflow**: repair the style system instead of stacking direct formatting
- **PDF-Level Verification**: structure correctness is not enough for thesis delivery
- **No Fabrication**: figures, architecture, and code must be grounded in real materials

### Features

- generate, extend, and revise thesis content
- revise Word documents under user-provided formatting requirements
- normalize body text, headings, figure captions, and table captions
- generate Mermaid architecture diagrams, flowcharts, and E-R diagrams from real source material
- use LaTeX-style code formatting in thesis documents
- export the thesis to PDF for page-by-page format review
- audit hidden OOXML issues such as:
  - `styleId`
  - `firstLineChars`
  - `titlePg`
  - REF field display text
  - section-level header/footer references

### Runtime Prerequisites

- On Windows, desktop Microsoft Word is strongly recommended for high-fidelity DOCX editing and PDF export
- Python scripts typically expect:
  - `python-docx`
  - `lxml`
- Mermaid rendering typically expects:
  - Node.js
  - `mmdc` or a usable `npx`

### Quick Start

After placing the skill locally, a good first-run sequence is:

```powershell
# 1) Check Word COM/DOM availability
powershell -ExecutionPolicy Bypass -File scripts/check_word_com.ps1 -Json

# 2) Audit the DOCX / OOXML before editing
python scripts/audit_docx_ooxml.py .\draft.docx --output_json .\draft.audit.json --output_txt .\draft.audit.txt

# 3) Dry-run style normalization first
powershell -ExecutionPolicy Bypass -File scripts/normalize_word_styles.ps1 -InputPath .\draft.docx -AuditOnly

# 4) Export to PDF for page-by-page review
powershell -ExecutionPolicy Bypass -File scripts/export_word_pdf.ps1 -DocPath .\draft.docx -PdfPath .\draft.audit.pdf
```

### Recommended Operating Order

1. Check Word and COM/DOM availability first
2. Read the school guide, template, screenshots, and existing styles before editing
3. Audit first, then decide whether batch normalization is justified
4. Fix only school-defined or user-defined requirements
5. Export to PDF and review page by page before claiming completion

### Repository Layout

```text
.
├── agents/
├── SKILL.md
├── scripts/
├── references/
├── examples/
├── README.md
└── LICENSE
```

### Directory Guide

- `SKILL.md`: main skill instructions
- `agents/`: UI-facing integration metadata
- `scripts/`: helper scripts invoked by AI when needed
- `references/`: supporting rules and workflow material
- `examples/`: sample inputs, configs, and final audit checklist

Recommended first reads:

- `references/paper-format-workflow.md`
- `references/failure-patterns-and-quality-gates.md`
- `references/script-usage.md`

### Intended For

This skill fits AI tools that can:

- read repository files
- read `SKILL.md`
- run local scripts or commands

### Installation

1. Download or clone this repository
2. Keep the whole repository folder intact
3. Copy the entire folder to the location where you want your AI tool to use it

Example:

```powershell
git clone https://github.com/the-shy123456/thesis-docx.git
```

Then copy the whole `thesis-docx` folder as-is. Do not split the files.

### Star History

[![Star History Chart](https://api.star-history.com/svg?repos=the-shy123456/thesis-docx&type=Date)](https://www.star-history.com/#the-shy123456/thesis-docx&Date)

### License

MIT
