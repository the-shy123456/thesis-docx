# thesis-docx-assistant

**中文**

`thesis-docx-assistant` 是一个面向毕业论文 / 学位论文场景的 Codex
skill，目标不是单纯“帮你写点文字”，而是把论文常见的几个高频痛点一起
收进来：

- 按学校模板或明确格式要求修订 Word 文档
- 统一正文、标题、图题注、表题注等样式
- 基于真实材料生成适合论文场景的 Mermaid 图
- 在论文中使用更规范的 LaTeX 风格代码排版

它适合开源分发，也适合直接 clone 到本地作为单个 skill 使用。

**English**

`thesis-docx-assistant` is a Codex skill for thesis and dissertation workflows.
It is designed for more than just text generation. It packages several common
academic-document tasks into one reusable skill:

- revising Word documents under explicit school formatting rules
- normalizing body text, headings, figure captions, and table captions
- generating Mermaid diagrams only from real project materials
- supporting LaTeX-oriented code listing workflows for thesis writing

This repository is structured as a **single-skill repository** so it can be
installed directly.

## Features

**中文**

- 先检查本机是否有桌面版 Microsoft Word，以及是否可用 COM/DOM 自动化
- 优先用“样式”而不是零散手工格式去统一论文排版
- 图的内容必须来自真实代码、表结构、项目文档或论文材料，禁止虚构
- 提供两个可执行 helper 脚本，减少重复造轮子

**English**

- checks whether desktop Microsoft Word and COM/DOM automation are available
- favors style-driven formatting over scattered manual formatting
- requires Mermaid figures to be grounded in real source materials
- ships helper scripts for repetitive and environment-sensitive operations

## Repository Layout

```text
.
├── SKILL.md
├── agents/
├── scripts/
├── references/
├── examples/
├── README.md
└── LICENSE
```

## Included Scripts

- `scripts/check_word_com.ps1`
  - Check whether Word COM/DOM automation is available.
- `scripts/normalize_word_styles.ps1`
  - Batch normalize thesis body text, Heading 1-3, figure captions, and table
    captions.
- `scripts/render_mermaid_figure.ps1`
  - Render Mermaid source into SVG, PNG, or PDF assets for thesis figures.

## Requirements

**中文**

完整工作流建议具备：

- Windows
- 桌面版 Microsoft Word
- PowerShell
- Node.js
- Python 3

**English**

Recommended environment for the full workflow:

- Windows
- desktop Microsoft Word
- PowerShell
- Node.js
- Python 3

Additional package for validation:

```powershell
python -m pip install pyyaml
```

## Install

### Recommended

直接把仓库 clone 到 Codex 的 skills 目录下，并保持目录名为
`thesis-docx-assistant`。

Clone this repository directly into your Codex skills directory and keep the
folder name as `thesis-docx-assistant`.

```powershell
git clone <your-repo-url> $env:USERPROFILE\.codex\skills\thesis-docx-assistant
```

### Alternative

如果你已经 clone 到别处，也可以手动复制整个仓库目录到：

If you cloned it elsewhere, you can also copy the whole repository directory
to:

```text
%USERPROFILE%\.codex\skills\thesis-docx-assistant
```

## Quick Start

### 1. Check Word automation

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\check_word_com.ps1 -Json
```

If Word COM or DOM automation is unavailable, stop the high-fidelity
automation path and ask the user to install desktop Microsoft Word.

### 2. Normalize thesis styles

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\normalize_word_styles.ps1 `
  -InputPath C:\path\to\draft.docx `
  -OutputPath C:\path\to\draft.normalized.docx
```

### 3. Render a Mermaid figure

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\render_mermaid_figure.ps1 `
  -InputPath .\examples\architecture.sample.mmd `
  -OutputPath C:\path\to\architecture.svg `
  -Theme base `
  -Width 1800 `
  -Height 1200 `
  -Scale 2
```

## School-Specific Style Overrides

如果学校模板里使用 `正文`、`图标注`、`表标注` 之类的自定义样式名，可直接
从这个示例开始改：

If your school template uses custom style names such as `正文`, `图标注`, or
`表标注`, start from:

```text
examples/word-style-config.sample.json
```

Example:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\normalize_word_styles.ps1 `
  -InputPath C:\path\to\draft.docx `
  -OutputPath C:\path\to\draft.normalized.docx `
  -ConfigPath .\examples\word-style-config.sample.json
```

## Open Source Positioning

**中文**

这个 skill 的定位是：

- 把论文工作流里的核心规则写进 `SKILL.md`
- 把重复、易错、环境敏感的操作固化成脚本
- 当脚本不适配时，仍允许 AI 在规则约束下给出临时方案

**English**

The intended design is:

- keep the core academic workflow in `SKILL.md`
- package repetitive and fragile operations into helper scripts
- still allow AI to improvise when a helper script does not fit a specific case

## Validate

```powershell
python C:\Users\85280\.codex\skills\.system\skill-creator\scripts\quick_validate.py .
```

## License

MIT
