# thesis-docx

## 中文

`thesis-docx` 是一个面向毕业论文 / 学位论文场景的开源工具项目。
它的核心目标很简单：把论文文档处理中那些重复、易错、又很依赖格式一致性的工作沉淀下来。

当前项目主要提供三类能力：

- 论文 Word 文档的样式归一化
- 基于真实材料的 Mermaid 论文插图渲染
- 面向 Codex 的 skill 接入

这个仓库 **不只用于 Codex**。

- `scripts/` 里的脚本可以独立运行，其他 AI、编辑器、自动化平台也能调用
- `references/` 里的规范说明也可以单独复用
- `SKILL.md` 和 `agents/openai.yaml` 只是 Codex 的接入层

也就是说，这个项目可以同时被当作：

1. 一个通用的论文文档处理脚本仓库
2. 一个 Codex skill

### 功能

- 检查本机是否存在桌面版 Microsoft Word，以及是否可用 COM/DOM 自动化
- 批量统一正文、标题、图题注、表题注等 Word 样式
- 将 Mermaid 源文件渲染为适合论文插图使用的 SVG、PNG 或 PDF
- 强制要求图内容来自真实代码、表结构、接口文档、项目文档或论文材料，避免虚构

### 仓库结构

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

### 主要脚本

- `scripts/check_word_com.ps1`
- `scripts/normalize_word_styles.ps1`
- `scripts/render_mermaid_figure.ps1`

### 运行要求

推荐环境：

- Windows
- 桌面版 Microsoft Word
- PowerShell
- Node.js
- Python 3

校验 skill 时需要：

```powershell
python -m pip install pyyaml
```

### 直接使用脚本

检查 Word 自动化环境：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\check_word_com.ps1 -Json
```

统一论文样式：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\normalize_word_styles.ps1 `
  -InputPath C:\path\to\draft.docx `
  -OutputPath C:\path\to\draft.normalized.docx
```

渲染 Mermaid 插图：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\render_mermaid_figure.ps1 `
  -InputPath .\examples\architecture.sample.mmd `
  -OutputPath C:\path\to\architecture.svg `
  -Theme base `
  -Width 1800 `
  -Height 1200 `
  -Scale 2
```

如果学校模板有自定义样式名，可基于这个示例配置调整：

```text
examples/word-style-config.sample.json
```

### 作为 Codex skill 使用

如果你希望 Codex 自动发现它，把仓库放到：

```text
%USERPROFILE%\.codex\skills\thesis-docx
```

或直接 clone：

```powershell
git clone https://github.com/the-shy123456/thesis-docx.git $env:USERPROFILE\.codex\skills\thesis-docx
```

### 校验

```powershell
python C:\Users\85280\.codex\skills\.system\skill-creator\scripts\quick_validate.py .
```

### License

MIT

## English

`thesis-docx` is an open-source toolkit for thesis and dissertation document
workflows. Its goal is to package the repetitive and format-sensitive parts of
thesis editing into reusable assets.

The project currently provides three kinds of functionality:

- Word style normalization for thesis documents
- Mermaid figure rendering based on real source material
- Codex skill integration

This repository is **not limited to Codex**.

- the scripts in `scripts/` can be executed directly from other AI tools,
  editors, or automation platforms
- the documents in `references/` can be reused independently
- `SKILL.md` and `agents/openai.yaml` are only the Codex integration layer

So the repository can be used both as:

1. a general-purpose thesis document automation toolkit
2. a Codex skill

### Features

- checks whether desktop Microsoft Word and COM/DOM automation are available
- batch normalizes body text, headings, figure captions, and table captions
- renders Mermaid source into SVG, PNG, or PDF assets for thesis figures
- requires figure content to be grounded in real code, schema, API docs,
  project docs, or thesis material

### Repository Layout

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

### Main Scripts

- `scripts/check_word_com.ps1`
- `scripts/normalize_word_styles.ps1`
- `scripts/render_mermaid_figure.ps1`

### Requirements

Recommended environment:

- Windows
- desktop Microsoft Word
- PowerShell
- Node.js
- Python 3

For skill validation:

```powershell
python -m pip install pyyaml
```

### Use the Scripts Directly

Check Word automation:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\check_word_com.ps1 -Json
```

Normalize thesis styles:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\normalize_word_styles.ps1 `
  -InputPath C:\path\to\draft.docx `
  -OutputPath C:\path\to\draft.normalized.docx
```

Render a Mermaid figure:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\render_mermaid_figure.ps1 `
  -InputPath .\examples\architecture.sample.mmd `
  -OutputPath C:\path\to\architecture.svg `
  -Theme base `
  -Width 1800 `
  -Height 1200 `
  -Scale 2
```

If your school template uses custom style names, start from:

```text
examples/word-style-config.sample.json
```

### Use as a Codex Skill

If you want Codex to discover the skill automatically, place the repository at:

```text
%USERPROFILE%\.codex\skills\thesis-docx
```

or clone it directly:

```powershell
git clone https://github.com/the-shy123456/thesis-docx.git $env:USERPROFILE\.codex\skills\thesis-docx
```

### Validate

```powershell
python C:\Users\85280\.codex\skills\.system\skill-creator\scripts\quick_validate.py .
```

### License

MIT
