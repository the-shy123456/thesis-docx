# thesis-docx

[![中文](https://img.shields.io/badge/中文-说明-1677ff)](#中文)
[![English](https://img.shields.io/badge/English-Docs-111827)](#english)

## 中文

`thesis-docx` 是一个用于毕业论文 / 学位论文文档处理的开源工具仓库。

它的重点不是绑定某一个 AI 平台，而是把论文处理里那些可复用、可脚本化、
又容易出错的部分沉淀成一套可以直接调用的工具和规范。

它可以被：

- Claude Code 使用
- Cursor / Windsurf / Roo Code 等可执行本地命令的工具使用
- 任何能读取仓库文件并调用 PowerShell 脚本的自动化流程使用

仓库里的核心内容是：

- `scripts/`：可直接执行的脚本
- `references/`：可复用的论文处理规则
- `SKILL.md`：给支持 skill 机制的工具使用的说明入口

### 功能

- 检查本机是否安装了桌面版 Microsoft Word，以及是否支持 COM/DOM 自动化
- 批量统一论文正文、标题、图题注、表题注等 Word 样式
- 将 Mermaid 源文件渲染成适合论文插图使用的 SVG、PNG 或 PDF
- 约束图内容必须基于真实材料，避免虚构结构

### 仓库结构

```text
.
├── SKILL.md
├── scripts/
├── references/
├── examples/
├── README.md
└── LICENSE
```

### 环境要求

推荐环境：

- Windows
- 桌面版 Microsoft Word
- PowerShell
- Node.js
- Python 3

其中：

- `normalize_word_styles.ps1` 依赖 Word COM 自动化
- `render_mermaid_figure.ps1` 依赖 `mmdc` 或 `npx @mermaid-js/mermaid-cli`

### 快速开始

1. 把仓库 clone 或下载到任意位置
2. 直接运行里面的脚本，或者把整个仓库文件夹复制到你需要的目录
3. 如果某个 AI 工具支持 skills / prompts / repo tools，也可以直接让它读取这个仓库

例如先下载到本地：

```powershell
git clone https://github.com/the-shy123456/thesis-docx.git
```

然后你可以：

- 直接在这个目录里运行脚本
- 把整个 `thesis-docx` 文件夹复制到其他工作区
- 把整个 `thesis-docx` 文件夹复制到支持 skill 目录机制的位置

### 直接运行脚本

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

### 可选：作为 skill 使用

如果你想把它接入支持 skill 目录机制的工具，最简单的方式不是拆文件，而是直接复制整个仓库目录到目标工具的 skill 目录。

例如：

```text
<your-skill-directory>\thesis-docx
```

也就是说，用户只需要：

1. 下载或 clone 这个仓库
2. 把整个仓库文件夹复制到目标位置

就可以了。

### License

MIT

## English

`thesis-docx` is an open-source repository for thesis and dissertation document
workflows.

The goal is not to bind the project to a single AI platform. Instead, it
packages the reusable, scriptable, and error-prone parts of thesis processing
into a toolkit that can be called directly.

It can be used by:

- Claude Code
- Cursor / Windsurf / Roo Code
- any automation pipeline that can read repository files and run local
  PowerShell scripts

The core parts of the repository are:

- `scripts/`: directly executable scripts
- `references/`: reusable thesis workflow rules
- `SKILL.md`: an entry point for tools that support skill-style workflows

### Features

- checks whether desktop Microsoft Word and COM/DOM automation are available
- batch normalizes body text, headings, figure captions, and table captions
- renders Mermaid source into SVG, PNG, or PDF assets for thesis figures
- requires figure content to be grounded in real project material instead of
  fabricated structure

### Repository Layout

```text
.
├── SKILL.md
├── scripts/
├── references/
├── examples/
├── README.md
└── LICENSE
```

### Requirements

Recommended environment:

- Windows
- desktop Microsoft Word
- PowerShell
- Node.js
- Python 3

Notes:

- `normalize_word_styles.ps1` depends on Word COM automation
- `render_mermaid_figure.ps1` depends on `mmdc` or `npx @mermaid-js/mermaid-cli`

### Quick Start

1. Clone or download the repository anywhere
2. Run the scripts directly, or copy the whole repository folder where you want
   to use it
3. If your AI tool supports skills, prompts, or repo-based tooling, point it to
   this repository directly

Clone example:

```powershell
git clone https://github.com/the-shy123456/thesis-docx.git
```

Then you can:

- run the scripts directly in this repository
- copy the whole `thesis-docx` folder into another workspace
- copy the whole `thesis-docx` folder into any location that supports a
  skill-directory workflow

### Run the Scripts Directly

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

### Optional: use as a skill

If you want to plug it into tools that support skill-directory workflows, copy
the whole repository folder directly into the target skill directory instead of
splitting files.

Example:

```text
<your-skill-directory>\thesis-docx
```

In other words, users only need to:

1. download or clone this repository
2. copy the entire repository folder to the target location

### License

MIT
