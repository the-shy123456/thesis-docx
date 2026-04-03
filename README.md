# thesis-docx

[![中文](https://img.shields.io/badge/中文-说明-1677ff)](#中文)
[![English](https://img.shields.io/badge/English-Docs-111827)](#english)

## 中文

`thesis-docx` 是一个面向毕业论文 / 学位论文场景的通用 skill。

它用于让 AI 在处理论文文档时遵循统一规则，而不是随意生成内容或随意改格式。

### 功能

- 生成、补充、改写论文内容
- 按用户给定模板或格式要求修订 Word 文档
- 统一正文、标题、图题注、表题注等样式
- 在论文中生成基于真实材料的 Mermaid 架构图、流程图、E-R 图等
- 在论文中使用 LaTeX 风格的代码排版

### 核心规则

- 优先检查桌面版 Microsoft Word 和 COM/DOM 自动化能力
- 如果没有 Word 或无法自动化，不强行执行高保真排版
- 用户给了格式要求时，必须严格遵守
- 图和代码内容必须基于用户提供的真实资料，不能虚构
- 仓库内置脚本供 AI 按需调用，通常不需要用户手动运行

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

### 目录说明

- `SKILL.md`：skill 主说明
- `scripts/`：供 AI 调用的辅助脚本
- `references/`：补充规则与说明
- `examples/`：示例输入与样式配置

### License

MIT

## English

`thesis-docx` is a general-purpose skill for thesis and dissertation workflows.

It is designed to help AI follow consistent rules when generating or revising
academic documents, instead of producing content or formatting arbitrarily.

### Features

- generate, extend, and revise thesis content
- revise Word documents under user-provided formatting requirements
- normalize body text, headings, figure captions, and table captions
- generate Mermaid architecture diagrams, flowcharts, and E-R diagrams from
  real source material
- use LaTeX-style code formatting in thesis documents

### Core Rules

- check desktop Microsoft Word and COM/DOM automation first
- do not force high-fidelity formatting when Word automation is unavailable
- follow user-provided formatting requirements strictly
- never fabricate figures or code content without real source material
- bundled scripts are internal helpers for AI tools and usually do not need to
  be run manually by end users

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

### Directory Guide

- `SKILL.md`: main skill instructions
- `scripts/`: helper scripts invoked by AI when needed
- `references/`: supporting rules and reference material
- `examples/`: sample inputs and style configuration examples

### License

MIT
