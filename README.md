# thesis-docx-assistant

A Codex skill for thesis and dissertation Word workflows, focused on:

- thesis content revision under explicit academic constraints
- style-driven Word formatting through Microsoft Word COM automation
- evidence-based Mermaid diagram generation for thesis figures
- LaTeX-oriented code listing guidance

This repository currently ships one skill:

- `thesis-docx-assistant/`

## Why this skill exists

Thesis editing is not just a writing task. It usually combines:

- Word template compliance
- strict heading, caption, and pagination rules
- diagram generation that must match real project materials
- code listings that should be suitable for academic documents

This skill packages those rules into a reusable Codex skill and includes helper
scripts for the most repetitive or fragile parts.

## What is included

- `thesis-docx-assistant/SKILL.md`
  - the skill definition and workflow
- `thesis-docx-assistant/scripts/check_word_com.ps1`
  - checks whether Microsoft Word COM/DOM automation is available
- `thesis-docx-assistant/scripts/normalize_word_styles.ps1`
  - batch normalizes thesis body text, headings, and captions
- `thesis-docx-assistant/scripts/render_mermaid_figure.ps1`
  - renders Mermaid source into thesis-ready figure assets
- `examples/word-style-config.sample.json`
  - sample style override file for school-specific templates
- `examples/architecture.sample.mmd`
  - sample Mermaid input file

## Requirements

### Required for the full workflow

- Windows
- desktop Microsoft Word
- PowerShell
- Node.js

### Required by specific operations

- `render_mermaid_figure.ps1`
  - `mmdc`, or `npx` with access to `@mermaid-js/mermaid-cli`
- skill validation
  - Python 3
  - `PyYAML`

## Install

### Option 1: copy the skill folder

Copy `thesis-docx-assistant/` into:

```text
%USERPROFILE%\.codex\skills\
```

Result:

```text
%USERPROFILE%\.codex\skills\thesis-docx-assistant\
```

### Option 2: clone this repository and copy the folder

```powershell
git clone <your-repo-url>
Copy-Item .\thesis-docx-assistant $env:USERPROFILE\.codex\skills -Recurse
```

## Quick start

### 1. Check Word automation

```powershell
powershell -ExecutionPolicy Bypass -File .\thesis-docx-assistant\scripts\check_word_com.ps1 -Json
```

If Word COM or DOM automation is unavailable, stop the high-fidelity automation
path and tell the user to install desktop Microsoft Word.

### 2. Normalize thesis styles

```powershell
powershell -ExecutionPolicy Bypass -File .\thesis-docx-assistant\scripts\normalize_word_styles.ps1 `
  -InputPath C:\path\to\draft.docx `
  -OutputPath C:\path\to\draft.normalized.docx
```

### 3. Render a Mermaid thesis figure

```powershell
powershell -ExecutionPolicy Bypass -File .\thesis-docx-assistant\scripts\render_mermaid_figure.ps1 `
  -InputPath .\examples\architecture.sample.mmd `
  -OutputPath C:\path\to\architecture.svg `
  -Theme base `
  -Width 1800 `
  -Height 1200 `
  -Scale 2
```

## School-specific styles

Use `examples/word-style-config.sample.json` as a starting point when a school
template uses different style names such as `正文`, `图标注`, or `表标注`.

Example:

```powershell
powershell -ExecutionPolicy Bypass -File .\thesis-docx-assistant\scripts\normalize_word_styles.ps1 `
  -InputPath C:\path\to\draft.docx `
  -OutputPath C:\path\to\draft.normalized.docx `
  -ConfigPath .\examples\word-style-config.sample.json
```

## Publishing notes

- Keep the skill folder name as `thesis-docx-assistant`
- Do not commit temporary Word lock files such as `~$*.docx`
- Do not generate diagrams from invented structure or fake project data
- Use Mermaid figures only when the source materials are real and sufficient

## Validate before release

```powershell
python -m pip install pyyaml
python C:\Users\85280\.codex\skills\.system\skill-creator\scripts\quick_validate.py .\thesis-docx-assistant
```

## License

MIT
