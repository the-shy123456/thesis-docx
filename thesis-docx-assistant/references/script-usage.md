# Script Usage

## 1. Normalize Word styles

Use `scripts/normalize_word_styles.ps1` after `check_word_com.ps1` confirms
that Word COM automation is available.

Example:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/normalize_word_styles.ps1 `
  -InputPath .\draft.docx `
  -OutputPath .\draft.normalized.docx
```

The script:

1. Opens the Word document through COM automation.
2. Resolves or creates target styles.
3. Applies style definitions for body text, Heading 1-3, figure captions, and
   table captions.
4. Uses existing style hints first, then falls back to text-pattern detection.
5. Saves a new file unless `-InPlace` is provided.

Use `-ConfigPath` to pass a UTF-8 JSON file with custom style names or school
rules.

Repository example:

```text
examples/word-style-config.sample.json
```

Example JSON:

```json
{
  "styles": {
    "Body": {
      "TargetStyleName": "正文",
      "FontName": "Times New Roman",
      "EastAsiaFontName": "宋体",
      "Size": 12,
      "LineSpacingMultiple": 1.5
    },
    "FigureCaption": {
      "TargetStyleName": "图标注"
    },
    "TableCaption": {
      "TargetStyleName": "表标注"
    }
  }
}
```

## 2. Render Mermaid figures

Use `scripts/render_mermaid_figure.ps1` to render Mermaid source into thesis
figure assets.

Example:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/render_mermaid_figure.ps1 `
  -InputPath .\architecture.mmd `
  -OutputPath .\figures\architecture.svg `
  -Theme base `
  -Width 1800 `
  -Height 1200 `
  -Scale 2
```

The script:

1. Prefers local `mmdc` when it exists.
2. Falls back to `npx @mermaid-js/mermaid-cli` when needed.
3. Generates a clean academic default theme config if no config is provided.
4. Produces SVG, PNG, or PDF output depending on the output filename and CLI
   behavior.

Repository example:

```text
examples/architecture.sample.mmd
```

For thesis work, prefer SVG first, then convert to PNG only when the target
submission system or Word workflow requires raster images.
