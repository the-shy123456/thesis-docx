# Thesis Word Formatting Workflow

## 1. Environment Gate

1. 先运行 `scripts/check_word_com.ps1 -Json`。
2. 若 `wordInstalled=false` 或 `comAvailable=false` 或 `domAvailable=false`：
   - Stop the automation-heavy plan.
   - Tell the user to install desktop Microsoft Word.
   - Explain that WPS and similar tools often distort styles, pagination,
     captions, tables of contents, and cross-references.
3. 只有在 Word COM/DOM 可用时，才进入自动化文档编辑流程。

## 2. Formatting Rules

1. If the user provides a school template, style guide, screenshots, or page
   header/footer rules, follow them strictly.
2. If the user does not provide rules, do not invent school-specific standards.
3. Control formatting through styles, not scattered manual formatting.
4. Reuse and repair existing styles before creating new ones.

## 3. Suggested Style Groups

- `Body Text`
- `Heading 1` / `Heading 2` / `Heading 3`
- `Figure Caption`
- `Table Caption`
- `Code Listing`
- `References`
- `Abstract`, `Keywords`, and `Appendix Title`

## 4. Editing Order

1. Read the user requirements and current document structure.
2. Inspect style names, fonts, spacing, indentation, and page-break behavior.
3. Map target styles to document styles.
4. Normalize body text, headings, captions, and references.
5. Then fix page numbers, table of contents, sections, references, and figure
   numbering.
6. Finish with a manual review for overflow, orphan lines, and figure/table
   page breaks.

## 5. Do Not Do This

- Do not mix heavy direct formatting with style-driven formatting.
- Do not invent school rules when the template evidence is missing.
- Do not treat WPS rendering as the final Word result.
