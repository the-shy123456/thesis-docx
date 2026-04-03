---
name: thesis-docx
description: Create, revise, and format thesis or dissertation Word documents with strict academic formatting control. Use when Codex needs to generate or revise thesis content, normalize Word styles, follow a school template, fix captions or page numbers or section levels, or produce evidence-based Mermaid figures and LaTeX-formatted code listings for a thesis document.
---

# Thesis DOCX

## Overview

Use this skill for thesis-oriented `.docx` work where content quality and
format fidelity both matter. Prefer Microsoft Word desktop automation over
WPS-like alternatives whenever the task involves batch formatting, styles,
captions, pagination, tables of contents, or cross-references.

## Workflow

1. Check Microsoft Word and COM/DOM automation first.
   - On Windows, run:
   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts/check_word_com.ps1 -Json
   ```
   - If Word is missing, COM is unavailable, or DOM access fails:
     - Stop the automation-heavy plan.
     - Tell the user to install desktop Microsoft Word.
     - Explain briefly that WPS or similar tools are likely to degrade layout
       fidelity for thesis formatting.
2. Read the user's real constraints before editing.
   - Collect the thesis template, school formatting guide, screenshots,
     existing document, sample pages, and any explicit chapter rules.
   - If the user gave formal requirements, follow them strictly.
   - If the user did not give requirements, do not invent school-specific
     standards. Use conservative academic defaults and say that they are
     defaults.
3. Standardize with styles, not scattered direct formatting.
   - Reuse and repair existing styles whenever possible.
   - Create missing styles only when no matching style exists.
   - Keep body text, headings, figure captions, table captions, references,
     abstract, and appendix styles separate and consistent.
4. Edit content only from user-provided facts.
   - Expand, polish, or reorganize thesis text only within the user's topic,
     evidence, codebase, notes, or source material.
   - Do not invent experimental data, system structure, entities, or results.
5. Generate figures only when the source material is sufficient.
   - Use Mermaid for architecture diagrams, E-R diagrams, flow charts, state
     diagrams, and similar thesis figures.
   - Base every node, field, relation, and dependency on real materials from
     the user, such as code, SQL schema, API docs, project docs, or the thesis
     draft itself.
   - If the materials are insufficient, refuse to fabricate the figure and ask
     for the missing source information.
6. Typeset code with LaTeX conventions when thesis code excerpts are needed.
   - Keep only the code relevant to the argument.
   - Preserve real identifiers from the user's code or design.
   - Avoid synthetic filler code written only to look complete.

## Style Strategy

- Treat styles as the single source of truth for formatting.
- Prefer these logical style buckets:
  - `Body Text`
  - `Heading 1` / `Heading 2` / `Heading 3`
  - `Figure Caption`
  - `Table Caption`
  - `References`
  - `Abstract`
  - `Keywords`
  - `Appendix Title`
- If the document already has equivalent styles, map to them and normalize
  their font, spacing, indentation, and numbering behavior.
- If direct formatting conflicts with styles, reduce the direct formatting and
  bring the document back under style control.

## Figure Rules

- Use Mermaid when a thesis figure is needed and the structure can be traced to
  real source material.
- Keep the diagram academically neutral and concise.
- Avoid decorative labels, chatty callouts, and speculative entities.
- Match the user's terminology unless it conflicts with the real materials.
- Read `references/figure-and-code-rules.md` before generating diagrams.

## Code Listing Rules

- Prefer LaTeX-oriented code presentation when the thesis includes code
  listings.
- Keep the listing faithful to the actual code.
- Trim non-essential boilerplate when it does not support the thesis argument.
- If the user requests a specific LaTeX package or listing style, follow it.

## Resource Guide

- `scripts/check_word_com.ps1`
  - Detect whether Microsoft Word desktop and COM/DOM automation are available.
- `scripts/normalize_word_styles.ps1`
  - Batch-normalize thesis body text, Heading 1-3, figure captions, and table
    captions through Word COM automation.
- `scripts/render_mermaid_figure.ps1`
  - Render Mermaid source into thesis-ready SVG, PNG, or PDF figure assets.
- `references/paper-format-workflow.md`
  - Read for the standard Word thesis formatting workflow.
- `references/figure-and-code-rules.md`
  - Read before generating Mermaid figures or LaTeX code listings.
- `references/script-usage.md`
  - Read for command examples and config file conventions.

## Final Checks

- Confirm the document is still style-driven after edits.
- Confirm captions, numbering, page breaks, and table of contents are coherent.
- Confirm every diagram and code block is grounded in user-provided material.
- If Word automation was unavailable, explicitly warn that layout fidelity was
  not guaranteed.
