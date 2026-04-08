# Thesis Voice and Style

Use this reference before generating, rewriting, or expanding thesis prose.

## 1. What Good Thesis Voice Looks Like

Preferred tone:

- neutral
- academic
- concise
- declarative
- readable as student thesis writing

Preferred patterns:

- `本文设计并实现了……`
- `系统采用了……`
- `在该模块中……`
- `为了实现……，系统……`
- `实验结果表明……`
- `综上所述……`

## 2. What Must Not Appear in the Thesis

Do not leak the writing workflow, prompting process, or source-feeding process
into the final paper text.

Forbidden expressions include:

- `根据已有工程……`
- `根据现有代码……`
- `根据用户提供的代码……`
- `结合任务书和代码……`
- `通过分析现有项目……`
- `将代码和任务书输入模型后……`
- `通过 AI 分析得出……`
- any sentence that reads like an explanation of how the assistant reasoned
  rather than how the system works

These belong to the hidden working process, not to the thesis itself.

## 3. Distinguish Source Basis From Meta Talk

Allowed:

- use the real project, codebase, schema, and task book as factual sources
- write system descriptions that are faithful to those sources

Not allowed:

- mention the feeding, parsing, inference, or reconstruction process
- narrate the existence of the AI assistant inside the thesis body

Bad:

- `根据现有代码可以看出系统采用了 Spring Boot。`

Good:

- `系统后端采用 Spring Boot 框架进行开发。`

Bad:

- `结合任务书与现有工程，本文确定了系统模块划分。`

Good:

- `本文围绕校园二手交易场景，对系统模块进行了划分。`

## 4. Student Thesis Voice

The writing should sound like a student presenting their own design and
implementation, not like a reviewer explaining how the draft was produced.

Preferred:

- `本文`
- `本系统`
- `该模块`
- `本研究`

Avoid:

- review-style instructions
- assistant-style disclaimers
- prompt-style meta statements
- exaggerated marketing copy

## 5. Revision Rule

When rewriting a paragraph:

1. remove workflow/meta wording first
2. rewrite into direct thesis narration
3. keep only the factual claim, design choice, or result
4. preserve real technical facts from the source material

If a sentence cannot be written cleanly without referring to the hidden source
analysis process, rewrite the sentence from the system perspective instead of
the writing-process perspective.

## 6. Final Self-Check Before Output

Before treating a thesis paragraph as finished, ask:

1. Is this sentence about the system, module, design choice, experiment, or
   result itself?
2. Or is it secretly explaining how the assistant learned that fact?

If it is the second kind, rewrite it.

Quick rejection checklist:

- contains `根据现有代码`
- contains `根据已有工程`
- contains `根据用户提供的代码`
- contains `结合任务书和代码`
- contains `通过分析现有项目`
- contains `通过 AI 分析`
- contains wording that sounds like prompt notes or hidden reasoning

Only deliver the thesis text after this checklist passes.
