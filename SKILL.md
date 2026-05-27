---
name: summarize-ppt-notes
description: Create detailed Word/Markdown/HTML study notes and final-exam review packs from PowerPoint decks or slide PDFs, including full-slide screenshots, extracted text, real Word Office Math equations, formulas, tables, images, charts, content-slide explanations, token-saving title/agenda/section slide compacting, teacher-style chapter learning paths, active-recall questions, small practice exercises with answers/solution steps, adaptive wrong-answer review paths, OCR/math-recognition JSON merge, practice-bank JSON matching, formula sheets, Anki flashcards, mistake-log templates, cram plans, and concept maps. Use when asked to summarize PPT/PPTX/PDF courseware, lecture slides, presentation notes, formulas, diagrams, or generate Chinese期末复习 materials that explain what each slide is for, what it says, complex content, what to learn first, likely exam questions, common mistakes, and worked examples for complex formulas.
---

# Summarize PPT Notes

## Overview

Use this skill to turn a slide deck (`.pptx`, `.ppt`, or slide `.pdf`) into a Word/Markdown/HTML study-notes document and a final-exam review pack. The expected output is a readable teacher-style review system, not a giant transcript. Default behavior should compress the deck like an information-theory problem: preserve high-yield definitions, formulas, examples, traps, exam signals, and diagrams; remove repeated wording, decorative transitions, and low-increment slides.

By default, compact low-value navigation slides instead of expanding them into full notes: title/cover pages, agenda/table-of-contents pages, section dividers, and closing/Q&A pages. They stay in `extraction.json` for traceability, but should not consume Word pages, practice questions, learning-path modules, flashcards, or prompt budget. Use `--content-filter all` only when the user explicitly needs full per-slide audit output.

Also by default, use `--review-depth compressed`. This creates three reading tiers:

- `必读深讲`: pages that deserve real explanation because they contain formulas, core definitions, examples, diagrams, or exam-heavy content.
- `快速扫读`: pages that add context or a small new conclusion; explain in 2-4 bullets, not long paragraphs.
- `参考/重复`: repeated or low-increment pages; list them in the reading index but do not request full note objects.

Use `--review-depth complete` only when the user explicitly asks for exhaustive per-slide notes.

Chinese is the default output language unless the user asks otherwise.

## Workflow

1. Locate the source slide file and choose an output path. If the user does not specify one, create the Word document next to the source file or on the Desktop.
2. Run `scripts/doctor.py` if this is the first use in the environment. Missing LibreOffice or PyMuPDF means screenshots may not render, but XML extraction can still work.
3. Run `scripts/ppt_notes_exporter.py` to create screenshots, extract raw slide content, generate a notes JSON template, and produce an initial Word skeleton:

   ```bash
   python3 /path/to/summarize-ppt-notes/scripts/ppt_notes_exporter.py "/path/to/slides.pptx" --output "/path/to/PPT学习笔记.docx"
   ```

4. Inspect the generated `extraction.json`, `prompt_pack.md`, `quality_report.md`, `notes_template.json`, `10_阅读取舍.md`, and slide screenshots in the work directory. Check `skipped_navigation_slides` and reading tiers. Title/agenda/section pages should usually be compacted; repeated reference pages should not be expanded; formula/chart/example/exercise pages should usually be `必读深讲` or `快速扫读`. Do not rely only on XML text extraction; visually inspect screenshots to catch formulas rendered as images, charts, SmartArt, handwritten symbols, and text embedded in pictures.
5. Fill a notes JSON file using the schema in `references/note-schema.md`. Fill only note-required slides in `notes_template.json`, not every original slide. For `必读深讲` slides, write:
   - `purpose`: this slide's role in the lecture/presentation.
   - `what_it_says`: a faithful explanation of the visible content.
   - `detailed_explanation`: deeper explanation for dense definitions, algorithms, diagrams, derivations, tables, or charts.
   - `visual_explanation`: what each important image, chart, diagram, or screenshot is showing and why it matters.
   - `formula_explanations`: every important formula, preferably in LaTeX, with variable meanings and assumptions.
   - `worked_examples`: concrete examples for complex formulas or procedures.
   - `exam_focus`: how this slide may appear in a final exam.
   - `key_takeaways`: the points the student must be able to recall without notes.
   - `likely_questions`: active-recall and exam-style questions with answer hints.
   - `practice_questions`: short exercises with answer, solution steps, difficulty, and source/source_url when adapted from open web material.
   - `common_mistakes`: traps, confusing pairs, missing conditions, and calculation pitfalls.
   - `memory_hooks`: concise memory aids, contrast rules, or step patterns.
   For `快速扫读` slides, keep the same fields but write concise content: 2-4 key bullets, one exam signal, one mistake if relevant, and a short explanation only when a formula/diagram requires it.
6. Re-run the script with the filled notes JSON to build the final Word document. Use `--ocr-json` when external OCR/math recognition has been produced, `--practice-bank-json` when an open/self-owned question bank is available, and `--wrong-answers-json` when the student has filled a wrong-answer log. Use the default teacher profile so the user sees one clean `START_HERE.md` entry instead of many intermediate files:

   ```bash
   python3 /path/to/summarize-ppt-notes/scripts/ppt_notes_exporter.py "/path/to/slides.pptx" --notes-json "/path/to/notes_filled.json" --output "/path/to/PPT学习笔记.docx" --notes-markdown "/path/to/PPT学习笔记.md" --study-mode final --layout study --output-profile teacher --fail-under 85
   ```

7. Verify the final `*_deliverables/START_HERE.md`, `10_阅读取舍.md`, `00_学习路径.md`, `07_小题练习.md`, `08_学习页面.html`, and `09_错题反馈路径.md` exist. The final answer should point the user to START_HERE first, then `10_阅读取舍.md`, then `00_学习路径.md`; mention debug/work directories only as optional. Check that slide count, reading tiers, screenshots, formulas, chart/diagram objects, extracted images, teacher-style learning path, active-recall questions, practice questions with answers/solution steps, formula sheet, flashcards, concept map, HTML study page, adaptive review path, and mistake-log template are represented. If any formula or visual element cannot be read confidently, mark it clearly as `需核对` and explain what is uncertain.

## Slide Analysis Standard

For each note-required study slide, cover the following. Do not write full notes for `参考/重复` slides.

- **页面截图**: include the rendered full-slide image when available.
- **原始内容**: preserve visible text, bullet points, tables, notes, formula candidates, and extracted images.
- **这一页是干什么用的**: explain the slide's teaching/presentation function, such as introducing a definition, proving a result, comparing methods, showing an example, or summarizing conclusions.
- **这一页讲了什么**: restate the content in clear Chinese without losing technical terms.
- **复杂内容详解**: expand dense logic step by step only for `必读深讲` slides. For `快速扫读` slides, compress to the minimum explanation needed to avoid misunderstanding.
- **公式说明和例题**: list formulas, define variables, state conditions, explain intuition, then give a small numeric or conceptual example when the formula is nontrivial.
- **期末复习字段**: identify likely exam forms, key takeaways, common mistakes, active-recall questions, memory hooks, and estimated review time.
- **学习路径字段**: use titles, tags, prerequisites, difficulty, formulas, visuals, examples, and exam focus to infer chapter-like modules, must-read pages, and self-test checkpoints.
- **小题练习字段**: create short exercises that can be solved quickly, with answer and solution steps. Prefer original questions generated from the PPT; if matching online examples, use open/clearly attributable sources, paraphrase or adapt instead of copying, and include `source_url`.
- **错题反馈字段**: if wrong-answer JSON is available, reorder review by missed slide and root cause; otherwise generate a fillable template and high-risk slide list.

## Writing Quality Standard

Avoid filler. The following patterns are not acceptable unless followed by exact slide-specific content:

- “放回主线理解”
- “先明确解决的问题”
- “建议仔细理解”
- “复习时按三步走”
- “本页需要理解”

Each `detailed_explanation` must include the actual concept or formula from the slide and follow this pattern when applicable:

1. Define the object or problem.
2. State the condition or assumption.
3. Explain why the theorem/formula/algorithm works.
4. Show how to use it on a small example.
5. Name the common mistake.

For formulas, do not stop at “说明变量含义”. Include log base/unit, condition, one numeric example, and the interpretation of the result.

For algorithms, include input, output, step order, why the greedy/recursive/iterative step is valid, and how to check the final answer.

For diagrams, name nodes, arrows, axes, regions, or table columns. Do not write generic “看图理解”.

## Compression Standard

Default outputs should be short enough that a student wants to read them.

- Treat repeated slides as redundancy. Merge their value into the nearest high-yield slide instead of writing another full page.
- Use information gain: if a slide adds no new definition, formula, condition, example, diagram conclusion, or exam trap, mark it as reference.
- For a 100-page deck, the expected result is usually around 15-25 `必读深讲` pages, 20-45 `快速扫读` pages, and the rest as reference/repeated pages, unless the deck is unusually dense.
- Use “核心结论 -> 为什么 -> 怎么考 -> 易错点” as the default explanation shape.
- Do not duplicate the original slide text. The original remains in screenshots and `extraction.json`; the notes should teach what the student must understand.
- Prefer short, concrete examples over long conceptual paragraphs.
- The student-facing order is `START_HERE.md` -> `10_阅读取舍.md` -> `00_学习路径.md` -> current module in `01_复习讲义.docx` -> active recall/practice.

## Formula Handling

- Treat formulas in three sources as important: Office Math/OMML extracted from `.pptx`, formula-like text detected by regex, and formulas visible only in screenshots or images.
- Keep formulas in LaTeX in JSON/Markdown when useful, but the Word output should write common formulas as Office Math / OMML equation objects instead of making raw LaTeX the primary view.
- Explain each variable and symbol. Do not skip constants, subscripts, superscripts, summations, integrals, matrix dimensions, or probability conditions.
- For a complex formula, include at least one worked example. A worked example can be numeric, symbolic, or a small scenario, but it must show how the formula is used.
- If formula recognition is uncertain, write `需核对` beside the formula and describe the ambiguity.

## Visual Content Handling

- Describe every non-decorative image, chart, table, graph, flowchart, architecture diagram, code screenshot, or dataset screenshot.
- For charts, mention axes, labels, units, trends, outliers, and the conclusion the slide likely wants the audience to draw.
- For diagrams, explain each part and relationship. For arrows, identify source, target, and meaning.
- For extracted pictures, include the image in the Word document when the script can extract it. If an image cannot be extracted separately, rely on the full-slide screenshot and describe it from the screenshot.

## Quality Checks

Before finalizing:

- There is a clean student-facing `START_HERE.md` that says what to read first, what is optional, and what is only for debugging.
- `10_阅读取舍.md` exists and separates slides into must-read, quick-scan, and reference/repeated.
- `00_学习路径.md` exists and tells the learner the chapter-style order, must-read slides, goals, checkpoints, and common traps.
- `07_小题练习.md` exists and contains question, answer, solution steps, difficulty, and source notes.
- `08_学习页面.html` exists as a local review page with module navigation, slide notes, formulas, and collapsible practice answers.
- `09_错题反馈路径.md` and `可选_错题输入模板.json` exist.
- Word and study-pack slide counts match the included study-slide count; `extraction.json` still preserves the original deck/PDF page count.
- Title, agenda, table-of-contents, section divider, and closing pages are compacted by default and not expanded into full explanations or practice questions unless the user requests `--content-filter all`.
- Every slide has a screenshot or a documented reason why rendering failed.
- No obvious formula, chart, table, or image is ignored.
- Dense slides have real explanation, not just rephrasing.
- Dense slides do not contain generic filler; they cite terms, formulas, algorithms, or examples visible on the slide.
- Complex formulas include examples.
- `quality_report.md` has no missing required fields for the intended delivery level.
- In `--study-mode final`, every slide has exam focus, key takeaways, likely questions, and common mistakes.
- `study_pack/` includes a usable learning path, cram plan, formula sheet, active-recall set, practice questions, HTML page, adaptive review path, flashcards, mistake log, one-page review, and concept map.
- The final answer to the user gives the Word document path and mentions any limitations or uncertain recognitions.

## Resources

- `scripts/ppt_notes_exporter.py`: extracts slide assets and generates Word documents.
- `scripts/doctor.py`: checks local runtime dependencies.
- `references/note-schema.md`: JSON schema for filled per-slide explanations.
- `benchmarks/run_benchmark.py`: quick regression benchmark for public fixtures.
