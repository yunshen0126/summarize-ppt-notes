---
name: summarize-ppt-notes
description: Create detailed Word/Markdown study notes and final-exam review packs from PowerPoint decks or slide PDFs, including full-slide screenshots, extracted text, formulas, tables, images, charts, per-slide explanations, active-recall questions, formula sheets, Anki flashcards, mistake-log templates, cram plans, and concept maps. Use when asked to summarize PPT/PPTX/PDF courseware, lecture slides, presentation notes, formulas, diagrams, or generate Chinese期末复习 materials that explain what each slide is for, what it says, complex content, likely exam questions, common mistakes, and worked examples for complex formulas.
---

# Summarize PPT Notes

## Overview

Use this skill to turn a slide deck (`.pptx`, `.ppt`, or slide `.pdf`) into a Word/Markdown study-notes document and a final-exam review pack. The expected output is not a brief summary: produce page-by-page notes with the original slide screenshot, complete extracted content, image/chart descriptions, formula recognition, teaching-style explanations, likely exam questions, common mistakes, memory hooks, and active-recall materials.

Chinese is the default output language unless the user asks otherwise.

## Workflow

1. Locate the source slide file and choose an output path. If the user does not specify one, create the Word document next to the source file or on the Desktop.
2. Run `scripts/doctor.py` if this is the first use in the environment. Missing LibreOffice or PyMuPDF means screenshots may not render, but XML extraction can still work.
3. Run `scripts/ppt_notes_exporter.py` to create screenshots, extract raw slide content, generate a notes JSON template, and produce an initial Word skeleton:

   ```bash
   python3 /path/to/summarize-ppt-notes/scripts/ppt_notes_exporter.py "/path/to/slides.pptx" --output "/path/to/PPT学习笔记.docx"
   ```

4. Inspect the generated `extraction.json`, `prompt_pack.md`, `quality_report.md`, `notes_template.json`, and slide screenshots in the work directory. Do not rely only on XML text extraction; visually inspect screenshots to catch formulas rendered as images, charts, SmartArt, handwritten symbols, and text embedded in pictures.
5. Fill a notes JSON file using the schema in `references/note-schema.md`. For every slide, write:
   - `purpose`: this slide's role in the lecture/presentation.
   - `what_it_says`: a faithful explanation of the visible content.
   - `detailed_explanation`: deeper explanation for dense definitions, algorithms, diagrams, derivations, tables, or charts.
   - `visual_explanation`: what each important image, chart, diagram, or screenshot is showing and why it matters.
   - `formula_explanations`: every important formula, preferably in LaTeX, with variable meanings and assumptions.
   - `worked_examples`: concrete examples for complex formulas or procedures.
   - `exam_focus`: how this slide may appear in a final exam.
   - `key_takeaways`: the points the student must be able to recall without notes.
   - `likely_questions`: active-recall and exam-style questions with answer hints.
   - `common_mistakes`: traps, confusing pairs, missing conditions, and calculation pitfalls.
   - `memory_hooks`: concise memory aids, contrast rules, or step patterns.
6. Re-run the script with the filled notes JSON to build the final Word document. Use a quality gate for serious work:

   ```bash
   python3 /path/to/summarize-ppt-notes/scripts/ppt_notes_exporter.py "/path/to/slides.pptx" --notes-json "/path/to/notes_filled.json" --output "/path/to/PPT学习笔记.docx" --notes-markdown "/path/to/PPT学习笔记.md" --study-mode final --fail-under 85
   ```

7. Verify the final Word, Markdown, and `study_pack/` outputs exist. Check that slide count, screenshots, formulas, chart/diagram objects, extracted images, active-recall questions, formula sheet, flashcards, concept map, and mistake-log template are represented. If any formula or visual element cannot be read confidently, mark it clearly as `需核对` and explain what is uncertain.

## Slide Analysis Standard

For each slide, cover the following:

- **页面截图**: include the rendered full-slide image when available.
- **原始内容**: preserve visible text, bullet points, tables, notes, formula candidates, and extracted images.
- **这一页是干什么用的**: explain the slide's teaching/presentation function, such as introducing a definition, proving a result, comparing methods, showing an example, or summarizing conclusions.
- **这一页讲了什么**: restate the content in clear Chinese without losing technical terms.
- **复杂内容详解**: expand dense logic step by step. For algorithms, describe input, output, process, and intuition. For diagrams/charts, explain axes, nodes, arrows, regions, trends, and takeaways.
- **公式说明和例题**: list formulas, define variables, state conditions, explain intuition, then give a small numeric or conceptual example when the formula is nontrivial.
- **期末复习字段**: identify likely exam forms, key takeaways, common mistakes, active-recall questions, memory hooks, and estimated review time.

## Formula Handling

- Treat formulas in three sources as important: Office Math/OMML extracted from `.pptx`, formula-like text detected by regex, and formulas visible only in screenshots or images.
- Convert formulas to LaTeX when possible. Keep the original visible form if conversion is uncertain.
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

- Slide count in the Word document matches the deck or PDF page count.
- Every slide has a screenshot or a documented reason why rendering failed.
- No obvious formula, chart, table, or image is ignored.
- Dense slides have real explanation, not just rephrasing.
- Complex formulas include examples.
- `quality_report.md` has no missing required fields for the intended delivery level.
- In `--study-mode final`, every slide has exam focus, key takeaways, likely questions, and common mistakes.
- `study_pack/` includes a usable cram plan, formula sheet, active-recall set, flashcards, mistake log, one-page review, and concept map.
- The final answer to the user gives the Word document path and mentions any limitations or uncertain recognitions.

## Resources

- `scripts/ppt_notes_exporter.py`: extracts slide assets and generates Word documents.
- `scripts/doctor.py`: checks local runtime dependencies.
- `references/note-schema.md`: JSON schema for filled per-slide explanations.
