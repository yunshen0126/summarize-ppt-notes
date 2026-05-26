# Open Source Comparison

This project should not compete as another generic PPT-to-Markdown converter. The strongest position is a slide-to-study-notes pipeline: preserve every page as evidence, extract structured signals, guide an AI/human reviewer through missing explanations, then export Word and Markdown notes.

## Comparison Matrix

| Project | Strength | Gap for study-note use | Our response |
| --- | --- | --- | --- |
| Microsoft MarkItDown | Broad document-to-Markdown conversion, useful for LLM ingestion, plugin-friendly ecosystem | General converter; not a page-by-page teaching-note workflow with Word equations, formula examples, learning route, practice loop, and quality gates | Keep Markdown output, but focus on slide pedagogy, screenshots, Word equations, formulas, examples, learning path, practice questions, and quality reports |
| pptx2md | Focused PPTX-to-Markdown converter with images and slide text extraction | Primarily conversion; no Word notes, no explanation schema, no formula/example quality gate, no teacher-style study path, no exercise pack | Provide Markdown plus DOCX, prompt pack, Office Math output, per-slide explanation schema, learning-path module grouping, practice questions, and validation |
| Unstructured | Mature document partitioning and ETL-style elements for PPT/PDF pipelines | Excellent extraction layer, but not a finished study-note product | Use similar structured-output thinking while keeping a simple local CLI and Codex Skill |
| Docling | Strong local document parsing, OCR/VLM-oriented roadmap, multiple exports | Broad document intelligence platform; may be heavier than a slide-note workflow | Stay lightweight and workflow-specific; add optional adapters later |

## Differentiation To Protect

- Evidence-first output: every slide keeps a screenshot inside the generated notes.
- Pedagogy-first schema: every slide requires purpose, explanation, visual interpretation, formula explanation, and examples.
- Word-native equations: common LaTeX formulas are emitted as Office Math / OMML objects.
- Quality gate: `quality_report.md` makes missing explanations visible.
- Teacher-style route: `learning_path.md` turns a deck into chapter-like modules with goals, must-read slides, checkpoints, and traps.
- Practice pack: `practice_questions.md` gives short questions with answers and solution steps.
- Adaptive review: `adaptive_review.md` can be regenerated from wrong-answer JSON.
- Local review page: `study_index.html` makes the pack easier to navigate.
- Final-exam pack: active recall, formula sheet, flashcards, mistake log, cram plan, one-page review, and concept map are generated from the same notes.
- Local-first extraction: source decks are processed on the user's machine.
- Skill-native operation: Codex can use the same repo as an executable workflow, not just a library.

## Feature Gaps Worth Closing Next

1. Production OCR/math-recognition engines behind the existing `--ocr-json` merge interface.
2. Broader LaTeX-to-OMML coverage for matrices, aligned equations, piecewise functions, and integrals.
3. Benchmark fixtures that snapshot extraction output for representative PPTX/PDF files.
4. Themeable Word templates for schools, teachers, and teams.
5. Optional ingestion adapters for MarkItDown, Unstructured, and Docling outputs.

## Product Decision

Do not chase every file format. The open-source hook should be:

> The best local-first pipeline for turning dense slide decks into explainable study notes.
