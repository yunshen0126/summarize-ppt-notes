# Final Exam Workflow

This workflow turns a slide deck into a repeatable review loop instead of a passive summary.

## 1. Convert Slides Into Evidence

Run the extractor first. Keep screenshots, text, formulas, image metadata, chart signals, and speaker notes. Screenshots matter because formulas and diagrams are often stored as images or shapes instead of text.

## 2. Fill The Teaching Notes

For each slide, fill:

- purpose
- what_it_says
- detailed_explanation
- visual_explanation
- formula_explanations
- worked_examples

This creates understanding.

## 3. Fill The Exam Fields

For each slide, fill:

- exam_focus: how it may be tested
- key_takeaways: what must be recalled from memory
- likely_questions: active-recall and exam-style questions
- common_mistakes: traps and confusing points
- memory_hooks: concise memory supports
- difficulty and estimated_review_minutes

This creates exam readiness.

## 4. Use The Study Pack

Recommended order:

1. Read `learning_path.md` to follow the chapter-style route instead of treating every slide equally.
2. Open the Word handout only for the current learning-path module.
3. Answer `active_recall_questions.md` without opening the slides.
4. Do `practice_questions.md` before reading the answers and solution steps.
5. Open `study_index.html` when a browser-based review surface is easier than reading separate Markdown files.
6. Rework formulas from `formula_sheet.md`.
7. Import `flashcards_anki.csv` into Anki or use `flashcards.md`.
8. Record every missed question in `mistake_log_template.md` or `wrong_answer_template.json`.
9. Regenerate with `--wrong-answers-json` and review `adaptive_review.md`.

## 5. Three-Pass Review

- Pass 1: Understand. Read final notes and explain each slide aloud.
- Pass 2: Recall. Use questions and flashcards without looking.
- Pass 3: Exam mode. Solve formula examples, explain diagrams, and review mistakes.

The goal is to reduce rereading and increase retrieval practice.
