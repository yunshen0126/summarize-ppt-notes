# Changelog

## 0.11.0

- Add default `--review-depth compressed` reading plan that separates study slides into `必读深讲`, `快速扫读`, and `参考/重复`.
- Add repeated-slide detection and `10_阅读取舍.md` so students know what to read first, what to skim, and what can wait.
- Reduce prompt and notes volume by only requiring note objects for deep/quick slides; reference slides stay traceable in `extraction.json`.
- Compact quick-scan slides in Word/Markdown output and lower quality checks for quick slides while preserving strict checks for deep slides.
- Update README, skill instructions, schema, and quality standards to emphasize information-gain compression over mechanical per-slide expansion.

## 0.10.0

- Add default `--content-filter study` mode that compacts title, agenda/table-of-contents, section divider, and closing slides.
- Keep compacted navigation slides in `extraction.json` while excluding them from notes templates, Word/Markdown study notes, prompt packs, learning paths, exercises, flashcards, formula sheets, HTML study pages, and quality scoring.
- Add `--content-filter all` for full per-slide audit output and report skipped navigation slides in `START_HERE.md` and quality reports.

## 0.9.0

- Rewrite README in Chinese with full usage for Word equations, OCR JSON, practice banks, adaptive review, HTML study page, and benchmarks.
- Write formulas as real Word Office Math / OMML objects for common LaTeX structures.
- Add `study_index.html`, `adaptive_review.md`, `wrong_answer_template.json`, OCR JSON merge, practice-bank JSON matching, and benchmark runner.

## 0.8.0

- Improve Word formula display so common LaTeX appears as readable formula text instead of raw source code.
- Add `practice_questions.md` and `practice_questions.json` with short exercises, answers, solution steps, difficulty, and source notes.
- Surface `07_小题练习.md` in clean deliverables and update START_HERE to include the practice loop.

## 0.7.0

- Add `learning_path.md`, a teacher-style chapter route with module goals, must-read slides, self-test checkpoints, and common traps.
- Surface `00_学习路径.md` first in the clean deliverables folder and `START_HERE.md`.
- Update study-pack dashboard and cram plan to start from the learning path before detailed reading.

## 0.6.0

- Add teacher output profile with a clean `*_deliverables/START_HERE.md` entry point.
- Copy the few student-facing files into a simple deliverables folder.
- Hide raw extraction artifacts from default CLI output while preserving complete/debug modes.

## 0.5.0

- Redesign default DOCX output into a polished study-handout layout.
- Add `--layout study|audit`; study mode suppresses raw extraction clutter, audit mode keeps it.
- Add anti-fluff quality checks and denser minimum explanation requirements.
- Add output quality documentation for deep, specific final-exam notes.

## 0.4.0

- Add final-exam study-pack generation.
- Add Anki CSV flashcards, active-recall questions, formula sheet, cram plan, mistake-log template, one-page review, and Mermaid concept map.
- Add final-exam schema fields and `--study-mode final` quality checks.
- Add final-exam workflow documentation.

## 0.3.0

- Add final study-notes Markdown output for GitHub preview and review workflows.
- Add `prompt_pack.md` generation for LLM/VLM-assisted slide review.
- Add open-source comparison documentation and sharpen product differentiation.

## 0.2.0

- Add GitHub-ready project metadata, README, license, CI, and tests.
- Add quality reports with optional `--fail-under` quality gate.
- Add slide selection through `--slides` and `--max-slides`.
- Add chart/diagram/embedded-object relationship extraction.
- Add `doctor.py` for local dependency checks.

## 0.1.0

- Initial Codex Skill and local PPT/PDF to Word notes exporter.
