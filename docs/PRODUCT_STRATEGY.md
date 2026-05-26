# Product Strategy

## Positioning

`summarize-ppt-notes` is a local-first slide understanding pipeline for students, teachers, researchers, and technical teams who need more than a short PPT summary. The core promise is:

> Turn dense slides into explainable study notes while preserving every page, formula, image, and diagram trail.

## Target Users

- Students reviewing lecture decks with formulas, diagrams, and derivations.
- Teachers creating handouts from courseware.
- Researchers summarizing technical conference slides.
- Product and engineering teams turning internal decks into onboarding notes.
- Training companies generating consistent Word handouts from slide libraries.

## Open-Core Boundary

Keep free and open source:

- local PPT/PDF extraction
- Word generation
- JSON templates
- quality reports
- Codex Skill workflow
- basic chart/image/formula candidate extraction

Potential paid layers:

- OCR and math recognition for screenshot-only equations
- hosted batch processing
- branded Word templates
- model-powered automatic explanations
- team review workflow
- private deployment for schools or companies
- API access for LMS or knowledge-base systems

## Packaging

Recommended public repo structure:

- `README.md`: product-level pitch and quick start
- `SKILL.md`: Codex Skill runtime instructions
- `scripts/`: deterministic local extraction and DOCX generation
- `references/`: schema and agent instructions
- `examples/`: small safe examples only
- `tests/`: regression tests for extraction and quality reports
- `.github/`: CI and contribution workflow

## Differentiation

- Local-first: users can inspect all extracted files before any AI step.
- Evidence-preserving: full-slide screenshots stay inside the Word output.
- Formula-aware: formula candidates trigger required explanations and examples.
- Word-native equations: common formulas are written as Office Math / OMML instead of raw LaTeX text.
- Quality-gated: missing explanation fields are caught before delivery.
- Skill-native: Codex can use the same repository as both documentation and automation.
- Markdown-native review: final notes and prompt packs can be reviewed directly on GitHub.
- Teacher-like guidance: the output starts with a chapter-style learning path so students know what to study first and when to move on.
- Practice loop: students get short exercises with answer and solution steps instead of only reading notes.
- Adaptive loop: wrong-answer JSON can regenerate a targeted review path.
- Local HTML review: a self-contained `study_index.html` gives a more product-like learning surface.
- Final-exam loop: the output includes active recall, flashcards, formula sheets, a mistake log, and a cram plan.

## Pricing Hypothesis

- Free: CLI + Codex Skill + local extraction.
- Pro: hosted OCR/math recognition, branded templates, batch queue, export themes.
- Team: shared projects, review workflow, template governance, private storage.
- Enterprise/Education: private deployment, SSO, audit logs, LMS integration.

## Star-Growth Loop

1. Make the README demo immediately understandable.
2. Keep the free local tool useful without signup.
3. Publish before/after examples using non-private decks.
4. Encourage users to contribute extraction edge cases.
5. Turn repeated edge cases into visible quality improvements.

## Near-Term Roadmap

1. Expand OMML coverage for matrices, aligned equations, piecewise functions, and integrals.
2. Add production OCR/math recognition adapters.
3. Add `.dotx` or template-based Word styling.
4. Add optional open-education web search with source attribution and license filtering.
5. Add benchmark decks with expected extraction snapshots.
6. Add a small local review UI for filling `notes_template.json`.
7. Add LMS export and Anki package (`.apkg`) generation.
