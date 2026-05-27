# Output Quality Standard

This project optimizes for useful final-exam review, not long documents.

## Good Notes

Good notes are:

- specific: they cite the actual formula, theorem, algorithm, diagram, or table on the slide
- deep: they explain why a result works, not only what it says
- operational: they show how to solve a question or check an answer
- compact: they remove generic transitions and empty advice
- compressed: they separate must-read, quick-scan, and reference/repeated slides instead of expanding every page equally
- exam-aware: they name likely question types and common mistakes
- guided: they tell students which module to study first, which slides are must-read, and when they are ready to move on
- actionable: they include short practice questions with answers and solution steps
- word-native: formulas in Word should be equation objects when the converter can represent them safely

## Bad Notes

Avoid lines like:

- "put this slide back into the chapter logic"
- "first understand the problem"
- "review this carefully"
- "this page is important"
- "look at the screenshot"

These sentences are only acceptable if they are followed by concrete slide-specific content.

## Required Explanation Pattern

For a theorem or formula:

1. State the object being measured or bounded.
2. State assumptions and units.
3. Explain the intuition.
4. Work through a small numeric example.
5. Name the common trap.

For an algorithm:

1. Input.
2. Output.
3. Ordered steps.
4. Why the key step is valid.
5. How to verify the final answer.

For a chart or diagram:

1. Identify axes/nodes/regions/arrows.
2. Explain what changes and what stays fixed.
3. State the conclusion the slide wants the student to learn.

For a learning path:

1. Group adjacent slides into chapter-like modules.
2. Name the learning goal and must-read slides.
3. Give closed-book checkpoints.
4. State common traps.
5. Avoid duplicating the full notes.

For reading compression:

1. Preserve formulas, definitions, examples, diagrams, conditions, and exam traps.
2. Mark low-increment or repeated pages as reference instead of asking for full explanations.
3. Keep quick-scan slides to 2-4 bullets and one exam signal.
4. Use the original screenshots and extraction JSON as traceability, not as material to rewrite.
5. A student should know exactly what to read first and what can wait.

For practice questions:

1. Keep each question short enough to solve in a few minutes.
2. Include final answer and solution steps.
3. Tie the question to a slide concept, formula, diagram, or common mistake.
4. Prefer original questions generated from the PPT.
5. If using web material, only use open or clearly attributable sources, paraphrase/adapt the question, and keep the source link.

For adaptive review:

1. Group wrong answers by slide and root cause.
2. Recommend a short re-study order instead of telling students to reread everything.
3. Turn each repeated error into a next-action rule.
4. Preserve a fillable JSON template when no wrong-answer data is supplied.
