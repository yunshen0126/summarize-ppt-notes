# Notes JSON Schema

Create a JSON file with this structure, then pass it to `ppt_notes_exporter.py` with `--notes-json`.

```json
{
  "language": "zh",
  "slides": [
    {
      "number": 1,
      "title": "可选标题",
      "purpose": "这一页是干什么用的。",
      "what_it_says": "这一页讲了什么，按内容完整说明。",
      "detailed_explanation": "复杂概念、流程、图表、算法或推导的详细讲解。",
      "visual_explanation": "图片、图表、流程图、架构图、代码截图等视觉元素说明。",
      "formula_explanations": [
        {
          "formula": "E = mc^2",
          "meaning": "说明每个符号和公式含义。",
          "conditions": "适用条件或前提；没有就写“无特殊条件”。",
          "example": "给出一个使用例子。"
        }
      ],
      "worked_examples": [
        "如果这一页有复杂公式或方法，在这里写完整例题或小案例。"
      ],
      "exam_focus": "这一页在期末考试中可能怎么考，属于概念题、计算题、推导题、图表题还是综合题。",
      "key_takeaways": [
        "必须背会、会解释、会套用的核心点。"
      ],
      "memory_hooks": [
        "帮助学生记忆的口诀、类比、步骤压缩或对比记忆。"
      ],
      "likely_questions": [
        {
          "question": "闭卷自测题或考试风格题。",
          "answer": "标准答案或答题要点。"
        }
      ],
      "practice_questions": [
        {
          "question": "基础小题，最好能直接动笔做。",
          "answer": "最终答案。",
          "solution": "分步骤解题思路。",
          "difficulty": "基础 / 中等 / 困难",
          "source": "PPT 内容原创生成 / 开放来源改编",
          "source_url": "如果参考了网上开放资料，填写链接；原创题留空。"
        }
      ],
      "common_mistakes": [
        "学生容易犯的错误、混淆点、漏写条件、计算陷阱。"
      ],
      "prerequisites": [
        "理解这一页前需要掌握的前置知识。"
      ],
      "difficulty": "基础 / 中等 / 困难",
      "estimated_review_minutes": "建议复习分钟数，例如 8",
      "tags": [
        "章节名",
        "知识点名"
      ],
      "uncertainties": [
        "如果某个公式、图片或文字无法准确识别，在这里标记需核对。"
      ]
    }
  ]
}
```

Guidelines:

- Use one object per slide and keep `number` aligned with the slide/page number.
- Keep formulas in LaTeX where possible. If the visual form is ambiguous, preserve the original visible text and add `需核对`.
- `formula_explanations` may be an empty list only when the slide has no formula or formula-like expression.
- `worked_examples` should not be empty for slides containing complex formulas, derivations, algorithms, or multi-step calculations.
- Do not omit non-text content. Put image/chart/diagram interpretation in `visual_explanation`.
- The quality report expects substantial content in `purpose`, `what_it_says`, and `detailed_explanation` for every slide.
- If the extraction reports formulas, fill both `formula_explanations` and `worked_examples`.
- If the extraction reports images, tables, charts, diagrams, or embedded objects, fill `visual_explanation`.
- For final-exam use, fill `exam_focus`, `key_takeaways`, `likely_questions`, and `common_mistakes` for every slide.
- `likely_questions` should include active-recall questions. For formulas, include at least one calculation or application question.
- `practice_questions` should be short enough for students to solve quickly and must include answer plus solution steps. Prefer original questions generated from the slide. If using online material, use open/clearly attributable sources, paraphrase or adapt, and keep `source_url`.
- `memory_hooks` should be short and useful, not decorative. Prefer contrast tables, acronyms, or "first do X, then do Y" patterns.
- Do not use generic filler such as “放回主线理解”, “先明确解决的问题”, “建议仔细理解”, or “复习时按三步走”.
- `detailed_explanation` should be high-density: define the concept, state conditions, explain why it works, show how to use it, and name the common mistake.
- For formulas, include log base/unit/condition when relevant and at least one numeric mini-example.
- For algorithms, include input, output, ordered steps, correctness intuition, and final-answer check.
