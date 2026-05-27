# Summarize PPT Notes

把课程 PPT / PDF 变成能复习、能理解、能做题的学习资料包。

这个项目既是一个 Codex Skill，也是一个本地 CLI 工具。它的目标不是简单总结 PPT，也不是把 100 页课件扩写成几十万字，而是尽量像老师一样处理课件：先压缩重复内容，分清哪些页必须深讲、哪些页只要扫读、哪些页只是参考，再生成学习路径、练习题、公式表、错题反馈和 Word 讲义。

## 核心定位

很多 PPT 工具只做“转 Markdown”或“短摘要”。本项目更适合：

- 期末复习
- 公式密集课程
- 图表/流程图较多的课件
- 需要逐页讲解的课程材料
- 老师或助教批量生成讲义

核心承诺：

> 不是把 PPT 机械扩写成长文，而是把 PPT 压缩成一套学生愿意读、知道先看什么、能检查自己懂没懂的复习系统。

## 主要功能

- 支持 `.pptx`、`.ppt`、课件型 `.pdf`
- 生成完整 Word 复习讲义
- 每页保留原始截图，方便核对公式、图片和图表
- 提取 PPT 文字、备注、表格、图片、图表关系和公式候选
- Word 中使用真正的 Office Math / OMML 公式对象，而不是只显示 LaTeX 源码
- 默认压缩标题页、目录页、章节过渡页和结束页，只把正文学习页写进讲义、题目、学习路径和 prompt，减少无效输出与 token 消耗
- 默认使用 `--review-depth compressed`，按信息增量分成 `必读深讲`、`快速扫读`、`参考/重复` 三档，避免一页页重复扩写
- 自动生成 `10_阅读取舍.md`，告诉学生什么必须看、什么扫一眼、什么暂时不用看
- 生成分层讲解：必读页讲清楚“干什么、讲什么、复杂内容、图表、公式、例题”；快速页只保留核心结论、考试信号和易错点
- 自动生成 `00_学习路径.md`，按老师讲课思路组织章节式学习路线
- 自动生成 `07_小题练习.md`，包含题目、答案、解题思路、难度和来源说明
- 自动生成 `08_学习页面.html`，本地浏览器打开即可按模块复习
- 自动生成 `09_错题反馈路径.md` 和错题输入模板
- 支持 `--wrong-answers-json`，根据错题重新生成二次学习路径
- 支持 `--ocr-json`，接入外部 OCR / 数学公式识别结果
- 支持 `--practice-bank-json`，接入开放教育题库或自建题库
- 生成 Anki CSV、公式速查、主动回忆题、错题本模板、冲刺计划、一页纸总览、概念图
- 生成质量报告，检查哪些页解释不够深、哪些字段缺失
- 默认本地处理文件，脚本本身不会上传课件

## 快速开始

先检查环境：

```bash
python3 scripts/doctor.py
```

生成初稿：

```bash
python3 scripts/ppt_notes_exporter.py "slides.pptx" --output "PPT学习笔记.docx"
```

第一次运行会生成一个工作目录：

```text
PPT学习笔记_work/
├── assets/
├── converted/
├── screenshots/
├── study_pack/
├── extraction.json
├── extraction.md
├── notes_template.json
├── prompt_pack.md
├── quality_report.json
└── quality_report.md
```

## 默认省 token：先过滤导航页，再按信息增量压缩

默认 `--content-filter study` 会自动识别并压缩低价值导航页：

- 标题/封面页
- 目录、大纲、Agenda、Outline
- 章节过渡页
- Thank you / Q&A 等结束页

这些页面仍会保留在 `extraction.json`，但不会进入完整讲义、学习路径、练习题、公式表、Anki 卡片和 prompt pack。这样可以避免把“标题页也讲一大段”，也能减少后续让模型填讲解时的 token。

默认 `--review-depth compressed` 会继续把正文学习页分成三档：

```text
必读深讲：公式、核心定义、例题、图表、期末高频页
快速扫读：有少量新增信息，但不值得长讲
参考/重复：重复铺垫、低增量页，只保留在索引和 extraction.json
```

这样 100 页 PPT 通常不会再生成 100 页逐页长讲，而是先生成阅读取舍和学习路径。学生先看必读页，扫读快速页，参考页只在看不懂上下文或核对原图时打开。

如果你确实要逐页全量审计，可以加：

```bash
python3 scripts/ppt_notes_exporter.py "slides.pptx" \
  --content-filter all \
  --review-depth complete \
  --layout audit \
  --output "PPT全量审计.docx"
```

把 `notes_template.json` 填成高质量逐页讲解后，重新生成最终版：

```bash
python3 scripts/ppt_notes_exporter.py "slides.pptx" \
  --notes-json "PPT学习笔记_work/notes_template.json" \
  --output "PPT学习笔记.docx" \
  --notes-markdown "PPT学习笔记_work/final_notes.md" \
  --study-mode final \
  --layout study \
  --output-profile teacher \
  --fail-under 85
```

## 学生应该先看什么

默认会生成一个干净的交付目录：

```text
PPT学习笔记_deliverables/
├── START_HERE.md
├── 00_学习路径.md
├── 01_复习讲义.docx
├── 02_完整讲义.md
├── 03_一页纸总览.md
├── 04_主动回忆题.md
├── 05_公式速查.md
├── 06_错题本模板.md
├── 07_小题练习.md
├── 08_学习页面.html
├── 09_错题反馈路径.md
├── 10_阅读取舍.md
├── 可选_错题输入模板.json
├── 可选_Anki卡片.csv
├── 可选_冲刺计划.md
└── 质量检查.md
```

推荐顺序：

1. 打开 `START_HERE.md`
2. 看 `10_阅读取舍.md`，先知道什么必须看、什么扫读、什么不用先看
3. 看 `00_学习路径.md`，按老师讲课顺序进入章节
4. 打开 `01_复习讲义.docx`，只读当前模块对应页面
5. 合上讲义做 `04_主动回忆题.md`
6. 做 `07_小题练习.md`，先写答案，再看解题思路
7. 打开 `08_学习页面.html` 做折叠式复习
8. 把错题写入 `06_错题本模板.md` 或 `可选_错题输入模板.json`
9. 考前只看一页纸总览、公式速查和错题反馈路径

## 真正的 Word 公式

脚本会把常见 LaTeX 公式转换成 Word 的 Office Math / OMML 结构。

例如 JSON 中写：

```json
{
  "formula": "MSE = \\frac{1}{n}\\sum_{i=1}^{n}(y_i-\\hat{y}_i)^2"
}
```

Word 里会以公式对象写入，而不是把 `\\frac{...}` 当普通文本显示。当前支持常见结构：

- `\frac{a}{b}`
- 上标和下标，如 `x_i^2`
- `\sum_{i=1}^{n}`
- `\sqrt{x}`
- 常见希腊字母和数学符号
- `\hat{y}`、`\bar{x}`、`\tilde{x}`

复杂 LaTeX 后续还可以继续扩展到更完整的 OMML 转换。

## OCR / 截图公式识别接口

很多 PPT 的公式其实是图片，无法从 XML 中直接提取。这个项目提供 `--ocr-json` 接口，可以把外部 OCR 或数学公式识别结果合并进提取结果：

```bash
python3 scripts/ppt_notes_exporter.py "slides.pptx" \
  --ocr-json "ocr_results.json" \
  --output "PPT学习笔记.docx"
```

`ocr_results.json` 示例：

```json
{
  "slides": [
    {
      "number": 3,
      "text": ["OCR 识别出的文字"],
      "formula_candidates": ["H(X) = -\\sum_x p(x)\\log p(x)"],
      "visual_explanation": "图中展示了信源符号概率分布。"
    }
  ]
}
```

## 开放题库 / 自建题库接入

默认小题会根据 PPT 内容原创生成。如果你有开放教育题库或自建题库，可以用 `--practice-bank-json` 接入：

```bash
python3 scripts/ppt_notes_exporter.py "slides.pptx" \
  --notes-json "notes_filled.json" \
  --practice-bank-json "practice_bank.json" \
  --output "PPT学习笔记.docx"
```

题库格式：

```json
{
  "questions": [
    {
      "terms": ["entropy", "information theory"],
      "question": "已知两个符号概率分别为 1/2 和 1/2，求信源熵。",
      "answer": "1 bit",
      "solution": "代入 H(X)=-sum p(x)log2 p(x)，得到 1 bit。",
      "difficulty": "基础",
      "source": "开放教育题源改编",
      "source_url": "https://example.edu/open-resource"
    }
  ]
}
```

注意：不要直接复制商业题库。建议使用开放授权资料，保留来源链接，并改写成适配当前 PPT 的练习题。

## 错题反馈路径

做完题后，把错题填入 `可选_错题输入模板.json`，再运行：

```bash
python3 scripts/ppt_notes_exporter.py "slides.pptx" \
  --notes-json "notes_filled.json" \
  --wrong-answers-json "可选_错题输入模板.json" \
  --output "PPT学习笔记.docx"
```

新的 `09_错题反馈路径.md` 会按错题页和错因重新安排复习顺序，告诉学生：

- 哪几页必须重看
- 错因集中在哪里
- 明天应该怎么复习
- 哪些题要重新做

## 用作 Codex Skill

把仓库放到 Codex skills 目录：

```bash
mkdir -p ~/.codex/skills
cp -R summarize-ppt-notes ~/.codex/skills/
```

然后对 Codex 说：

```text
Use $summarize-ppt-notes to turn this PPT into final-exam review notes.
```

## 常用 CLI 参数

```bash
python3 scripts/ppt_notes_exporter.py "slides.pptx" \
  --slides 1,3-5 \
  --max-slides 10 \
  --content-filter study \
  --review-depth compressed \
  --max-deep-slides 24 \
  --dpi 180 \
  --output notes.docx \
  --workdir notes_work \
  --notes-json notes_filled.json \
  --ocr-json ocr_results.json \
  --practice-bank-json practice_bank.json \
  --wrong-answers-json wrong_answers.json \
  --fail-under 90
```

常用模式：

- `--template-only`：只生成提取文件和模板，不生成 Word
- `--no-render`：跳过截图渲染，适合快速 XML 提取
- `--slides 1,3-5`：只处理指定页
- `--max-slides 10`：只处理前 10 页
- `--content-filter study`：默认，压缩标题、目录、章节过渡和结束页，只展开正文学习页
- `--content-filter all`：全量逐页输出，适合审计或需要保留目录页讲解的场景
- `--notes-json`：传入填好的逐页讲解
- `--ocr-json`：合并外部 OCR / 公式识别结果
- `--practice-bank-json`：接入开放题库或自建题库
- `--wrong-answers-json`：根据错题生成二次学习路径
- `--study-mode final`：质量检查要求期末复习字段完整
- `--layout study`：生成学生友好的复习讲义
- `--layout audit`：保留更多原始提取内容，适合调试
- `--output-profile teacher`：默认，只输出干净交付目录
- `--output-profile complete`：同时打印所有中间文件路径
- `--output-profile debug`：只打印调试文件路径

## Benchmark

可以用 `benchmarks/run_benchmark.py` 检查更新是否破坏输出：

```bash
python3 benchmarks/run_benchmark.py examples/your_deck.pptx --max-slides 5
```

它会输出页数、质量分、公式页数量、图表页数量，以及学习路径、小题练习、HTML 页面、错题反馈是否生成。

## 依赖

必须：

- Python 3.9+

推荐：

- LibreOffice：用于 PPT/PPTX/PPT 转 PDF 和截图渲染
- PyMuPDF：用于 PDF 文本和截图处理
- Pillow：用于 Word 中图片尺寸计算

安装开发依赖：

```bash
python3 -m pip install --upgrade pip setuptools
python3 -m pip install -e ".[dev,render]"
```

## 项目结构

```text
summarize-ppt-notes/
├── SKILL.md
├── scripts/
│   ├── ppt_notes_exporter.py
│   └── doctor.py
├── references/
│   └── note-schema.md
├── docs/
├── examples/
├── benchmarks/
└── tests/
```

## 和普通转换工具的区别

本项目不追求成为所有格式的万能转换器。它更聚焦：

- PPT 到学习讲义
- 公式和图表解释
- 期末复习路径
- 主动回忆和小题练习
- 错题反馈
- 本地优先和可检查输出

MarkItDown、Unstructured、Docling 这类项目更适合作为上游解析或导入适配器；本项目的差异点是面向学生复习的教学闭环。

## 商业化方向

开源核心保持免费：

- 本地提取
- Word 生成
- 学习路径
- 小题练习
- 错题反馈
- 质量检查
- Codex Skill 工作流

可商业化层：

- 批量处理
- 更强 OCR / 数学公式识别
- 学校或机构模板
- 自动讲解模型
- 团队协作和审核
- LMS / Canvas / Moodle 导出
- 私有化部署

## 隐私

脚本默认在本地处理文件，不会主动上传课件。如果你使用 Codex 或其他模型填写 `notes_template.json`，请根据课件隐私级别自行判断是否可以发送给模型服务。

## License

MIT。见 [LICENSE](LICENSE)。
