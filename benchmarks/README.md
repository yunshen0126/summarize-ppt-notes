# Benchmark

这个目录用于放公开、可分享的测试课件，验证每次更新是否真的提升学习笔记质量。

建议的 fixture 类型：

- 公式密集型 PPT
- 图表密集型 PPT
- 信息论 / 机器学习 / 数学课程 PPT
- 扫描图或截图公式较多的 PDF

运行：

```bash
python3 benchmarks/run_benchmark.py examples/your_deck.pptx --max-slides 5
```

输出会统计：

- 页数
- 是否生成 Word
- 是否生成学习路径、小题练习、HTML 学习页、错题反馈路径
- 质量分
- 公式页和图表页数量

