from __future__ import annotations

import base64
import json
import subprocess
import sys
import unittest
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "ppt_notes_exporter.py"

sys.path.insert(0, str(ROOT / "scripts"))
import ppt_notes_exporter as exporter  # noqa: E402


PNG_1X1 = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII="
)


def write_minimal_pptx(path: Path) -> None:
    slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:txBody>
          <a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:t>MSE = (1/n) sum((y_i - yhat_i)^2)</a:t></a:r></a:p>
          <a:p><a:r><a:t>This slide explains prediction error.</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:pic>
        <p:nvPicPr><p:cNvPr id="3" name="Picture 1" descr="error plot"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
        <p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
        <p:spPr/>
      </p:pic>
    </p:spTree>
  </p:cSld>
</p:sld>
"""
    presentation_xml = """<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
</p:presentation>
"""
    presentation_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>
"""
    slide_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>
"""
    chart_xml = """<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart><c:title><c:tx><c:rich><c:p><c:r><c:t>Error trend</c:t></c:r></c:p></c:rich></c:tx></c:title></c:chart>
  <c:v>Series A</c:v><c:v>0.1</c:v><c:v>0.4</c:v>
</c:chartSpace>
"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as pptx:
        pptx.writestr("ppt/presentation.xml", presentation_xml)
        pptx.writestr("ppt/_rels/presentation.xml.rels", presentation_rels)
        pptx.writestr("ppt/slides/slide1.xml", slide_xml)
        pptx.writestr("ppt/slides/_rels/slide1.xml.rels", slide_rels)
        pptx.writestr("ppt/charts/chart1.xml", chart_xml)
        pptx.writestr("ppt/media/image1.png", PNG_1X1)


class PptNotesExporterTest(unittest.TestCase):
    def test_extract_pptx_reads_text_images_and_related_objects(self) -> None:
        with self.subTest("minimal pptx extraction"):
            import tempfile

            with tempfile.TemporaryDirectory() as tmp:
                tmp_path = Path(tmp)
                source = tmp_path / "sample.pptx"
                write_minimal_pptx(source)

                extraction = exporter.extract_source(source, tmp_path / "work", dpi=90, no_render=True)
                slide = extraction["slides"][0]

                self.assertEqual(extraction["slide_count"], 1)
                self.assertIn("MSE", " ".join(slide["text"]))
                self.assertTrue(slide["formula_candidates"])
                self.assertTrue(slide["images"][0]["filename"].endswith(".png"))
                self.assertEqual(slide["related_objects"][0]["source"], "ppt/charts/chart1.xml")

    def test_cli_creates_docx_and_quality_report(self) -> None:
        import tempfile

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "sample.pptx"
            output = tmp_path / "notes.docx"
            workdir = tmp_path / "work"
            write_minimal_pptx(source)

            notes = {
                "slides": [
                    {
                        "number": 1,
                        "title": "MSE",
                        "purpose": "这一页用于说明回归模型的预测误差如何被量化，核心对象是 MSE 平均平方误差。",
                        "what_it_says": "页面给出 MSE 公式 MSE=(1/n)sum((y_i-yhat_i)^2)，说明它把每个样本的真实值 y_i 与预测值 yhat_i 做差、平方、求平均，用一个非负数衡量预测偏差。",
                        "detailed_explanation": "MSE 的定义对象是一组回归预测结果。条件是每个样本都有真实值 y_i 和预测值 yhat_i，并且样本数量为 n。计算时先逐样本求误差 y_i-yhat_i，再平方以避免正负误差相互抵消，同时让较大的偏差受到更重惩罚，最后除以 n 得到平均损失。它的数值越小，表示预测整体越接近真实值；但由于平方会放大异常误差，所以对离群点较敏感。期末题通常要求代入两到三个样本手算，或解释为什么平方误差比绝对误差更强调大偏差。",
                        "visual_explanation": "图表区域表示误差趋势或预测偏差的可视化：横向可理解为样本或时间，纵向可理解为误差大小。复习时需要把图中的偏差距离与 MSE 公式里的平方项对应起来。",
                        "formula_explanations": [
                            {
                                "formula": "MSE = \\frac{1}{n}\\sum_i(y_i-\\hat y_i)^2",
                                "meaning": "n 是样本数，y_i 是真实值，\\hat y_i 是预测值。",
                                "conditions": "适用于回归任务。",
                                "example": "真实值 3、5，预测值 2、7，则 MSE=(1+4)/2=2.5。"
                            }
                        ],
                        "worked_examples": ["两条样本逐项计算误差平方，再除以样本数。"],
                        "exam_focus": "期末可能给出一组真实值和预测值，要求写出 MSE 公式、逐项计算平方误差、求平均，并解释平方项为什么会放大大误差。",
                        "key_takeaways": ["MSE 是平均平方误差。", "平方会放大大误差。"],
                        "memory_hooks": ["差值 -> 平方 -> 平均。"],
                        "likely_questions": [
                            {
                                "question": "给定两条样本，如何计算 MSE？",
                                "answer": "逐项计算误差平方，再除以样本数。"
                            }
                        ],
                        "common_mistakes": ["忘记除以样本数。"],
                        "prerequisites": ["平均值", "平方"],
                        "difficulty": "基础",
                        "estimated_review_minutes": "6",
                        "tags": ["mse", "loss"],
                        "uncertainties": []
                    }
                ]
            }
            notes_path = tmp_path / "notes.json"
            notes_path.write_text(json.dumps(notes, ensure_ascii=False), encoding="utf-8")

            result = subprocess.run(
                [
                    sys.executable,
                    str(SCRIPT),
                    str(source),
                    "--no-render",
                    "--notes-json",
                    str(notes_path),
                    "--output",
                    str(output),
                    "--workdir",
                    str(workdir),
                    "--fail-under",
                    "90",
                    "--study-mode",
                    "final",
                ],
                text=True,
                capture_output=True,
                check=False,
            )

            self.assertEqual(result.returncode, 0, result.stderr + result.stdout)
            self.assertTrue(output.exists())
            self.assertTrue(zipfile.is_zipfile(output))
            self.assertTrue((workdir / "final_notes.md").exists())
            self.assertTrue((workdir / "prompt_pack.md").exists())
            self.assertTrue((workdir / "study_pack" / "flashcards_anki.csv").exists())
            self.assertTrue((workdir / "study_pack" / "active_recall_questions.md").exists())
            self.assertTrue((workdir / "study_pack" / "formula_sheet.md").exists())
            self.assertTrue((workdir / "study_pack" / "mistake_log_template.md").exists())
            deliverables = output.with_suffix("").with_name(output.stem + "_deliverables")
            self.assertTrue((deliverables / "START_HERE.md").exists())
            self.assertTrue((deliverables / "01_复习讲义.docx").exists())
            self.assertTrue((deliverables / "04_主动回忆题.md").exists())
            report = json.loads((workdir / "quality_report.json").read_text(encoding="utf-8"))
            self.assertTrue(report["passed"])
            self.assertGreaterEqual(report["score"], 90)


if __name__ == "__main__":
    unittest.main()
