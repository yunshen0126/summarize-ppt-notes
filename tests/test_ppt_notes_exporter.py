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
                        "purpose": "这一页用于说明模型预测误差的度量方式。",
                        "what_it_says": "页面给出 MSE 公式，并说明它用于衡量真实值和预测值之间的平均平方差。",
                        "detailed_explanation": "先计算每个样本的误差，再平方，最后求平均。平方可以避免正负抵消，并突出较大的预测偏差。",
                        "visual_explanation": "图片和图表表示误差趋势，用于帮助理解预测偏差如何随样本变化。",
                        "formula_explanations": [
                            {
                                "formula": "MSE = \\frac{1}{n}\\sum_i(y_i-\\hat y_i)^2",
                                "meaning": "n 是样本数，y_i 是真实值，\\hat y_i 是预测值。",
                                "conditions": "适用于回归任务。",
                                "example": "真实值 3、5，预测值 2、7，则 MSE=(1+4)/2=2.5。"
                            }
                        ],
                        "worked_examples": ["两条样本逐项计算误差平方，再除以样本数。"],
                        "exam_focus": "期末可能要求计算 MSE 或解释平方误差的含义。",
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
            report = json.loads((workdir / "quality_report.json").read_text(encoding="utf-8"))
            self.assertTrue(report["passed"])
            self.assertGreaterEqual(report["score"], 90)


if __name__ == "__main__":
    unittest.main()
