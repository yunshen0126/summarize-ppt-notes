#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import subprocess
import sys
import tempfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
EXPORTER = ROOT / "scripts" / "ppt_notes_exporter.py"


def main() -> int:
    parser = argparse.ArgumentParser(description="Run a small summarize-ppt-notes benchmark fixture.")
    parser.add_argument("source", help="PPT/PPTX/PDF fixture path")
    parser.add_argument("--max-slides", type=int, default=5)
    parser.add_argument("--notes-json")
    args = parser.parse_args()

    source = Path(args.source).expanduser().resolve()
    if not source.exists():
        print(f"fixture not found: {source}", file=sys.stderr)
        return 2

    with tempfile.TemporaryDirectory(prefix="summarize-ppt-notes-bench-") as tmp:
        tmp_path = Path(tmp)
        output = tmp_path / "notes.docx"
        workdir = tmp_path / "work"
        cmd = [
            sys.executable,
            str(EXPORTER),
            str(source),
            "--no-render",
            "--max-slides",
            str(args.max_slides),
            "--output",
            str(output),
            "--workdir",
            str(workdir),
            "--study-mode",
            "final",
        ]
        if args.notes_json:
            cmd += ["--notes-json", str(Path(args.notes_json).expanduser().resolve())]
        result = subprocess.run(cmd, text=True, capture_output=True, check=False)
        if result.returncode not in (0, 1):
            print(result.stdout)
            print(result.stderr, file=sys.stderr)
            return result.returncode
        extraction = json.loads((workdir / "extraction.json").read_text(encoding="utf-8"))
        quality = json.loads((workdir / "quality_report.json").read_text(encoding="utf-8"))
        study_pack = workdir / "study_pack"
        metrics = {
            "source": str(source),
            "slides": extraction.get("slide_count", 0),
            "quality_score": quality.get("score", 0),
            "docx": output.exists(),
            "formula_slides": sum(1 for slide in extraction.get("slides", []) if slide.get("formulas") or slide.get("formula_candidates")),
            "visual_slides": sum(1 for slide in extraction.get("slides", []) if slide.get("images") or slide.get("tables") or slide.get("related_objects")),
            "learning_path": (study_pack / "learning_path.md").exists(),
            "practice_questions": (study_pack / "practice_questions.md").exists(),
            "study_html": (study_pack / "study_index.html").exists(),
            "adaptive_review": (study_pack / "adaptive_review.md").exists(),
        }
        print(json.dumps(metrics, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
