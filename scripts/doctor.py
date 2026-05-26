#!/usr/bin/env python3
from __future__ import annotations

import argparse
import importlib.util
import shutil
import sys
from pathlib import Path
from typing import List, Optional


def has_module(name: str) -> bool:
    return importlib.util.find_spec(name) is not None


def has_soffice() -> bool:
    candidates = [
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    return any(candidate and Path(candidate).exists() for candidate in candidates)


def status_line(ok: bool, label: str, detail: str) -> str:
    return f"[{'OK' if ok else 'WARN'}] {label}: {detail}"


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Check summarize-ppt-notes runtime dependencies.")
    parser.add_argument("--strict", action="store_true", help="Return a non-zero exit code when optional render dependencies are missing")
    args = parser.parse_args(argv)

    checks = [
        (sys.version_info >= (3, 9), "Python", sys.version.split()[0] + " (requires >= 3.9)"),
        (has_module("fitz"), "PyMuPDF", "needed for PDF text extraction and slide screenshots"),
        (has_module("PIL"), "Pillow", "needed for accurate image sizing in DOCX output"),
        (has_soffice(), "LibreOffice", "needed to render PPT/PPTX/PPT to full-slide screenshots"),
    ]
    for ok, label, detail in checks:
        print(status_line(ok, label, detail))

    hard_fail = not checks[0][0]
    optional_fail = any(not ok for ok, label, _ in checks if label != "Python")
    if hard_fail or (args.strict and optional_fail):
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
