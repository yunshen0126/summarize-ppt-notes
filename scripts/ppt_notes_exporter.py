#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import hashlib
from html import escape as html_escape
import io
import json
import mimetypes
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from datetime import date as Date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape


NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

WORD_NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "m": NS_M,
}

VERSION = "0.11.0"

BANNED_FLUFF_PATTERNS = [
    r"放回.*主线.*理解",
    r"先明确解决的问题",
    r"建议.*仔细.*理解",
    r"复习时按三步走",
    r"本页需要.*理解",
    r"请.*回看.*截图",
    r"待补写",
    r"无或待补写",
    r"本页属于.*模块",
]

STYLE_COLORS = {
    "ink": "1F2937",
    "muted": "6B7280",
    "blue": "1D4ED8",
    "blue_bg": "EFF6FF",
    "green": "047857",
    "green_bg": "ECFDF5",
    "amber": "B45309",
    "amber_bg": "FFFBEB",
    "red": "B91C1C",
    "red_bg": "FEF2F2",
    "slate_bg": "F8FAFC",
}

TOPIC_STOPWORDS = {
    "slide",
    "slides",
    "page",
    "pages",
    "lecture",
    "chapter",
    "part",
    "section",
    "example",
    "examples",
    "exercise",
    "exercises",
    "homework",
    "summary",
    "note",
    "notes",
    "the",
    "and",
    "for",
    "with",
    "from",
    "into",
    "onto",
    "about",
    "between",
    "of",
    "to",
    "in",
    "on",
    "by",
    "vs",
    "problem",
    "proof",
    "definition",
    "theorem",
    "题目",
    "练习",
    "习题",
    "总结",
    "定义",
    "定理",
    "证明",
    "概念",
    "方法",
    "问题",
    "引言",
    "复习",
    "页面",
    "章节",
    "课程",
    "讲义",
    "材料",
    "内容",
    "学习",
}

NAVIGATION_CONTENT_KINDS = {"title", "agenda", "section", "closing"}

REVIEW_TIERS = {"deep", "quick", "reference"}

REVIEW_DEPTH_LABELS = {
    "compressed": "压缩复习",
    "balanced": "均衡复习",
    "complete": "完整讲义",
}

REVIEW_TIER_LABELS = {
    "deep": "必读深讲",
    "quick": "快速扫读",
    "reference": "参考/重复",
}

AGENDA_MARKERS = (
    "agenda",
    "outline",
    "contents",
    "roadmap",
    "overview",
    "目录",
    "大纲",
    "提纲",
    "本章内容",
    "本讲内容",
    "本节内容",
    "主要内容",
    "课程安排",
    "学习目标",
)

TITLE_MARKERS = (
    "lecture",
    "chapter",
    "chap.",
    "course",
    "presentation",
    "slides",
    "final review",
    "期末复习",
    "课程",
    "讲义",
    "课件",
)

CLOSING_MARKERS = (
    "thank you",
    "thanks",
    "q&a",
    "questions?",
    "questions",
    "谢谢",
    "答疑",
    "提问",
)

MATH_GLYPH_MARKERS = (
    "\uf03d",  # Symbol-font equals in many converted PDFs.
    "\uf0e5",  # Symbol-font summation.
    "\uf0a3",
    "\uf0b3",
    "\uf02d",
    "\uf02b",
    "\uf0d7",
    "\uf0f9",
    "\uf0fa",
    "\uf0fb",
)

CONNECTOR_HINTS = {
    "example",
    "examples",
    "exercise",
    "exercises",
    "homework",
    "练习",
    "习题",
    "例题",
    "作业",
    "总结",
    "回顾",
    "review",
    "summary",
}

ROLE_LABELS = {
    "concept": "概念",
    "formula": "公式/推导",
    "algorithm": "方法/算法",
    "example": "例题/案例",
    "practice": "练习",
    "summary": "总结",
    "visual": "图表理解",
}

SUPERSCRIPT_CHARS = str.maketrans(
    {
        "0": "⁰",
        "1": "¹",
        "2": "²",
        "3": "³",
        "4": "⁴",
        "5": "⁵",
        "6": "⁶",
        "7": "⁷",
        "8": "⁸",
        "9": "⁹",
        "+": "⁺",
        "-": "⁻",
        "=": "⁼",
        "(": "⁽",
        ")": "⁾",
        "n": "ⁿ",
        "i": "ⁱ",
    }
)

SUBSCRIPT_CHARS = str.maketrans(
    {
        "0": "₀",
        "1": "₁",
        "2": "₂",
        "3": "₃",
        "4": "₄",
        "5": "₅",
        "6": "₆",
        "7": "₇",
        "8": "₈",
        "9": "₉",
        "+": "₊",
        "-": "₋",
        "=": "₌",
        "(": "₍",
        ")": "₎",
        "a": "ₐ",
        "e": "ₑ",
        "h": "ₕ",
        "i": "ᵢ",
        "j": "ⱼ",
        "k": "ₖ",
        "l": "ₗ",
        "m": "ₘ",
        "n": "ₙ",
        "o": "ₒ",
        "p": "ₚ",
        "r": "ᵣ",
        "s": "ₛ",
        "t": "ₜ",
        "u": "ᵤ",
        "v": "ᵥ",
        "x": "ₓ",
    }
)

LATEX_SYMBOLS = {
    r"\alpha": "α",
    r"\beta": "β",
    r"\gamma": "γ",
    r"\delta": "δ",
    r"\epsilon": "ε",
    r"\varepsilon": "ε",
    r"\theta": "θ",
    r"\lambda": "λ",
    r"\mu": "μ",
    r"\pi": "π",
    r"\sigma": "σ",
    r"\tau": "τ",
    r"\phi": "φ",
    r"\omega": "ω",
    r"\Delta": "Δ",
    r"\Sigma": "Σ",
    r"\sum": "∑",
    r"\prod": "∏",
    r"\int": "∫",
    r"\infty": "∞",
    r"\leq": "≤",
    r"\le": "≤",
    r"\geq": "≥",
    r"\ge": "≥",
    r"\neq": "≠",
    r"\approx": "≈",
    r"\cdot": "·",
    r"\times": "×",
    r"\div": "÷",
    r"\pm": "±",
    r"\rightarrow": "→",
    r"\to": "→",
    r"\leftarrow": "←",
    r"\Rightarrow": "⇒",
    r"\forall": "∀",
    r"\exists": "∃",
    r"\in": "∈",
    r"\notin": "∉",
    r"\cup": "∪",
    r"\cap": "∩",
    r"\log": "log",
    r"\ln": "ln",
}


def local_name(tag: str) -> str:
    if "}" in tag:
        return tag.rsplit("}", 1)[1]
    return tag


def namespace(tag: str) -> str:
    if tag.startswith("{") and "}" in tag:
        return tag[1:].split("}", 1)[0]
    return ""


def parse_xml(data: bytes) -> Optional[ET.Element]:
    try:
        return ET.fromstring(data)
    except ET.ParseError:
        return None


def unique_preserve(items: Iterable[str]) -> List[str]:
    seen = set()
    result = []
    for item in items:
        normalized = re.sub(r"\s+", " ", item or "").strip()
        if not normalized or normalized in seen:
            continue
        seen.add(normalized)
        result.append(normalized)
    return result


def normalize_zip_path(base_dir: str, target: str) -> str:
    if target.startswith("/"):
        target = target[1:]
        base_dir = ""
    combined = posixpath.join(base_dir, target)
    parts: List[str] = []
    for part in combined.split("/"):
        if part in ("", "."):
            continue
        if part == "..":
            if parts:
                parts.pop()
        else:
            parts.append(part)
    return "/".join(parts)


def paragraph_texts(root: ET.Element) -> List[str]:
    texts = []
    for elem in root.iter():
        if local_name(elem.tag) != "p" or namespace(elem.tag) != NS_A:
            continue
        parts = []
        for child in elem.iter():
            if local_name(child.tag) == "t" and child.text:
                parts.append(child.text)
        text = "".join(parts).strip()
        if text:
            texts.append(text)
    return unique_preserve(texts)


def element_text(elem: ET.Element) -> str:
    parts = []
    for child in elem.iter():
        if local_name(child.tag) == "t" and child.text:
            parts.append(child.text)
    return "".join(parts).strip()


def extract_tables(root: ET.Element) -> List[List[List[str]]]:
    tables: List[List[List[str]]] = []
    for tbl in root.iter():
        if local_name(tbl.tag) != "tbl" or namespace(tbl.tag) != NS_A:
            continue
        rows: List[List[str]] = []
        for tr in tbl.iter():
            if local_name(tr.tag) != "tr" or namespace(tr.tag) != NS_A:
                continue
            row = []
            for tc in tr.iter():
                if local_name(tc.tag) == "tc" and namespace(tc.tag) == NS_A:
                    row.append(element_text(tc))
            if row:
                rows.append(row)
        if rows:
            tables.append(rows)
    return tables


def extract_alt_texts(root: ET.Element) -> List[str]:
    values = []
    for elem in root.iter():
        if local_name(elem.tag) != "cNvPr":
            continue
        name = (elem.attrib.get("name") or "").strip()
        descr = (elem.attrib.get("descr") or "").strip()
        combined = " - ".join(part for part in [name, descr] if part)
        if combined and not re.fullmatch(r"(Rectangle|TextBox|Picture|Image|Oval|Line)\s*\d*", combined, re.I):
            values.append(combined)
    return unique_preserve(values)


def extract_formula_elements(root: ET.Element) -> List[Dict[str, str]]:
    formulas: List[Dict[str, str]] = []
    seen = set()
    for elem in root.iter():
        if namespace(elem.tag) != NS_M or local_name(elem.tag) not in ("oMath", "oMathPara"):
            continue
        text = element_text(elem)
        xml = ET.tostring(elem, encoding="unicode")
        key = text or xml
        if key in seen:
            continue
        seen.add(key)
        formulas.append({"text": text, "omml": xml})
    return formulas


FORMULA_RE = re.compile(
    r"(\\frac|\\sum|\\int|\\sqrt|[=≈≠≤≥∑∫√∞±×÷∂∇]|[A-Za-z0-9]\s*[\^_]\s*[A-Za-z0-9({]|"
    r"\b(sin|cos|tan|log|ln|argmax|argmin|max|min|softmax|sigmoid)\s*[\(\[]|"
    r"\b(P|Pr|E|Var|Cov)\s*[\(\[]|[α-ωΑ-Ω])"
)


def formula_candidates_from_text(texts: Iterable[str]) -> List[str]:
    candidates = []
    for text in texts:
        cleaned = re.sub(r"\s+", " ", text).strip()
        if len(cleaned) < 2:
            continue
        if FORMULA_RE.search(cleaned):
            candidates.append(cleaned)
    return unique_preserve(candidates)


def all_xml_text_values(root: ET.Element) -> List[str]:
    values = []
    for elem in root.iter():
        if local_name(elem.tag) in {"t", "v"} and elem.text:
            values.append(elem.text)
    return unique_preserve(values)


def parse_relationships(zf: zipfile.ZipFile, rels_path: str) -> Dict[str, Dict[str, str]]:
    if rels_path not in zf.namelist():
        return {}
    root = parse_xml(zf.read(rels_path))
    if root is None:
        return {}
    rels: Dict[str, Dict[str, str]] = {}
    for rel in root:
        if local_name(rel.tag) != "Relationship":
            continue
        rel_id = rel.attrib.get("Id")
        if not rel_id:
            continue
        rels[rel_id] = {
            "target": rel.attrib.get("Target", ""),
            "type": rel.attrib.get("Type", ""),
            "mode": rel.attrib.get("TargetMode", ""),
        }
    return rels


def slide_paths_in_order(zf: zipfile.ZipFile) -> List[str]:
    names = set(zf.namelist())
    ordered: List[str] = []
    if "ppt/presentation.xml" in names and "ppt/_rels/presentation.xml.rels" in names:
        pres = parse_xml(zf.read("ppt/presentation.xml"))
        rels = parse_relationships(zf, "ppt/_rels/presentation.xml.rels")
        if pres is not None:
            for elem in pres.iter():
                if local_name(elem.tag) != "sldId":
                    continue
                rid = elem.attrib.get(f"{{{NS_R}}}id")
                target = rels.get(rid or "", {}).get("target", "")
                if not target:
                    continue
                path = normalize_zip_path("ppt", target)
                if path in names and path not in ordered:
                    ordered.append(path)
    if ordered:
        return ordered

    slide_re = re.compile(r"ppt/slides/slide(\d+)\.xml$")
    slides = []
    for name in names:
        match = slide_re.match(name)
        if match:
            slides.append((int(match.group(1)), name))
    return [name for _, name in sorted(slides)]


def extract_slide_images(
    zf: zipfile.ZipFile,
    slide_path: str,
    root: ET.Element,
    slide_number: int,
    assets_dir: Path,
) -> List[Dict[str, str]]:
    rels_path = posixpath.join(
        posixpath.dirname(slide_path),
        "_rels",
        posixpath.basename(slide_path) + ".rels",
    )
    rels = parse_relationships(zf, rels_path)
    found_rids = []
    for elem in root.iter():
        if local_name(elem.tag) != "blip":
            continue
        rid = elem.attrib.get(f"{{{NS_R}}}embed") or elem.attrib.get(f"{{{NS_R}}}link")
        if rid:
            found_rids.append(rid)

    images = []
    used = set()
    for image_index, rid in enumerate(found_rids, start=1):
        if rid in used:
            continue
        used.add(rid)
        rel = rels.get(rid)
        if not rel or rel.get("mode") == "External":
            images.append({"relationship_id": rid, "target": rel.get("target", "") if rel else "", "external": "true"})
            continue
        target = rel.get("target", "")
        media_path = normalize_zip_path(posixpath.dirname(slide_path), target)
        if media_path not in zf.namelist():
            images.append({"relationship_id": rid, "target": media_path, "missing": "true"})
            continue
        ext = Path(media_path).suffix or ".bin"
        out_path = assets_dir / f"slide_{slide_number:03d}_image_{image_index:02d}{ext.lower()}"
        out_path.write_bytes(zf.read(media_path))
        images.append(
            {
                "relationship_id": rid,
                "source": media_path,
                "path": str(out_path.resolve()),
                "filename": out_path.name,
            }
        )
    return images


def extract_slide_related_objects(zf: zipfile.ZipFile, slide_path: str) -> List[Dict[str, Any]]:
    rels_path = posixpath.join(
        posixpath.dirname(slide_path),
        "_rels",
        posixpath.basename(slide_path) + ".rels",
    )
    rels = parse_relationships(zf, rels_path)
    objects: List[Dict[str, Any]] = []
    for rid, rel in sorted(rels.items()):
        rel_type = rel.get("type", "")
        target = rel.get("target", "")
        marker = f"{rel_type} {target}".lower()
        if not any(token in marker for token in ("chart", "diagram", "oleobject", "package", "embedded")):
            continue
        item: Dict[str, Any] = {
            "relationship_id": rid,
            "type": rel_type,
            "target": target,
        }
        if rel.get("mode") == "External":
            item["external"] = True
            objects.append(item)
            continue
        part_path = normalize_zip_path(posixpath.dirname(slide_path), target)
        item["source"] = part_path
        if part_path in zf.namelist():
            if part_path.lower().endswith(".xml"):
                root = parse_xml(zf.read(part_path))
                if root is not None:
                    item["text_values"] = all_xml_text_values(root)[:120]
                    item["formula_candidates"] = formula_candidates_from_text(item["text_values"])
            else:
                item["size_bytes"] = len(zf.read(part_path))
        else:
            item["missing"] = True
        objects.append(item)
    return objects


def notes_path_for_slide(slide_path: str) -> Optional[str]:
    match = re.search(r"slide(\d+)\.xml$", slide_path)
    if not match:
        return None
    return f"ppt/notesSlides/notesSlide{match.group(1)}.xml"


def extract_pptx(source: Path, assets_dir: Path, warnings: List[str]) -> List[Dict[str, Any]]:
    slides: List[Dict[str, Any]] = []
    try:
        zf = zipfile.ZipFile(source)
    except zipfile.BadZipFile:
        warnings.append(f"Cannot parse {source.name} as pptx zip.")
        return slides

    with zf:
        paths = slide_paths_in_order(zf)
        for index, slide_path in enumerate(paths, start=1):
            root = parse_xml(zf.read(slide_path))
            if root is None:
                warnings.append(f"Cannot parse XML for slide {index}.")
                continue

            text = paragraph_texts(root)
            notes_text: List[str] = []
            notes_path = notes_path_for_slide(slide_path)
            if notes_path and notes_path in zf.namelist():
                notes_root = parse_xml(zf.read(notes_path))
                if notes_root is not None:
                    notes_text = paragraph_texts(notes_root)

            formulas = extract_formula_elements(root)
            formula_candidates = formula_candidates_from_text(text)
            table_data = extract_tables(root)
            alt_texts = extract_alt_texts(root)
            images = extract_slide_images(zf, slide_path, root, index, assets_dir)
            related_objects = extract_slide_related_objects(zf, slide_path)

            title = text[0] if text else f"Slide {index}"
            slides.append(
                {
                    "number": index,
                    "title": title,
                    "source_slide_xml": slide_path,
                    "text": text,
                    "notes": notes_text,
                    "tables": table_data,
                    "formulas": formulas,
                    "formula_candidates": formula_candidates,
                    "alt_texts": alt_texts,
                    "images": images,
                    "related_objects": related_objects,
                    "screenshot": "",
                }
            )
    return slides


def find_soffice() -> Optional[str]:
    for candidate in [
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]:
        if candidate and Path(candidate).exists():
            return str(candidate)
    return None


def convert_with_libreoffice(source: Path, out_dir: Path, target_ext: str, warnings: List[str]) -> Optional[Path]:
    soffice = find_soffice()
    if not soffice:
        warnings.append("LibreOffice/soffice was not found, so Office-to-PDF/PPTX conversion was skipped.")
        return None
    out_dir.mkdir(parents=True, exist_ok=True)
    profile = tempfile.mkdtemp(prefix="lo-profile-")
    before = set(out_dir.glob(f"*.{target_ext}"))
    cmd = [
        soffice,
        "--headless",
        f"-env:UserInstallation=file://{profile}",
        "--convert-to",
        target_ext,
        "--outdir",
        str(out_dir),
        str(source),
    ]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    except Exception as exc:
        warnings.append(f"LibreOffice conversion to {target_ext} failed: {exc}")
        shutil.rmtree(profile, ignore_errors=True)
        return None
    finally:
        shutil.rmtree(profile, ignore_errors=True)

    if result.returncode != 0:
        detail = (result.stderr or result.stdout or "").strip()
        warnings.append(f"LibreOffice conversion to {target_ext} failed: {detail}")
        return None

    expected = out_dir / f"{source.stem}.{target_ext}"
    if expected.exists():
        return expected
    after = set(out_dir.glob(f"*.{target_ext}")) - before
    if after:
        return sorted(after, key=lambda p: p.stat().st_mtime)[-1]
    warnings.append(f"LibreOffice reported success, but no .{target_ext} file was found.")
    return None


def render_pdf_with_fitz(pdf_path: Path, screenshots_dir: Path, dpi: int, warnings: List[str]) -> Tuple[List[str], List[str]]:
    screenshots_dir.mkdir(parents=True, exist_ok=True)
    try:
        import fitz  # type: ignore
    except Exception:
        warnings.append("PyMuPDF/fitz is unavailable, so PDF screenshot rendering was skipped.")
        return [], []

    screenshots: List[str] = []
    page_texts: List[str] = []
    try:
        doc = fitz.open(str(pdf_path))
        zoom = dpi / 72.0
        matrix = fitz.Matrix(zoom, zoom)
        for index, page in enumerate(doc, start=1):
            page_texts.append((page.get_text("text") or "").strip())
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            out_path = screenshots_dir / f"slide_{index:03d}.png"
            pix.save(str(out_path))
            screenshots.append(str(out_path.resolve()))
        doc.close()
    except Exception as exc:
        warnings.append(f"PDF rendering failed: {exc}")
    return screenshots, page_texts


def ensure_slides_for_screenshots(slides: List[Dict[str, Any]], screenshots: List[str], page_texts: List[str]) -> List[Dict[str, Any]]:
    max_count = max(len(slides), len(screenshots), len(page_texts))
    while len(slides) < max_count:
        n = len(slides) + 1
        raw_text = page_texts[n - 1] if n - 1 < len(page_texts) else ""
        text_items = unique_preserve(raw_text.splitlines())
        slides.append(
            {
                "number": n,
                "title": text_items[0] if text_items else f"Slide {n}",
                "source_slide_xml": "",
                "text": text_items,
                "notes": [],
                "tables": [],
                "formulas": [],
                "formula_candidates": formula_candidates_from_text(text_items),
                "alt_texts": [],
                "images": [],
                "related_objects": [],
                "screenshot": "",
            }
        )
    for index, screenshot in enumerate(screenshots, start=1):
        slides[index - 1]["screenshot"] = screenshot
    for index, raw_text in enumerate(page_texts, start=1):
        if index - 1 >= len(slides) or not raw_text:
            continue
        text_items = unique_preserve(raw_text.splitlines())
        existing = slides[index - 1].get("text") or []
        slides[index - 1]["text"] = unique_preserve(list(existing) + text_items)
        slides[index - 1]["formula_candidates"] = unique_preserve(
            list(slides[index - 1].get("formula_candidates") or []) + formula_candidates_from_text(text_items)
        )
    return slides


def extract_source(source: Path, workdir: Path, dpi: int, no_render: bool) -> Dict[str, Any]:
    warnings: List[str] = []
    assets_dir = workdir / "assets"
    screenshots_dir = workdir / "screenshots"
    converted_dir = workdir / "converted"
    assets_dir.mkdir(parents=True, exist_ok=True)
    screenshots_dir.mkdir(parents=True, exist_ok=True)
    converted_dir.mkdir(parents=True, exist_ok=True)

    source = source.resolve()
    ext = source.suffix.lower()
    parse_source = source
    slides: List[Dict[str, Any]] = []

    if ext == ".ppt":
        converted = convert_with_libreoffice(source, converted_dir, "pptx", warnings)
        if converted:
            parse_source = converted
            ext = ".pptx"

    if ext == ".pptx":
        slides = extract_pptx(parse_source, assets_dir, warnings)
    elif ext == ".pdf":
        slides = []
    else:
        warnings.append(f"Unsupported source extension '{source.suffix}'. Supported: .pptx, .ppt, .pdf.")

    screenshots: List[str] = []
    page_texts: List[str] = []
    if not no_render:
        pdf_path: Optional[Path]
        if source.suffix.lower() == ".pdf":
            pdf_path = source
        else:
            pdf_path = convert_with_libreoffice(source, converted_dir, "pdf", warnings)
        if pdf_path:
            screenshots, page_texts = render_pdf_with_fitz(pdf_path, screenshots_dir, dpi, warnings)

    slides = ensure_slides_for_screenshots(slides, screenshots, page_texts)
    if not slides:
        warnings.append("No slides/pages were extracted.")

    return {
        "tool": "summarize-ppt-notes",
        "version": VERSION,
        "source": str(source),
        "source_sha256": sha256_file(source),
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "workdir": str(workdir.resolve()),
        "slide_count": len(slides),
        "slides": slides,
        "warnings": warnings,
    }


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def parse_slide_selector(selector: Optional[str]) -> Optional[List[int]]:
    if not selector:
        return None
    selected: List[int] = []
    for part in selector.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            start_text, end_text = part.split("-", 1)
            start = int(start_text)
            end = int(end_text)
            if start <= 0 or end <= 0 or end < start:
                raise ValueError(f"Invalid slide range: {part}")
            selected.extend(range(start, end + 1))
        else:
            value = int(part)
            if value <= 0:
                raise ValueError(f"Invalid slide number: {part}")
            selected.append(value)
    return list(dict.fromkeys(selected))


def filter_extraction_slides(extraction: Dict[str, Any], selector: Optional[str], max_slides: Optional[int]) -> None:
    slides = extraction.get("slides", [])
    selected_numbers = parse_slide_selector(selector)
    if selected_numbers is not None:
        allowed = set(selected_numbers)
        slides = [slide for slide in slides if int(slide.get("number", 0) or 0) in allowed]
        extraction["selected_slides"] = selected_numbers
    if max_slides is not None:
        if max_slides <= 0:
            raise ValueError("--max-slides must be greater than 0")
        slides = slides[:max_slides]
        extraction["max_slides"] = max_slides
    extraction["slides"] = slides
    extraction["slide_count"] = len(slides)


def slide_visible_text_items(slide: Dict[str, Any]) -> List[str]:
    items: List[str] = []
    for value in as_list(slide.get("title")) + as_list(slide.get("text")):
        text = plain(value)
        if text:
            items.append(text)
    return unique_preserve(items)


def slide_has_formula_signal(slide: Dict[str, Any]) -> bool:
    if slide.get("formulas"):
        return True
    visible_text = "\n".join(slide_visible_text_items(slide))
    if any(marker in visible_text for marker in MATH_GLYPH_MARKERS):
        return True
    if any(token in visible_text for token in ("=", "≈", "≠", "≤", "≥", "\\frac", "\\sum", "\\int", "\\sqrt", "∑", "∫", "√")):
        return True
    if re.search(r"[A-Za-z0-9]\s*[\^_]\s*[A-Za-z0-9({]", visible_text):
        return True
    for candidate in as_list(slide.get("formula_candidates")):
        text = plain(candidate)
        if not text:
            continue
        if any(token in text for token in ("=", "≈", "≠", "≤", "≥", "\\frac", "\\sum", "\\int", "\\sqrt", "∑", "∫", "√")):
            return True
        if re.search(r"[A-Za-z0-9]\s*[\^_]\s*[A-Za-z0-9({]", text):
            return True
    return False


def slide_has_structured_learning_signal(slide: Dict[str, Any]) -> bool:
    return bool(
        slide_has_formula_signal(slide)
        or slide.get("tables")
        or slide.get("related_objects")
        or slide.get("ocr_visual_explanations")
    )


def text_has_any_marker(text: str, markers: Tuple[str, ...]) -> bool:
    lower = text.lower()
    return any(marker in lower for marker in markers)


def is_agenda_like_slide(title: str, text: str, line_count: int, text_units: int) -> bool:
    title_lower = title.lower()
    text_lower = text.lower()
    strong_title_markers = ("agenda", "outline", "contents", "roadmap", "目录", "大纲", "提纲", "本章内容", "本讲内容", "本节内容", "课程安排")
    weak_title_markers = ("overview", "context", "主要内容", "学习目标")
    if re.search(r"\b(lecture|chapter|chap\.?)\b", title_lower) and line_count >= 5 and text_units <= 150:
        return True
    if any(marker in title_lower for marker in strong_title_markers):
        return text_units <= 120
    if any(marker in title_lower for marker in weak_title_markers) and line_count >= 3:
        return text_units <= 100
    if any(marker in text_lower for marker in AGENDA_MARKERS) and line_count >= 3 and text_units <= 120:
        return True
    return False


def is_cover_like_slide(number: int, title: str, text: str, line_count: int, text_units: int) -> bool:
    if number <= 2 and line_count <= 5 and text_units <= 45:
        return True
    if line_count <= 4 and text_units <= 35 and text_has_any_marker(f"{title}\n{text}", TITLE_MARKERS):
        return True
    return False


def is_section_like_slide(title: str, text: str, line_count: int, text_units: int) -> bool:
    if is_section_break_title(title) and text_units <= 70:
        return True
    if re.search(r"^\s*(part|chapter|section|module|unit)\s+[\w\divx.-]+", title, re.I) and text_units <= 70:
        return True
    if re.search(r"^\s*第\s*[一二三四五六七八九十0-9]+\s*[章节部分讲]\b", title) and text_units <= 70:
        return True
    return line_count <= 2 and text_units <= 16


def classify_slide_content(slide: Dict[str, Any]) -> Tuple[str, str]:
    number = int(slide.get("number", 0) or 0)
    title = plain(slide.get("title"))
    visible_items = slide_visible_text_items(slide)
    text = "\n".join(visible_items)
    text_units = count_cjk_or_words(text)
    line_count = len(visible_items)
    lower = text.lower()
    has_learning_signal = slide_has_structured_learning_signal(slide)

    if any(word in lower for word in ("练习", "习题", "作业", "quiz", "exercise", "homework", "problem set")):
        return "exercise", "练习/习题页保留为学习内容"
    if any(word in lower for word in ("例题", "举例", "示例", "example", "case study")):
        return "example", "例题/案例页保留为学习内容"
    if has_learning_signal:
        return "content", "包含公式、表格、图表或 OCR 视觉信息"
    if not text:
        return "content", "无结构化文字，保留截图供人工核对"
    if is_agenda_like_slide(title, text, line_count, text_units):
        return "agenda", "目录/大纲页默认压缩"
    if text_has_any_marker(text, CLOSING_MARKERS) and text_units <= 30:
        return "closing", "结束/答疑页默认压缩"
    if number <= 2 and is_cover_like_slide(number, title, text, line_count, text_units):
        return "title", "封面/标题页默认压缩"
    if is_section_like_slide(title, text, line_count, text_units):
        return "section", "章节过渡页默认压缩"
    if is_cover_like_slide(number, title, text, line_count, text_units):
        return "title", "封面/标题页默认压缩"
    return "content", "正文学习页"


def annotate_slide_kinds(extraction: Dict[str, Any], content_filter: str = "study") -> None:
    if content_filter not in {"study", "all"}:
        raise ValueError("--content-filter must be 'study' or 'all'")
    skipped: List[Dict[str, Any]] = []
    for slide in extraction.get("slides", []):
        kind, reason = classify_slide_content(slide)
        include = content_filter == "all" or kind not in NAVIGATION_CONTENT_KINDS
        slide["content_kind"] = kind
        slide["study_include"] = include
        slide["study_skip_reason"] = "" if include else reason
        if not include:
            skipped.append(
                {
                    "number": slide.get("number"),
                    "title": slide.get("title", ""),
                    "content_kind": kind,
                    "reason": reason,
                }
            )
    extraction["content_filter"] = content_filter
    extraction["study_slide_count"] = sum(1 for slide in extraction.get("slides", []) if slide.get("study_include", True))
    extraction["skipped_navigation_slide_count"] = len(skipped)
    extraction["skipped_navigation_slides"] = skipped


def study_slides(extraction: Dict[str, Any]) -> List[Dict[str, Any]]:
    return [slide for slide in extraction.get("slides", []) if slide.get("study_include", True)]


def note_required_slides(extraction: Dict[str, Any]) -> List[Dict[str, Any]]:
    return [slide for slide in study_slides(extraction) if slide.get("note_required", True)]


def deep_review_slides(extraction: Dict[str, Any]) -> List[Dict[str, Any]]:
    return [slide for slide in study_slides(extraction) if slide.get("review_tier") == "deep"]


def quick_review_slides(extraction: Dict[str, Any]) -> List[Dict[str, Any]]:
    return [slide for slide in study_slides(extraction) if slide.get("review_tier") == "quick"]


def reference_review_slides(extraction: Dict[str, Any]) -> List[Dict[str, Any]]:
    return [slide for slide in study_slides(extraction) if slide.get("review_tier") == "reference"]


def skipped_navigation_slides(extraction: Dict[str, Any]) -> List[Dict[str, Any]]:
    return [item for item in extraction.get("skipped_navigation_slides", []) if isinstance(item, dict)]


def slide_filter_summary(extraction: Dict[str, Any]) -> str:
    total = int(extraction.get("slide_count", len(extraction.get("slides", []))) or 0)
    study_count = int(extraction.get("study_slide_count", len(study_slides(extraction))) or 0)
    skipped = int(extraction.get("skipped_navigation_slide_count", 0) or 0)
    return f"学习页 {study_count}/{total}；已压缩导航页 {skipped} 页"


def review_plan_summary(extraction: Dict[str, Any]) -> str:
    depth = extraction.get("review_depth", "compressed")
    deep_count = int(extraction.get("deep_slide_count", len(deep_review_slides(extraction))) or 0)
    quick_count = int(extraction.get("quick_slide_count", len(quick_review_slides(extraction))) or 0)
    reference_count = int(extraction.get("reference_slide_count", len(reference_review_slides(extraction))) or 0)
    duplicate_count = int(extraction.get("duplicate_slide_count", 0) or 0)
    return (
        f"{REVIEW_DEPTH_LABELS.get(depth, depth)}：必读深讲 {deep_count} 页，"
        f"快速扫读 {quick_count} 页，参考/重复 {reference_count} 页；识别重复 {duplicate_count} 页"
    )


def slide_signature_text(slide: Dict[str, Any]) -> str:
    parts = as_list(slide.get("title")) + as_list(slide.get("text")) + as_list(slide.get("formula_candidates"))
    for table in as_list(slide.get("tables")):
        parts.append(plain(table))
    return plain(parts).lower()


def signature_tokens(text: str) -> List[str]:
    tokens = re.findall(r"[a-z0-9_]{2,}|[\u4e00-\u9fff]{2,}", text.lower())
    return [token for token in tokens if token not in TOPIC_STOPWORDS]


def token_similarity(left: List[str], right: List[str]) -> float:
    left_set = set(left)
    right_set = set(right)
    if not left_set or not right_set:
        return 0.0
    return len(left_set & right_set) / len(left_set | right_set)


def slide_has_exercise_signal(slide: Dict[str, Any]) -> bool:
    text = slide_signature_text(slide)
    return bool(
        slide.get("content_kind") == "exercise"
        or any(marker in text for marker in ("练习", "习题", "作业", "quiz", "exercise", "homework", "problem"))
    )


def slide_has_example_signal(slide: Dict[str, Any]) -> bool:
    text = slide_signature_text(slide)
    return bool(slide.get("content_kind") == "example" or any(marker in text for marker in ("例题", "举例", "示例", "example", "case")))


def slide_importance(slide: Dict[str, Any]) -> Tuple[int, List[str]]:
    score = 10
    reasons: List[str] = []
    text_units = count_cjk_or_words(slide_signature_text(slide))
    if slide_has_exercise_signal(slide):
        score += 45
        reasons.append("练习/题型")
    if slide_has_example_signal(slide):
        score += 38
        reasons.append("例题/案例")
    if slide_has_formula_signal(slide):
        score += 36
        reasons.append("公式/推导")
    if slide.get("tables"):
        score += 16
        reasons.append("表格")
    if slide.get("related_objects") or slide.get("ocr_visual_explanations"):
        score += 16
        reasons.append("图表/结构")
    if slide.get("images"):
        score += 8
        reasons.append("图片")
    if text_units >= 140:
        score += 8
        reasons.append("信息密度高")
    elif text_units <= 35:
        score -= 8
        reasons.append("信息少")
    return score, unique_preserve(reasons)


def review_depth_budgets(count: int, review_depth: str, max_deep_slides: Optional[int]) -> Tuple[int, int]:
    if count <= 0:
        return 0, 0
    if review_depth == "complete":
        return count, count
    if review_depth == "balanced":
        deep = max(10, min(42, round(count * 0.34)))
        quick = max(deep, round(count * 0.68), deep + min(18, max(0, count - deep)))
    else:
        deep = max(8, min(24, round(count * 0.22)))
        quick = max(deep, round(count * 0.45), deep + min(12, max(0, (count - deep + 1) // 2)))
    if max_deep_slides is not None:
        deep = max(1, min(deep, max_deep_slides, count))
        quick = max(deep, min(quick, count))
    return min(deep, count), min(quick, count)


def annotate_review_plan(
    extraction: Dict[str, Any],
    review_depth: str = "compressed",
    max_deep_slides: Optional[int] = None,
    dedupe_threshold: float = 0.86,
) -> None:
    if review_depth not in REVIEW_DEPTH_LABELS:
        raise ValueError("--review-depth must be compressed, balanced, or complete")
    slides = study_slides(extraction)
    signatures: List[Tuple[int, List[str]]] = []
    duplicate_count = 0
    for slide in slides:
        number = int(slide.get("number", 0) or 0)
        tokens = signature_tokens(slide_signature_text(slide))
        best_number = 0
        best_similarity = 0.0
        for previous_number, previous_tokens in signatures[-10:]:
            similarity = token_similarity(tokens, previous_tokens)
            if similarity > best_similarity:
                best_similarity = similarity
                best_number = previous_number
        if best_similarity >= dedupe_threshold and tokens:
            slide["duplicate_of"] = best_number
            slide["duplicate_similarity"] = round(best_similarity, 3)
            duplicate_count += 1
        else:
            slide["duplicate_of"] = None
            slide["duplicate_similarity"] = 0.0
        signatures.append((number, tokens))

        score, reasons = slide_importance(slide)
        if slide.get("duplicate_of"):
            score -= 24
            reasons.append(f"与 S{slide.get('duplicate_of')} 重复")
        slide["importance_score"] = score
        slide["importance_reasons"] = reasons

    deep_budget, quick_budget = review_depth_budgets(len(slides), review_depth, max_deep_slides)
    ranked = sorted(slides, key=lambda item: (-int(item.get("importance_score", 0) or 0), int(item.get("number", 0) or 0)))
    deep_numbers = {int(slide.get("number", 0) or 0) for slide in ranked[:deep_budget]}
    quick_numbers = {int(slide.get("number", 0) or 0) for slide in ranked[:quick_budget]}

    for slide in slides:
        number = int(slide.get("number", 0) or 0)
        if review_depth == "complete" or number in deep_numbers:
            tier = "deep"
        elif number in quick_numbers:
            tier = "quick"
        else:
            tier = "reference"
        slide["review_tier"] = tier
        slide["review_tier_label"] = REVIEW_TIER_LABELS[tier]
        slide["note_required"] = tier in {"deep", "quick"}
        if tier == "deep":
            guidance = "深讲：只解释本页真正新增的概念、公式、例题或易错点，不复述整页文字。"
        elif tier == "quick":
            guidance = "速读：用 2-4 条 bullet 说明核心结论、考试信号和是否需要回看。"
        else:
            guidance = "参考：通常不写逐页讲解；只在学习路径中标出用途，必要时回看截图或 extraction.json。"
        if slide.get("duplicate_of"):
            guidance += f" 本页与 S{slide.get('duplicate_of')} 高度相似，避免重复讲解。"
        slide["compression_guidance"] = guidance

    extraction["review_depth"] = review_depth
    extraction["deep_slide_count"] = len(deep_review_slides(extraction))
    extraction["quick_slide_count"] = len(quick_review_slides(extraction))
    extraction["reference_slide_count"] = len(reference_review_slides(extraction))
    extraction["note_required_slide_count"] = len(note_required_slides(extraction))
    extraction["duplicate_slide_count"] = duplicate_count
    extraction["dedupe_threshold"] = dedupe_threshold


def apply_ocr_json(extraction: Dict[str, Any], ocr_path: Optional[Path]) -> None:
    if not ocr_path:
        return
    raw = json.loads(ocr_path.read_text(encoding="utf-8"))
    items = raw if isinstance(raw, list) else raw.get("slides", [])
    by_number = {int(slide.get("number", 0) or 0): slide for slide in extraction.get("slides", [])}
    merged = 0
    for item in items:
        if not isinstance(item, dict):
            continue
        try:
            number = int(item.get("number") or item.get("slide") or item.get("page"))
        except Exception:
            continue
        slide = by_number.get(number)
        if not slide:
            continue
        for key in ("text", "formula_candidates", "alt_texts"):
            values = as_list(item.get(key))
            if values:
                slide[key] = unique_preserve(list(slide.get(key) or []) + [plain(value) for value in values if plain(value)])
        if item.get("formulas"):
            slide["formulas"] = list(slide.get("formulas") or []) + as_list(item.get("formulas"))
        if item.get("visual_explanation"):
            slide.setdefault("ocr_visual_explanations", []).append(plain(item.get("visual_explanation")))
        merged += 1
    extraction.setdefault("warnings", []).append(f"Merged OCR/math-recognition JSON from {ocr_path} for {merged} slide(s).")


def make_notes_template(extraction: Dict[str, Any], language: str) -> Dict[str, Any]:
    slides = []
    for slide in note_required_slides(extraction):
        slides.append(
            {
                "number": slide.get("number"),
                "title": slide.get("title", ""),
                "content_kind": slide.get("content_kind", "content"),
                "review_tier": slide.get("review_tier", "deep"),
                "review_tier_label": slide.get("review_tier_label", "必读深讲"),
                "importance_reasons": slide.get("importance_reasons", []),
                "duplicate_of": slide.get("duplicate_of"),
                "compression_guidance": slide.get("compression_guidance", ""),
                "purpose": "",
                "what_it_says": "",
                "detailed_explanation": "",
                "visual_explanation": "",
                "formula_explanations": [],
                "worked_examples": [],
                "uncertainties": [],
                "exam_focus": "",
                "key_takeaways": [],
                "memory_hooks": [],
                "likely_questions": [],
                "practice_questions": [],
                "common_mistakes": [],
                "prerequisites": [],
                "difficulty": "",
                "estimated_review_minutes": "",
                "tags": [],
            }
        )
    return {
        "language": language,
        "source": extraction.get("source", ""),
        "content_filter": extraction.get("content_filter", "study"),
        "review_depth": extraction.get("review_depth", "compressed"),
        "study_slide_count": extraction.get("study_slide_count", len(study_slides(extraction))),
        "note_required_slide_count": len(slides),
        "deep_slide_count": extraction.get("deep_slide_count", len(deep_review_slides(extraction))),
        "quick_slide_count": extraction.get("quick_slide_count", len(quick_review_slides(extraction))),
        "reference_slide_count": extraction.get("reference_slide_count", len(reference_review_slides(extraction))),
        "total_slide_count": extraction.get("slide_count", len(extraction.get("slides", []))),
        "skipped_navigation_slides": skipped_navigation_slides(extraction),
        "reference_slides": [
            {
                "number": slide.get("number"),
                "title": slide.get("title", ""),
                "duplicate_of": slide.get("duplicate_of"),
                "importance_reasons": slide.get("importance_reasons", []),
            }
            for slide in reference_review_slides(extraction)
        ],
        "slides": slides,
    }


def load_notes(path: Optional[Path]) -> Dict[int, Dict[str, Any]]:
    if not path:
        return {}
    raw = json.loads(path.read_text(encoding="utf-8"))
    slide_items = raw if isinstance(raw, list) else raw.get("slides", [])
    notes: Dict[int, Dict[str, Any]] = {}
    for item in slide_items:
        if not isinstance(item, dict):
            continue
        number = item.get("number") or item.get("slide") or item.get("page")
        try:
            n = int(number)
        except Exception:
            continue
        notes[n] = item
    return notes


def meaningful_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, list):
        return "\n".join(meaningful_text(item) for item in value).strip()
    if isinstance(value, dict):
        return json.dumps(value, ensure_ascii=False)
    return str(value).strip()


def count_cjk_or_words(text: str) -> int:
    cjk = re.findall(r"[\u4e00-\u9fff]", text)
    words = re.findall(r"[A-Za-z0-9_]+", text)
    return len(cjk) + len(words)


def has_banned_fluff(value: Any) -> bool:
    text = meaningful_text(value)
    return any(re.search(pattern, text) for pattern in BANNED_FLUFF_PATTERNS)


def has_specific_terms(value: Any, slide: Dict[str, Any]) -> bool:
    text = meaningful_text(value)
    source_terms = []
    for item in slide.get("text") or []:
        source_terms.extend(re.findall(r"[A-Za-z][A-Za-z0-9_-]{2,}|[\u4e00-\u9fff]{2,}", str(item)))
    source_terms = [term for term in dict.fromkeys(source_terms) if len(term) >= 2]
    hits = sum(1 for term in source_terms[:30] if term in text)
    return hits >= 2 or bool(slide.get("formula_candidates") and any(str(c) in text for c in slide.get("formula_candidates") or []))


def field_has_content(note: Dict[str, Any], min_chars: int, *keys: str) -> bool:
    return len(meaningful_text(note_field(note, *keys))) >= min_chars


def field_is_dense(note: Dict[str, Any], min_units: int, *keys: str) -> bool:
    return count_cjk_or_words(meaningful_text(note_field(note, *keys))) >= min_units


def slide_has_formula(slide: Dict[str, Any]) -> bool:
    return bool(slide.get("formulas") or slide.get("formula_candidates"))


def slide_has_visual(slide: Dict[str, Any]) -> bool:
    return bool(slide.get("images") or slide.get("related_objects") or slide.get("tables") or slide.get("screenshot"))


def build_quality_report(
    extraction: Dict[str, Any],
    notes: Dict[int, Dict[str, Any]],
    fail_under: Optional[float] = None,
    study_mode: str = "notes",
) -> Dict[str, Any]:
    slide_reports = []
    total_points = 0
    earned_points = 0
    slides_to_score = note_required_slides(extraction)
    for slide in slides_to_score:
        number = int(slide.get("number", 0) or 0)
        note = notes.get(number, {})
        tier = slide.get("review_tier", "deep")
        checks = [
            ("purpose", "这一页是干什么用的", field_is_dense(note, 18 if tier == "deep" else 10, "purpose", "page_purpose")),
            ("what_it_says", "这一页讲了什么", field_is_dense(note, 45 if tier == "deep" else 24, "what_it_says", "summary")),
            ("specificity", "讲解引用本页具体术语", has_specific_terms(note_field(note, "what_it_says", "detailed_explanation", "summary"), slide)),
            ("anti_fluff", "讲解不能是套话/空话", not has_banned_fluff(note)),
        ]
        if tier == "deep":
            checks.append(("detailed_explanation", "复杂内容详解", field_is_dense(note, 90, "detailed_explanation", "complex_explanation")))
        elif slide_has_formula(slide) or slide_has_visual(slide):
            checks.append(("compressed_explanation", "速读页保留核心解释", field_is_dense(note, 35, "detailed_explanation", "complex_explanation", "what_it_says", "summary")))
        if slide_has_visual(slide):
            checks.append(("visual_explanation", "图片/图表/表格说明", field_is_dense(note, 35 if tier == "deep" else 18, "visual_explanation", "image_explanation")))
        if slide_has_formula(slide):
            checks.append(("formula_explanations", "公式说明", bool(note_field(note, "formula_explanations", "formulas"))))
            checks.append(("worked_examples", "公式或方法例题", bool(note_field(note, "worked_examples", "examples"))))
        if study_mode == "final":
            checks.extend(
                [
                    ("exam_focus", "期末考点定位", field_is_dense(note, 25, "exam_focus")),
                    ("key_takeaways", "核心记忆点", bool(note_field(note, "key_takeaways"))),
                    ("likely_questions", "可能考法/自测题", bool(note_field(note, "likely_questions"))),
                    ("common_mistakes", "常见错误", bool(note_field(note, "common_mistakes"))),
                ]
            )
        missing = [label for _, label, ok in checks if not ok]
        points = len(checks)
        earned = sum(1 for _, _, ok in checks if ok)
        total_points += points
        earned_points += earned
        slide_reports.append(
            {
                "number": number,
                "title": note_field(note, "title") or slide.get("title", ""),
                "score": round((earned / points) * 100, 2) if points else 100.0,
                "earned": earned,
                "possible": points,
                "missing": missing,
                "has_formula": slide_has_formula(slide),
                "has_visual": slide_has_visual(slide),
                "review_tier": tier,
            }
        )

    score = round((earned_points / total_points) * 100, 2) if total_points else 0.0
    passed = fail_under is None or score >= fail_under
    return {
        "tool": "summarize-ppt-notes",
        "version": VERSION,
        "source": extraction.get("source", ""),
        "slide_count": len(slides_to_score),
        "total_slide_count": extraction.get("slide_count", len(extraction.get("slides", []))),
        "study_slide_count": extraction.get("study_slide_count", len(study_slides(extraction))),
        "note_required_slide_count": extraction.get("note_required_slide_count", len(slides_to_score)),
        "content_filter": extraction.get("content_filter", "study"),
        "review_depth": extraction.get("review_depth", "compressed"),
        "deep_slide_count": extraction.get("deep_slide_count", len(deep_review_slides(extraction))),
        "quick_slide_count": extraction.get("quick_slide_count", len(quick_review_slides(extraction))),
        "reference_slide_count": extraction.get("reference_slide_count", len(reference_review_slides(extraction))),
        "duplicate_slide_count": extraction.get("duplicate_slide_count", 0),
        "skipped_navigation_slide_count": extraction.get("skipped_navigation_slide_count", 0),
        "skipped_navigation_slides": skipped_navigation_slides(extraction),
        "score": score,
        "earned": earned_points,
        "possible": total_points,
        "fail_under": fail_under,
        "study_mode": study_mode,
        "passed": passed,
        "slides": slide_reports,
    }


def write_quality_report(report: Dict[str, Any], json_path: Path, md_path: Path) -> None:
    json_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    lines = [
        "# Quality Report",
        "",
        f"- Source: {report.get('source', '')}",
        f"- Score: {report.get('score', 0)}%",
        f"- Note-required slides evaluated: {report.get('slide_count', 0)} / study slides {report.get('study_slide_count', report.get('slide_count', 0))} / original slides {report.get('total_slide_count', report.get('slide_count', 0))}",
        f"- Review depth: {report.get('review_depth', 'compressed')}",
        f"- Reading tiers: deep {report.get('deep_slide_count', 0)}, quick {report.get('quick_slide_count', 0)}, reference {report.get('reference_slide_count', 0)}, duplicates {report.get('duplicate_slide_count', 0)}",
        f"- Skipped navigation slides: {report.get('skipped_navigation_slide_count', 0)}",
        f"- Passed: {report.get('passed')}",
        "",
    ]
    skipped = report.get("skipped_navigation_slides") or []
    if skipped:
        lines += ["## Skipped Navigation Slides", ""]
        for item in skipped:
            lines.append(f"- S{item.get('number')}: {item.get('content_kind')} - {item.get('title', '')}")
        lines.append("")
    lines += ["## Slide Checks", ""]
    for slide in report.get("slides", []):
        missing = slide.get("missing") or []
        status = "PASS" if not missing else "TODO"
        lines.append(f"### Slide {slide.get('number')} [{slide.get('review_tier', 'deep')}]: {status} ({slide.get('score')}%)")
        lines.append("")
        if missing:
            for item in missing:
                lines.append(f"- Missing: {item}")
        else:
            lines.append("- All required explanation fields are covered.")
        lines.append("")
    md_path.write_text("\n".join(lines), encoding="utf-8")


def content_type_for(path: Path) -> str:
    ext = path.suffix.lower().lstrip(".")
    explicit = {
        "png": "image/png",
        "jpg": "image/jpeg",
        "jpeg": "image/jpeg",
        "gif": "image/gif",
        "bmp": "image/bmp",
        "tif": "image/tiff",
        "tiff": "image/tiff",
        "emf": "image/x-emf",
        "wmf": "image/x-wmf",
    }
    return explicit.get(ext) or mimetypes.guess_type(path.name)[0] or "application/octet-stream"


def image_size(path: Path) -> Tuple[int, int]:
    try:
        from PIL import Image  # type: ignore

        with Image.open(path) as img:
            return img.size
    except Exception:
        return (1280, 720)


def emu_dimensions(path: Path, max_width_inches: float) -> Tuple[int, int]:
    width_px, height_px = image_size(path)
    if width_px <= 0 or height_px <= 0:
        width_px, height_px = 1280, 720
    width_inches = min(max_width_inches, max(1.0, width_px / 160.0))
    height_inches = width_inches * (height_px / width_px)
    return int(width_inches * 914400), int(height_inches * 914400)


def w_text(text: Any) -> str:
    return escape(str(text), {'"': "&quot;"})


def strip_latex_delimiters(text: str) -> str:
    text = text.strip()
    if text.startswith("$$") and text.endswith("$$"):
        return text[2:-2].strip()
    if text.startswith("$") and text.endswith("$"):
        return text[1:-1].strip()
    if text.startswith(r"\[") and text.endswith(r"\]"):
        return text[2:-2].strip()
    if text.startswith(r"\(") and text.endswith(r"\)"):
        return text[2:-2].strip()
    return text


def parse_latex_group(text: str, start: int) -> Tuple[Optional[str], int]:
    while start < len(text) and text[start].isspace():
        start += 1
    if start >= len(text) or text[start] != "{":
        return None, start
    depth = 0
    for index in range(start, len(text)):
        char = text[index]
        if char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                return text[start + 1 : index], index + 1
    return None, start


def replace_latex_command_groups(text: str, command: str, arity: int, formatter: Any) -> str:
    position = 0
    while True:
        index = text.find(command, position)
        if index < 0:
            return text
        cursor = index + len(command)
        groups = []
        ok = True
        for _ in range(arity):
            group, cursor = parse_latex_group(text, cursor)
            if group is None:
                ok = False
                break
            groups.append(group)
        if not ok:
            position = index + len(command)
            continue
        replacement = formatter(*groups)
        text = text[:index] + replacement + text[cursor:]
        position = index + len(replacement)


def to_superscript(text: Any) -> str:
    raw = plain(text)
    converted = raw.translate(SUPERSCRIPT_CHARS)
    return converted if converted != raw or len(raw) <= 3 else f"^({raw})"


def to_subscript(text: Any) -> str:
    raw = plain(text)
    converted = raw.translate(SUBSCRIPT_CHARS)
    return converted if converted != raw or len(raw) <= 3 else f"_({raw})"


def compact_fraction(num: str, den: str) -> str:
    left = latex_to_readable_formula(num)
    right = latex_to_readable_formula(den)
    if re.fullmatch(r"[\wα-ωΑ-Ω₀-₉ᵢⱼ₊₋₌]+", left) and re.fullmatch(r"[\wα-ωΑ-Ω₀-₉ᵢⱼ₊₋₌]+", right):
        return f"{left}⁄{right}"
    return f"({left})⁄({right})"


def decorated_symbol(kind: str, value: str) -> str:
    text = latex_to_readable_formula(value).strip()
    if not text:
        return ""
    if kind == "hat":
        return {"y": "ŷ", "x": "x̂", "p": "p̂", "q": "q̂"}.get(text, text + "\u0302")
    if kind == "bar":
        return text + "\u0304"
    if kind == "tilde":
        return text + "\u0303"
    return text


def latex_to_readable_formula(value: Any) -> str:
    text = strip_latex_delimiters(plain(value))
    if not text:
        return ""
    text = text.replace(r"\left", "").replace(r"\right", "")
    text = text.replace(r"\,", " ").replace(r"\;", " ").replace(r"\!", "")
    for _ in range(8):
        updated = replace_latex_command_groups(text, r"\frac", 2, compact_fraction)
        updated = replace_latex_command_groups(updated, r"\sqrt", 1, lambda body: f"√({latex_to_readable_formula(body)})")
        updated = replace_latex_command_groups(updated, r"\hat", 1, lambda body: decorated_symbol("hat", body))
        updated = replace_latex_command_groups(updated, r"\bar", 1, lambda body: decorated_symbol("bar", body))
        updated = replace_latex_command_groups(updated, r"\tilde", 1, lambda body: decorated_symbol("tilde", body))
        if updated == text:
            break
        text = updated
    text = re.sub(r"\\hat\s+([A-Za-z])", lambda m: decorated_symbol("hat", m.group(1)), text)
    text = re.sub(r"\\bar\s+([A-Za-z])", lambda m: decorated_symbol("bar", m.group(1)), text)
    text = re.sub(r"\\tilde\s+([A-Za-z])", lambda m: decorated_symbol("tilde", m.group(1)), text)
    for command, replacement in sorted(LATEX_SYMBOLS.items(), key=lambda item: -len(item[0])):
        text = text.replace(command, replacement)
    text = re.sub(r"_\{([^{}]+)\}", lambda match: to_subscript(match.group(1)), text)
    text = re.sub(r"\^\{([^{}]+)\}", lambda match: to_superscript(match.group(1)), text)
    text = re.sub(r"_([A-Za-z0-9+\-=()])", lambda match: to_subscript(match.group(1)), text)
    text = re.sub(r"\^([A-Za-z0-9+\-=()])", lambda match: to_superscript(match.group(1)), text)
    text = text.replace("{", "").replace("}", "")
    text = re.sub(r"\\([A-Za-z]+)", r"\1", text)
    text = text.replace("\\", "")
    text = text.replace(" - ", " − ").replace("-", "−")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def omml_run(text: Any) -> str:
    return f"<m:r><m:t>{w_text(text)}</m:t></m:r>"


def omml_group(xml: str) -> str:
    return f"<m:e>{xml or omml_run('□')}</m:e>"


def parse_latex_script(text: str, start: int) -> Tuple[str, int]:
    while start < len(text) and text[start].isspace():
        start += 1
    if start >= len(text):
        return "", start
    if text[start] == "{":
        group, end = parse_latex_group(text, start)
        if group is None:
            return "", start
        return latex_to_omml_inner(group), end
    if text[start] == "\\":
        match = re.match(r"\\[A-Za-z]+", text[start:])
        if match:
            raw = match.group(0)
            return latex_to_omml_inner(raw), start + len(raw)
    return latex_to_omml_inner(text[start]), start + 1


def omml_apply_scripts(base_xml: str, sub_xml: str, sup_xml: str) -> str:
    if sub_xml and sup_xml:
        return f"<m:sSubSup>{omml_group(base_xml)}<m:sub>{sub_xml}</m:sub><m:sup>{sup_xml}</m:sup></m:sSubSup>"
    if sub_xml:
        return f"<m:sSub>{omml_group(base_xml)}<m:sub>{sub_xml}</m:sub></m:sSub>"
    if sup_xml:
        return f"<m:sSup>{omml_group(base_xml)}<m:sup>{sup_xml}</m:sup></m:sSup>"
    return base_xml


def parse_latex_atom_to_omml(text: str, start: int) -> Tuple[str, int]:
    if text.startswith(r"\frac", start):
        cursor = start + len(r"\frac")
        numerator, cursor = parse_latex_group(text, cursor)
        denominator, cursor = parse_latex_group(text, cursor)
        if numerator is not None and denominator is not None:
            return (
                f"<m:f><m:num>{latex_to_omml_inner(numerator)}</m:num><m:den>{latex_to_omml_inner(denominator)}</m:den></m:f>",
                cursor,
            )
    if text.startswith(r"\sqrt", start):
        cursor = start + len(r"\sqrt")
        body, cursor = parse_latex_group(text, cursor)
        if body is not None:
            return f"<m:rad><m:radPr><m:degHide m:val=\"1\"/></m:radPr>{omml_group(latex_to_omml_inner(body))}</m:rad>", cursor
    for command, kind in [(r"\hat", "hat"), (r"\bar", "bar"), (r"\tilde", "tilde")]:
        if text.startswith(command, start):
            cursor = start + len(command)
            body, cursor = parse_latex_group(text, cursor)
            if body is None and cursor < len(text):
                body = text[cursor]
                cursor += 1
            return omml_run(decorated_symbol(kind, body or "")), cursor
    if text[start] == "\\":
        match = re.match(r"\\[A-Za-z]+", text[start:])
        if match:
            raw = match.group(0)
            return omml_run(LATEX_SYMBOLS.get(raw, raw[1:])), start + len(raw)
    if text[start] == "{":
        group, cursor = parse_latex_group(text, start)
        if group is not None:
            return latex_to_omml_inner(group), cursor
    end = start
    while end < len(text) and text[end] not in "\\_^{}":
        end += 1
    if end == start:
        return omml_run(text[start]), start + 1
    return omml_run(text[start:end]), end


def latex_to_omml_inner(value: Any) -> str:
    text = strip_latex_delimiters(plain(value))
    if not text:
        return ""
    text = text.replace(r"\left", "").replace(r"\right", "")
    text = text.replace(r"\,", " ").replace(r"\;", " ").replace(r"\!", "")
    parts: List[str] = []
    index = 0
    while index < len(text):
        if text[index].isspace():
            start = index
            while index < len(text) and text[index].isspace():
                index += 1
            parts.append(omml_run(text[start:index]))
            continue
        base_xml, index = parse_latex_atom_to_omml(text, index)
        sub_xml = ""
        sup_xml = ""
        while index < len(text) and text[index] in "_^":
            marker = text[index]
            script_xml, index = parse_latex_script(text, index + 1)
            if marker == "_":
                sub_xml = script_xml
            else:
                sup_xml = script_xml
        parts.append(omml_apply_scripts(base_xml, sub_xml, sup_xml))
    return "".join(parts)


def latex_to_omml_paragraph(value: Any) -> str:
    formula = plain(value)
    if not formula:
        formula = "□"
    math_xml = latex_to_omml_inner(formula)
    if not math_xml:
        math_xml = omml_run(latex_to_readable_formula(formula) or formula)
    return (
        '<w:p><w:pPr><w:pStyle w:val="FormulaDisplay"/></w:pPr>'
        '<m:oMathPara><m:oMath>'
        f"{math_xml}"
        "</m:oMath></m:oMathPara></w:p>"
    )


class DocxBuilder:
    def __init__(self, title: str) -> None:
        self.title = title
        self.body: List[str] = []
        self.image_rels: List[Tuple[str, str, Path]] = []
        self.image_id = 1

    def add_heading(self, text: str, level: int = 1) -> None:
        style = "Heading1" if level == 1 else "Heading2"
        self.add_paragraph(text, style=style)

    def add_title(self, title: str, subtitle: str = "") -> None:
        self.add_paragraph(title, style="Title")
        if subtitle:
            self.add_paragraph(subtitle, style="Subtitle")

    def add_meta(self, text: Any) -> None:
        self.add_paragraph(text, style="Meta")

    def add_page_break(self) -> None:
        self.body.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')

    def add_section_label(self, text: str) -> None:
        self.add_paragraph(text, style="SectionLabel")

    def add_callout(self, title: str, value: Any, style: str = "InsightBox", placeholder: str = "无") -> None:
        self.add_paragraph(title, style=style, bold=True)
        self.add_structured_value(value, placeholder=placeholder, bullet_style=f"{style}Text", paragraph_style=f"{style}Text")

    def add_structured_value(
        self,
        value: Any,
        placeholder: str = "无",
        bullet_style: str = "ListBullet",
        paragraph_style: Optional[str] = None,
    ) -> None:
        if isinstance(value, list):
            if value:
                for item in value:
                    self.add_bullet(item if not isinstance(item, dict) else json.dumps(item, ensure_ascii=False), style=bullet_style)
            else:
                self.add_paragraph(placeholder, style=paragraph_style)
        elif value:
            self.add_paragraph(value, style=paragraph_style)
        else:
            self.add_paragraph(placeholder, style=paragraph_style)

    def add_paragraph(self, text: Any = "", style: Optional[str] = None, bold: bool = False) -> None:
        text = "" if text is None else str(text)
        style_xml = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
        run_props = "<w:rPr><w:b/></w:rPr>" if bold else ""
        lines = text.splitlines() or [""]
        run_parts = []
        for index, line in enumerate(lines):
            if index:
                run_parts.append("<w:br/>")
            run_parts.append(f'<w:t xml:space="preserve">{w_text(line)}</w:t>')
        self.body.append(f"<w:p>{style_xml}<w:r>{run_props}{''.join(run_parts)}</w:r></w:p>")

    def add_formula_display(self, formula: Any) -> None:
        self.body.append(latex_to_omml_paragraph(formula))

    def add_bullet(self, text: Any, style: str = "ListBullet") -> None:
        self.add_paragraph(f"- {text}", style=style)

    def add_image(self, path_text: str, max_width_inches: float = 6.4) -> bool:
        if not path_text:
            return False
        path = Path(path_text)
        if not path.exists() or not path.is_file():
            return False
        rid = f"rId{len(self.image_rels) + 1}"
        ext = path.suffix.lower() or ".png"
        media_name = f"image{len(self.image_rels) + 1}{ext}"
        cx, cy = emu_dimensions(path, max_width_inches)
        doc_pr_id = self.image_id
        self.image_id += 1
        self.image_rels.append((rid, media_name, path))
        self.body.append(
            f"""
<w:p>
  <w:r>
    <w:drawing>
      <wp:inline distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="{cx}" cy="{cy}"/>
        <wp:docPr id="{doc_pr_id}" name="{w_text(path.name)}"/>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic>
              <pic:nvPicPr>
                <pic:cNvPr id="0" name="{w_text(path.name)}"/>
                <pic:cNvPicPr/>
              </pic:nvPicPr>
              <pic:blipFill>
                <a:blip r:embed="{rid}"/>
                <a:stretch><a:fillRect/></a:stretch>
              </pic:blipFill>
              <pic:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              </pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>
</w:p>
""".strip()
        )
        return True

    def document_xml(self) -> str:
        ns = " ".join(f'xmlns:{k}="{v}"' for k, v in WORD_NS.items())
        body = "\n".join(self.body)
        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document {ns}>
  <w:body>
    {body}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="720" w:right="820" w:bottom="720" w:left="820" w:header="560" w:footer="560" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>'''

    def rels_xml(self) -> str:
        rels = [
            f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/{media_name}"/>'
            for rid, media_name, _ in self.image_rels
        ]
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<Relationships xmlns="{NS_REL}">' + "".join(rels) + "</Relationships>"
        )

    def content_types_xml(self) -> str:
        defaults = {
            "rels": "application/vnd.openxmlformats-package.relationships+xml",
            "xml": "application/xml",
        }
        for _, media_name, path in self.image_rels:
            defaults[Path(media_name).suffix.lower().lstrip(".")] = content_type_for(path)
        default_xml = "".join(
            f'<Default Extension="{w_text(ext)}" ContentType="{w_text(content_type)}"/>'
            for ext, content_type in sorted(defaults.items())
        )
        overrides = """
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
""".strip()
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            f"{default_xml}{overrides}</Types>"
        )

    def write(self, output: Path) -> None:
        output.parent.mkdir(parents=True, exist_ok=True)
        now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as docx:
            docx.writestr("[Content_Types].xml", self.content_types_xml())
            docx.writestr("_rels/.rels", root_rels_xml())
            docx.writestr("docProps/core.xml", core_xml(self.title, now))
            docx.writestr("docProps/app.xml", app_xml())
            docx.writestr("word/document.xml", self.document_xml())
            docx.writestr("word/styles.xml", styles_xml())
            docx.writestr("word/_rels/document.xml.rels", self.rels_xml())
            for _, media_name, path in self.image_rels:
                docx.write(path, f"word/media/{media_name}")


def root_rels_xml() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{NS_REL}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''


def core_xml(title: str, now: str) -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:dcmitype="http://purl.org/dc/dcmitype/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>{w_text(title)}</dc:title>
  <dc:creator>summarize-ppt-notes</dc:creator>
  <cp:lastModifiedBy>summarize-ppt-notes</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>'''


def app_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>summarize-ppt-notes</Application>
</Properties>'''


def styles_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:pPr><w:spacing w:before="0" w:after="120" w:line="330" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Aptos" w:eastAsia="Microsoft YaHei" w:hAnsi="Aptos"/><w:color w:val="1F2937"/><w:sz w:val="21"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="240" w:after="120"/><w:jc w:val="center"/></w:pPr>
    <w:rPr><w:b/><w:rFonts w:ascii="Aptos Display" w:eastAsia="Microsoft YaHei" w:hAnsi="Aptos Display"/><w:color w:val="0F172A"/><w:sz w:val="44"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Subtitle">
    <w:name w:val="Subtitle"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:after="360"/><w:jc w:val="center"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Aptos" w:eastAsia="Microsoft YaHei" w:hAnsi="Aptos"/><w:color w:val="475569"/><w:sz w:val="23"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Meta">
    <w:name w:val="Meta"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="0" w:after="80"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Aptos" w:eastAsia="Microsoft YaHei" w:hAnsi="Aptos"/><w:color w:val="64748B"/><w:sz w:val="18"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr><w:spacing w:before="360" w:after="160"/><w:keepNext/><w:outlineLvl w:val="0"/><w:pBdr><w:bottom w:val="single" w:sz="8" w:space="3" w:color="2563EB"/></w:pBdr></w:pPr>
    <w:rPr><w:b/><w:rFonts w:ascii="Aptos Display" w:eastAsia="Microsoft YaHei" w:hAnsi="Aptos Display"/><w:color w:val="1E3A8A"/><w:sz w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr><w:spacing w:before="220" w:after="90"/><w:keepNext/><w:outlineLvl w:val="1"/></w:pPr>
    <w:rPr><w:b/><w:rFonts w:ascii="Aptos" w:eastAsia="Microsoft YaHei" w:hAnsi="Aptos"/><w:color w:val="0F766E"/><w:sz w:val="24"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="SectionLabel">
    <w:name w:val="Section Label"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="180" w:after="70"/></w:pPr>
    <w:rPr><w:b/><w:caps/><w:rFonts w:ascii="Aptos" w:eastAsia="Microsoft YaHei" w:hAnsi="Aptos"/><w:color w:val="334155"/><w:sz w:val="18"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListBullet">
    <w:name w:val="List Bullet"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:after="70"/><w:ind w:left="360" w:hanging="180"/></w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="SourceText">
    <w:name w:val="Source Text"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:after="60"/><w:shd w:val="clear" w:color="auto" w:fill="F8FAFC"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Aptos" w:eastAsia="Microsoft YaHei" w:hAnsi="Aptos"/><w:color w:val="475569"/><w:sz w:val="18"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="InsightBox">
    <w:name w:val="Insight Box"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="120" w:after="40"/><w:shd w:val="clear" w:color="auto" w:fill="EFF6FF"/><w:pBdr><w:left w:val="single" w:sz="18" w:space="5" w:color="2563EB"/></w:pBdr></w:pPr>
    <w:rPr><w:b/><w:rFonts w:ascii="Aptos" w:eastAsia="Microsoft YaHei" w:hAnsi="Aptos"/><w:color w:val="1D4ED8"/><w:sz w:val="21"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="InsightBoxText">
    <w:name w:val="Insight Box Text"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:after="90"/><w:shd w:val="clear" w:color="auto" w:fill="EFF6FF"/><w:ind w:left="260"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Aptos" w:eastAsia="Microsoft YaHei" w:hAnsi="Aptos"/><w:color w:val="1E3A8A"/><w:sz w:val="20"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="FormulaBox">
    <w:name w:val="Formula Box"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="120" w:after="40"/><w:shd w:val="clear" w:color="auto" w:fill="ECFDF5"/><w:pBdr><w:left w:val="single" w:sz="18" w:space="5" w:color="059669"/></w:pBdr></w:pPr>
    <w:rPr><w:b/><w:rFonts w:ascii="Aptos" w:eastAsia="Microsoft YaHei" w:hAnsi="Aptos"/><w:color w:val="047857"/><w:sz w:val="21"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="FormulaBoxText">
    <w:name w:val="Formula Box Text"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:after="80"/><w:shd w:val="clear" w:color="auto" w:fill="ECFDF5"/><w:ind w:left="260"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Aptos" w:eastAsia="Microsoft YaHei" w:hAnsi="Aptos"/><w:color w:val="064E3B"/><w:sz w:val="20"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="FormulaDisplay">
    <w:name w:val="Formula Display"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="80" w:after="90"/><w:jc w:val="center"/><w:shd w:val="clear" w:color="auto" w:fill="ECFDF5"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Cambria Math" w:eastAsia="Microsoft YaHei" w:hAnsi="Cambria Math"/><w:color w:val="064E3B"/><w:sz w:val="26"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ExamBox">
    <w:name w:val="Exam Box"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="120" w:after="40"/><w:shd w:val="clear" w:color="auto" w:fill="FFFBEB"/><w:pBdr><w:left w:val="single" w:sz="18" w:space="5" w:color="D97706"/></w:pBdr></w:pPr>
    <w:rPr><w:b/><w:rFonts w:ascii="Aptos" w:eastAsia="Microsoft YaHei" w:hAnsi="Aptos"/><w:color w:val="B45309"/><w:sz w:val="21"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ExamBoxText">
    <w:name w:val="Exam Box Text"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:after="80"/><w:shd w:val="clear" w:color="auto" w:fill="FFFBEB"/><w:ind w:left="260"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Aptos" w:eastAsia="Microsoft YaHei" w:hAnsi="Aptos"/><w:color w:val="78350F"/><w:sz w:val="20"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="MistakeBox">
    <w:name w:val="Mistake Box"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="120" w:after="40"/><w:shd w:val="clear" w:color="auto" w:fill="FEF2F2"/><w:pBdr><w:left w:val="single" w:sz="18" w:space="5" w:color="DC2626"/></w:pBdr></w:pPr>
    <w:rPr><w:b/><w:rFonts w:ascii="Aptos" w:eastAsia="Microsoft YaHei" w:hAnsi="Aptos"/><w:color w:val="B91C1C"/><w:sz w:val="21"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="MistakeBoxText">
    <w:name w:val="Mistake Box Text"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:after="80"/><w:shd w:val="clear" w:color="auto" w:fill="FEF2F2"/><w:ind w:left="260"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Aptos" w:eastAsia="Microsoft YaHei" w:hAnsi="Aptos"/><w:color w:val="7F1D1D"/><w:sz w:val="20"/></w:rPr>
  </w:style>
</w:styles>'''


def add_list_section(doc: DocxBuilder, title: str, items: List[Any], empty_text: str = "无") -> None:
    doc.add_heading(title, 2)
    if not items:
        doc.add_paragraph(empty_text)
        return
    for item in items:
        if isinstance(item, (dict, list)):
            doc.add_bullet(json.dumps(item, ensure_ascii=False))
        else:
            doc.add_bullet(item)


def note_field(note: Dict[str, Any], *keys: str) -> Any:
    for key in keys:
        if key in note and note[key] not in (None, ""):
            return note[key]
    return ""


def add_note_text(doc: DocxBuilder, label: str, value: Any, placeholder: str) -> None:
    doc.add_section_label(label)
    doc.add_structured_value(value, placeholder=placeholder)


def format_question_item(item: Any) -> str:
    if isinstance(item, dict):
        question = item.get("question") or item.get("front") or ""
        answer = item.get("answer") or item.get("back") or item.get("answer_hint") or ""
        if question and answer:
            return f"Q: {question}\nA: {answer}"
        return json.dumps(item, ensure_ascii=False)
    return str(item)


def add_formula_notes(doc: DocxBuilder, formulas: Any) -> None:
    doc.add_paragraph("公式说明和例题", style="FormulaBox", bold=True)
    if not formulas:
        doc.add_paragraph("无。", style="FormulaBoxText")
        return
    if isinstance(formulas, str):
        doc.add_formula_display(formulas)
        return
    if not isinstance(formulas, list):
        doc.add_paragraph(json.dumps(formulas, ensure_ascii=False), style="FormulaBoxText")
        return
    for item in formulas:
        if isinstance(item, dict):
            formula = item.get("formula", "")
            meaning = item.get("meaning", "")
            conditions = item.get("conditions", "")
            example = item.get("example", "")
            if formula:
                doc.add_paragraph("公式：", style="FormulaBoxText", bold=True)
                doc.add_formula_display(formula)
            else:
                doc.add_paragraph("公式：未填写", style="FormulaBoxText")
            if meaning:
                doc.add_paragraph(f"含义：{meaning}", style="FormulaBoxText")
            if conditions:
                doc.add_paragraph(f"条件：{conditions}", style="FormulaBoxText")
            if example:
                doc.add_paragraph(f"例题：{example}", style="FormulaBoxText")
        else:
            doc.add_formula_display(item)


def compact_source_items(items: List[Any], limit: int = 8) -> List[Any]:
    if len(items) <= limit:
        return items
    return items[:limit] + [f"... 其余 {len(items) - limit} 条见 extraction.json"]


def add_question_notes(doc: DocxBuilder, questions: Any) -> None:
    doc.add_paragraph("可能考法/自测题", style="ExamBox", bold=True)
    items = as_list(questions)
    if not items:
        doc.add_paragraph("待补写：至少包含 1 道闭卷自测题和 1 道考试风格题。", style="ExamBoxText")
        return
    for item in items:
        doc.add_bullet(format_question_item(item), style="ExamBoxText")


def build_docx(extraction: Dict[str, Any], notes: Dict[int, Dict[str, Any]], output: Path, layout: str = "study") -> None:
    title = f"PPT学习笔记 - {Path(extraction.get('source', 'slides')).stem}"
    doc = DocxBuilder(title)
    doc.add_title(title, "期末复习讲义 · 先理解，再回忆，最后刷错题")
    doc.add_meta(f"来源文件：{extraction.get('source', '')}")
    doc.add_meta(f"生成时间：{extraction.get('generated_at', '')}")
    doc.add_meta(f"页数：{slide_filter_summary(extraction)}")
    doc.add_meta(f"阅读分层：{review_plan_summary(extraction)}")
    doc.add_meta(f"排版模式：{'复习讲义' if layout == 'study' else '审计全量'}")
    if skipped_navigation_slides(extraction) and layout == "study":
        doc.add_meta("标题页、目录页和章节过渡页已默认压缩；需要全量讲义时使用 --content-filter all。")
    doc.add_callout(
        "使用方式",
        [
            "先读 `START_HERE` 和学习路径，明确哪些页必读、哪些页只扫一眼。",
            "必读深讲页看“考点定位-核心结论-深度讲解”；快速扫读页只抓新增结论和考试信号。",
            "遇到公式页，按“公式-变量-条件-例题”四步复述；最后用主动回忆题闭卷检查。",
        ],
        style="InsightBox",
    )

    warnings = extraction.get("warnings") or []
    if warnings:
        add_list_section(doc, "处理提示", warnings)

    doc_slides = extraction.get("slides", []) if layout == "audit" else note_required_slides(extraction)
    for index, slide in enumerate(doc_slides, start=1):
        number = int(slide.get("number", 0) or 0)
        note = notes.get(number, {})
        title_text = note_field(note, "title") or slide.get("title") or f"Slide {number}"
        tier = slide.get("review_tier", "deep")
        if index > 1:
            doc.add_page_break()
        doc.add_heading(f"第 {number} 页：{title_text}", 1)
        if layout != "audit":
            reasons = "、".join(as_list(slide.get("importance_reasons"))) or "正文学习页"
            doc.add_meta(f"阅读层级：{REVIEW_TIER_LABELS.get(tier, tier)}；原因：{reasons}")
            if slide.get("duplicate_of"):
                doc.add_meta(f"重复压缩：本页与第 {slide.get('duplicate_of')} 页相似，避免重复展开。")

        if layout != "audit" and tier == "quick":
            doc.add_callout("速读结论", note_field(note, "key_takeaways") or note_field(note, "what_it_says", "summary"), style="InsightBox", placeholder="待补写：用 2-4 条写出本页新增结论。")
            doc.add_callout("考试信号", note_field(note, "exam_focus"), style="ExamBox", placeholder="待补写：说明本页是否常考、怎么考；不常考就写“了解即可”。")
            add_note_text(doc, "一句话解释", clamp_text(note_field(note, "detailed_explanation", "complex_explanation") or note_field(note, "what_it_says", "summary"), 260), "待补写：只解释核心概念，不要展开成长文。")
            if slide_has_formula(slide) or note_field(note, "formula_explanations", "formulas"):
                add_formula_notes(doc, note_field(note, "formula_explanations", "formulas"))
            if slide_has_visual(slide):
                add_note_text(doc, "图表/图片只看什么", clamp_text(note_field(note, "visual_explanation", "image_explanation"), 220), "待补写：指出图表结论或需要核对的部分。")
            doc.add_section_label("原文核对")
            for item in compact_source_items(slide.get("text") or [], limit=4):
                doc.add_paragraph(item, style="SourceText")
            continue

        doc.add_callout("考点定位", note_field(note, "exam_focus"), style="ExamBox", placeholder="待补写：说明本页在期末考试中怎么考。")
        doc.add_callout("必须掌握", note_field(note, "key_takeaways"), style="InsightBox", placeholder="待补写：列出本页真正需要记住的 2-4 个点。")
        add_note_text(doc, "这一页讲什么", note_field(note, "what_it_says", "summary"), "待补写：用本页具体术语解释内容，不能泛泛而谈。")
        add_note_text(doc, "深度讲解", note_field(note, "detailed_explanation", "complex_explanation"), "待补写：写清推理链、算法步骤、公式来源或图表含义。")
        add_formula_notes(doc, note_field(note, "formula_explanations", "formulas"))
        doc.add_callout("例题/套用", note_field(note, "worked_examples", "examples"), style="FormulaBox", placeholder="待补写：至少给一个能算、能判断或能复述的例子。")
        add_question_notes(doc, note_field(note, "likely_questions"))
        doc.add_callout("常见错误", note_field(note, "common_mistakes"), style="MistakeBox", placeholder="待补写：列出容易混淆、漏条件、算错的地方。")
        doc.add_callout("记忆钩子", note_field(note, "memory_hooks"), style="InsightBox", placeholder="无。")

        screenshot = slide.get("screenshot", "")
        doc.add_heading("页面截图与核对", 2)
        if not doc.add_image(screenshot, max_width_inches=6.7):
            doc.add_paragraph("未生成页面截图。")

        formulas = slide.get("formulas") or []
        candidates = slide.get("formula_candidates") or []
        if formulas or candidates:
            doc.add_paragraph("从 PPT 结构化文本识别到的公式/候选表达式", style="FormulaBox", bold=True)
            for formula in formulas:
                formula_text = formula.get("text") or "[OMML formula]"
                doc.add_formula_display(formula_text)
            for candidate in candidates:
                doc.add_formula_display(candidate)

        add_note_text(doc, "图片/图表说明", note_field(note, "visual_explanation", "image_explanation"), "待补写：说明图片、图表、流程图、架构图等视觉元素。")
        if layout == "audit":
            add_list_section(doc, "页面文字", slide.get("text") or [])
            if slide.get("notes"):
                add_list_section(doc, "演讲者备注", slide.get("notes") or [])
            if slide.get("tables"):
                add_list_section(doc, "表格内容", slide.get("tables") or [])
            if slide.get("alt_texts"):
                add_list_section(doc, "对象/图片替代文本", slide.get("alt_texts") or [])
            if slide.get("related_objects"):
                add_list_section(doc, "图表/嵌入对象数据", slide.get("related_objects") or [])

            doc.add_heading("图片素材", 2)
            images = slide.get("images") or []
            if not images:
                doc.add_paragraph("未提取到独立图片素材；如本页有图片，请以页面截图为准。")
            for image in images:
                path = image.get("path", "")
                label = image.get("filename") or image.get("target") or "图片"
                doc.add_paragraph(label, bold=True)
                if path and not doc.add_image(path, max_width_inches=3.2):
                    doc.add_paragraph(path)
                elif not path:
                    doc.add_paragraph(json.dumps(image, ensure_ascii=False))
        else:
            doc.add_section_label("原文核对")
            for item in compact_source_items(slide.get("text") or [], limit=6):
                doc.add_paragraph(item, style="SourceText")

        add_note_text(doc, "不确定内容", note_field(note, "uncertainties"), "无。")

    doc.write(output)


def write_markdown(extraction: Dict[str, Any], output: Path) -> None:
    lines = [
        f"# PPT学习笔记提取稿 - {Path(extraction.get('source', 'slides')).stem}",
        "",
        f"- 来源文件: {extraction.get('source', '')}",
        f"- 页数: {extraction.get('slide_count', 0)}",
        "",
    ]
    warnings = extraction.get("warnings") or []
    if warnings:
        lines += ["## 处理提示", ""]
        lines += [f"- {w}" for w in warnings]
        lines.append("")
    for slide in extraction.get("slides", []):
        lines += [f"## 第 {slide.get('number')} 页：{slide.get('title', '')}", ""]
        if slide.get("screenshot"):
            lines += [f"![slide]({slide.get('screenshot')})", ""]
        for section, key in [
            ("页面文字", "text"),
            ("演讲者备注", "notes"),
            ("公式候选", "formula_candidates"),
            ("图表/嵌入对象", "related_objects"),
            ("图片素材", "images"),
        ]:
            lines += [f"### {section}", ""]
            items = slide.get(key) or []
            if items:
                for item in items:
                    lines.append(f"- {json.dumps(item, ensure_ascii=False) if isinstance(item, (dict, list)) else item}")
            else:
                lines.append("- 无")
            lines.append("")
    output.write_text("\n".join(lines), encoding="utf-8")


def md_value(value: Any) -> str:
    if value is None or value == "":
        return "待补写"
    if isinstance(value, str):
        return value
    if isinstance(value, list):
        if not value:
            return "无"
        return "\n".join(f"- {md_value(item)}" for item in value)
    if isinstance(value, dict):
        return json.dumps(value, ensure_ascii=False, indent=2)
    return str(value)


def markdown_note_field(note: Dict[str, Any], placeholder: str, *keys: str) -> str:
    value = note_field(note, *keys)
    if value in ("", None, []):
        return placeholder
    return md_value(value)


def build_notes_markdown(extraction: Dict[str, Any], notes: Dict[int, Dict[str, Any]]) -> str:
    title = f"PPT学习笔记 - {Path(extraction.get('source', 'slides')).stem}"
    lines = [
        f"# {title}",
        "",
        f"- 来源文件: {extraction.get('source', '')}",
        f"- {slide_filter_summary(extraction)}",
        f"- {review_plan_summary(extraction)}",
        f"- 工具版本: {extraction.get('version', VERSION)}",
        "",
    ]
    warnings = extraction.get("warnings") or []
    if warnings:
        lines += ["## 处理提示", ""]
        lines += [f"- {warning}" for warning in warnings]
        lines.append("")

    skipped = skipped_navigation_slides(extraction)
    if skipped:
        lines += ["## 已压缩的导航页", ""]
        for item in skipped:
            lines.append(f"- S{item.get('number')}: {item.get('content_kind')} - {item.get('title', '')}")
        lines.append("")

    references = reference_review_slides(extraction)
    if references:
        lines += ["## 参考/重复页（默认不展开）", ""]
        lines.append("这些页通常是过渡、重复铺垫或低增量内容；复习时按学习路径需要再回看截图。")
        lines.append("")
        for slide in references:
            duplicate = f"，重复 S{slide.get('duplicate_of')}" if slide.get("duplicate_of") else ""
            reasons = "、".join(as_list(slide.get("importance_reasons"))) or "低增量"
            lines.append(f"- S{slide.get('number')}: {slide.get('title', '')} ({reasons}{duplicate})")
        lines.append("")

    for slide in note_required_slides(extraction):
        number = int(slide.get("number", 0) or 0)
        note = notes.get(number, {})
        slide_title = note_field(note, "title") or slide.get("title") or f"Slide {number}"
        tier = slide.get("review_tier", "deep")
        lines += [f"## 第 {number} 页：{slide_title}", ""]
        lines += [
            f"- 阅读层级: {REVIEW_TIER_LABELS.get(tier, tier)}",
            f"- 压缩原因: {'、'.join(as_list(slide.get('importance_reasons'))) or '正文学习页'}",
            "",
        ]
        if slide.get("screenshot"):
            lines += [f"![第 {number} 页截图]({slide.get('screenshot')})", ""]

        if tier == "quick":
            lines += [
                "### 速读结论",
                "",
                markdown_note_field(note, "待补写：用 2-4 条写出本页新增结论。", "key_takeaways", "what_it_says", "summary"),
                "",
                "### 考试信号",
                "",
                markdown_note_field(note, "待补写：说明本页怎么考；不常考就写“了解即可”。", "exam_focus"),
                "",
                "### 一句话解释",
                "",
                clamp_text(note_field(note, "detailed_explanation", "complex_explanation") or note_field(note, "what_it_says", "summary"), 320) or "待补写：只解释核心概念，不要展开成长文。",
                "",
                "### 常见错误",
                "",
                markdown_note_field(note, "无或待补写。", "common_mistakes"),
                "",
            ]
            continue

        lines += [
            "### 这一页是干什么用的",
            "",
            markdown_note_field(note, "待补写：说明本页在整套 PPT 中的作用。", "purpose", "page_purpose"),
            "",
            "### 这一页讲了什么",
            "",
            markdown_note_field(note, "待补写：完整解释本页内容。", "what_it_says", "summary"),
            "",
            "### 复杂内容详解",
            "",
            markdown_note_field(note, "待补写：对复杂概念、推导、图表或算法做展开说明。", "detailed_explanation", "complex_explanation"),
            "",
            "### 图片/图表说明",
            "",
            markdown_note_field(note, "待补写：说明图片、图表、流程图、架构图等视觉元素。", "visual_explanation", "image_explanation"),
            "",
            "### 公式说明和例题",
            "",
        ]
        formulas = note_field(note, "formula_explanations", "formulas")
        if formulas:
            if isinstance(formulas, list):
                for item in formulas:
                    if isinstance(item, dict):
                        lines.append(f"- 公式: {item.get('formula', '未填写')}")
                        if item.get("meaning"):
                            lines.append(f"  - 含义: {item.get('meaning')}")
                        if item.get("conditions"):
                            lines.append(f"  - 条件: {item.get('conditions')}")
                        if item.get("example"):
                            lines.append(f"  - 例题: {item.get('example')}")
                    else:
                        lines.append(f"- {item}")
            else:
                lines.append(md_value(formulas))
        else:
            lines.append("无或待补写。")
        lines += [
            "",
            "### 补充例题/案例",
            "",
            markdown_note_field(note, "无或待补写。", "worked_examples", "examples"),
            "",
            "### 期末考点定位",
            "",
            markdown_note_field(note, "待补写：说明本页在期末考试中可能怎么考。", "exam_focus"),
            "",
            "### 核心记忆点",
            "",
            markdown_note_field(note, "待补写：列出必须背会或能复述的要点。", "key_takeaways"),
            "",
            "### 记忆钩子",
            "",
            markdown_note_field(note, "无或待补写。", "memory_hooks"),
            "",
            "### 可能考法/自测题",
            "",
            markdown_note_field(note, "待补写：生成主动回忆题或考试风格问题。", "likely_questions"),
            "",
            "### 常见错误",
            "",
            markdown_note_field(note, "待补写：说明学生容易错在哪里。", "common_mistakes"),
            "",
            "### 原始提取内容",
            "",
        ]
        for label, key in [
            ("页面文字", "text"),
            ("公式候选", "formula_candidates"),
            ("表格", "tables"),
            ("图片素材", "images"),
            ("图表/嵌入对象", "related_objects"),
        ]:
            items = slide.get(key) or []
            lines.append(f"#### {label}")
            lines.append("")
            if items:
                for item in items:
                    lines.append(f"- {json.dumps(item, ensure_ascii=False) if isinstance(item, (dict, list)) else item}")
            else:
                lines.append("- 无")
            lines.append("")
    return "\n".join(lines)


def write_notes_markdown(extraction: Dict[str, Any], notes: Dict[int, Dict[str, Any]], output: Path) -> None:
    output.write_text(build_notes_markdown(extraction, notes), encoding="utf-8")


def write_prompt_pack(extraction: Dict[str, Any], output: Path, language: str) -> None:
    lines = [
        "# Prompt Pack",
        "",
        "Use this pack to fill `notes_template.json` for the extracted slide deck.",
        "",
        "## Compression Contract",
        "",
        "- Your job is to teach the course with minimum sufficient information, not to rewrite the whole PPT.",
        "- Treat the deck like an information-theory compression problem: preserve high-yield definitions, formulas, examples, traps, and exam signals; remove repeated wording, decorative transitions, and low-increment restatements.",
        "- Deep slides get compact but real explanation. Quick slides get 2-4 bullets plus one exam signal. Reference/repeated slides are listed for orientation only and should not receive full note objects.",
        "- Do not exceed the slide tier. If a slide says `快速扫读`, do not write a long `detailed_explanation`.",
        "- Prefer one concrete example over five paragraphs of abstract advice.",
        "",
        "## Required Output",
        "",
        "- Return valid JSON matching `references/note-schema.md`.",
        "- Keep one object per note-required slide only: `必读深讲` and `快速扫读`.",
        "- Title pages, agenda/table-of-contents pages, and section dividers are compacted by default; do not expand them into full notes unless `--content-filter all` was used.",
        "- For `必读深讲` slides: explain purpose, content, complex ideas, visual elements, formulas, examples, exam focus, and mistakes, but stay dense.",
        "- For `快速扫读` slides: fill purpose, what_it_says, exam_focus, key_takeaways, common_mistakes; keep detailed_explanation short unless the slide has a formula or hard diagram.",
        "- Add final-exam fields: exam_focus, key_takeaways, memory_hooks, likely_questions, common_mistakes, prerequisites, difficulty, estimated_review_minutes, and tags.",
        "- likely_questions should include active-recall questions and at least one exam-style question for important formulas or algorithms.",
        "- Add `practice_questions` when possible: short original or open-source-adapted exercises with answer, solution steps, difficulty, and source/source_url if externally inspired.",
        "- If OCR/math-recognition data is merged, treat it as a signal but mark uncertain formulas as `需核对`.",
        "- For external practice-bank questions, preserve source/source_url and adapt the wording to the current slide rather than copying long passages.",
        "- Avoid generic filler. Do not write vague lines such as 'put this slide back into the chapter logic' unless you name the exact concept, formula, or algorithm.",
        "- detailed_explanation must include a reasoning chain: definition -> condition -> why it works -> how to use it -> where students make mistakes.",
        "- For each important formula, explain units/base/log convention and give a concrete numeric mini-example.",
        "- Mark uncertain visual or formula recognition as `需核对`.",
        f"- Output language: {language}.",
        "",
        "## Source",
        "",
        f"- File: {extraction.get('source', '')}",
        f"- SHA256: {extraction.get('source_sha256', '')}",
        f"- {slide_filter_summary(extraction)}",
        f"- {review_plan_summary(extraction)}",
        "",
    ]
    skipped = skipped_navigation_slides(extraction)
    if skipped:
        lines += ["## Compacted Navigation Slides", ""]
        for item in skipped:
            lines.append(f"- S{item.get('number')}: {item.get('content_kind')} - {item.get('title', '')}")
        lines += ["", "Only use these pages for orientation; do not create full explanatory note objects for them.", ""]

    references = reference_review_slides(extraction)
    if references:
        lines += ["## Reference / Repeated Study Slides", ""]
        lines.append("These study slides were intentionally excluded from note objects to reduce repetition. Use them only to understand flow or to merge a missing detail into a nearby deep/quick slide.")
        lines.append("")
        for slide in references:
            duplicate = f"; duplicate of S{slide.get('duplicate_of')}" if slide.get("duplicate_of") else ""
            reasons = "、".join(as_list(slide.get("importance_reasons"))) or "low incremental value"
            lines.append(f"- S{slide.get('number')}: {slide.get('title', '')} ({reasons}{duplicate})")
        lines.append("")

    for slide in note_required_slides(extraction):
        lines += [
            f"## Slide {slide.get('number')}: {slide.get('title', '')}",
            "",
            f"- Review tier: {slide.get('review_tier_label', REVIEW_TIER_LABELS.get(slide.get('review_tier', 'deep'), '必读深讲'))}",
            f"- Why included: {'、'.join(as_list(slide.get('importance_reasons'))) or 'study content'}",
            f"- Compression guidance: {slide.get('compression_guidance', '')}",
            f"- Screenshot: {slide.get('screenshot') or 'not rendered'}",
            "",
            "### Extracted Text",
            "",
        ]
        text = slide.get("text") or []
        lines += [f"- {item}" for item in text] if text else ["- No structured text extracted."]
        lines += ["", "### Formula Candidates", ""]
        formulas = slide.get("formulas") or []
        candidates = slide.get("formula_candidates") or []
        if formulas:
            for formula in formulas:
                lines.append(f"- {formula.get('text') or '[OMML formula; inspect screenshot]'}")
        if candidates:
            lines += [f"- {candidate}" for candidate in candidates]
        if not formulas and not candidates:
            lines.append("- None extracted; still inspect screenshot for image-only formulas.")
        lines += ["", "### Visual / Object Signals", ""]
        visual_items = {
            "images": slide.get("images") or [],
            "tables": slide.get("tables") or [],
            "related_objects": slide.get("related_objects") or [],
            "alt_texts": slide.get("alt_texts") or [],
        }
        has_visual = False
        for key, values in visual_items.items():
            if not values:
                continue
            has_visual = True
            lines.append(f"#### {key}")
            for value in values:
                lines.append(f"- {json.dumps(value, ensure_ascii=False) if isinstance(value, (dict, list)) else value}")
            lines.append("")
        if not has_visual:
            lines.append("- No separate visual metadata extracted; inspect full-slide screenshot.")
            lines.append("")
    output.write_text("\n".join(lines), encoding="utf-8")


def as_list(value: Any) -> List[Any]:
    if value in (None, ""):
        return []
    if isinstance(value, list):
        return value
    return [value]


def plain(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return re.sub(r"\s+", " ", value).strip()
    if isinstance(value, dict):
        return "; ".join(f"{key}: {plain(item)}" for key, item in value.items() if plain(item))
    if isinstance(value, list):
        return "; ".join(plain(item) for item in value if plain(item))
    return str(value)


def slide_tags(note: Dict[str, Any], slide: Dict[str, Any]) -> str:
    tags = [str(tag).strip().replace(" ", "_") for tag in as_list(note_field(note, "tags")) if str(tag).strip()]
    tags.append(f"slide_{int(slide.get('number', 0) or 0):03d}")
    if slide_has_formula(slide):
        tags.append("formula")
    if slide_has_visual(slide):
        tags.append("visual")
    return " ".join(dict.fromkeys(tags))


def clamp_text(value: Any, limit: int = 120) -> str:
    text = plain(value)
    if len(text) <= limit:
        return text
    return text[: limit - 1].rstrip() + "..."


def normalize_topic_term(term: Any) -> str:
    text = plain(term).strip(" -_:/：,，.。;；|[]()（）{}<>")
    if not text:
        return ""
    text = re.sub(r"\s+", " ", text)
    lower = text.lower()
    if lower in TOPIC_STOPWORDS or text in TOPIC_STOPWORDS:
        return ""
    if lower.startswith("slide_"):
        return ""
    if text.isdigit() or len(text) < 2:
        return ""
    if re.fullmatch(r"[A-Za-z]{2}", text) and not text.isupper():
        return ""
    if len(text) > 24:
        return ""
    return text


def topic_terms_from_text(value: Any, limit: int = 8) -> List[str]:
    text = plain(value)
    if not text:
        return []
    terms: List[str] = []
    for raw in re.findall(r"[A-Za-z][A-Za-z0-9_+\-/]{1,23}", text):
        term = normalize_topic_term(raw)
        if term:
            terms.append(term)
    chunks = re.split(r"[\s,，。；;:：、/|()[\]{}<>《》]+", text)
    for chunk in chunks:
        if not re.fullmatch(r"[\u4e00-\u9fffA-Za-z0-9_-]{2,24}", chunk or ""):
            continue
        term = normalize_topic_term(chunk)
        if term and not re.fullmatch(r"[A-Za-z0-9_+\-/]+", term):
            terms.append(term)
    return unique_preserve(terms)[:limit]


def extract_topic_terms(note: Dict[str, Any], slide: Dict[str, Any], limit: int = 8) -> List[str]:
    terms: List[str] = []
    for tag in as_list(note_field(note, "tags")):
        term = normalize_topic_term(tag)
        if term:
            terms.append(term)
    for value in [
        note_field(note, "title"),
        slide.get("title", ""),
        note_field(note, "exam_focus"),
        note_field(note, "key_takeaways"),
        note_field(note, "prerequisites"),
    ]:
        terms.extend(topic_terms_from_text(value, limit=4))
    for value in (slide.get("formula_candidates") or [])[:3]:
        terms.extend(topic_terms_from_text(value, limit=2))
    for value in (slide.get("text") or [])[:3]:
        terms.extend(topic_terms_from_text(value, limit=3))
    return unique_preserve(terms)[:limit]


def slide_has_extracted_visual(slide: Dict[str, Any]) -> bool:
    return bool(slide.get("images") or slide.get("related_objects") or slide.get("tables") or slide.get("alt_texts"))


def infer_slide_role(note: Dict[str, Any], slide: Dict[str, Any]) -> str:
    text = plain(
        [
            note_field(note, "title"),
            slide.get("title", ""),
            note_field(note, "purpose"),
            note_field(note, "what_it_says", "summary"),
            note_field(note, "exam_focus"),
            slide.get("text") or [],
        ]
    )
    lower = text.lower()
    if any(word in lower for word in ("练习", "习题", "作业", "quiz", "exercise", "homework")):
        return "practice"
    if any(word in lower for word in ("例题", "案例", "example", "case study")):
        return "example"
    if any(word in lower for word in ("总结", "回顾", "小结", "summary", "review")):
        return "summary"
    if any(word in lower for word in ("algorithm", "算法", "步骤", "流程", "procedure", "迭代", "递归", "贪心")):
        return "algorithm"
    if slide_has_formula(slide) or note_field(note, "formula_explanations", "formulas"):
        return "formula"
    if slide_has_extracted_visual(slide):
        return "visual"
    return "concept"


def parse_review_minutes(note: Dict[str, Any], slide: Dict[str, Any]) -> int:
    raw = plain(note_field(note, "estimated_review_minutes"))
    match = re.search(r"\d+", raw)
    if match:
        return max(3, min(int(match.group(0)), 45))
    difficulty = plain(note_field(note, "difficulty"))
    if "困难" in difficulty or difficulty.lower() in {"hard", "difficult"}:
        return 12
    if slide_has_formula(slide):
        return 10
    if slide_has_extracted_visual(slide):
        return 8
    return 6


def slide_difficulty_label(note: Dict[str, Any], slide: Dict[str, Any]) -> str:
    difficulty = plain(note_field(note, "difficulty"))
    if difficulty:
        return difficulty
    if slide_has_formula(slide):
        return "困难"
    if slide_has_extracted_visual(slide):
        return "中等"
    return "基础"


def is_section_break_title(title: str) -> bool:
    text = plain(title)
    if not text:
        return False
    lower = text.lower()
    if any(word in lower for word in ("chapter", "section", "module", "unit", "part ")):
        return True
    return bool(re.search(r"(第\s*[一二三四五六七八九十0-9]+\s*[章节部分讲]|^目录$|^大纲$|^outline$)", text))


def term_overlap_score(left: List[str], right: List[str]) -> float:
    left_set = {item.lower() for item in left if item}
    right_set = {item.lower() for item in right if item}
    if not left_set or not right_set:
        return 0.0
    return len(left_set & right_set) / max(len(left_set | right_set), 1)


def slide_span(numbers: List[int]) -> str:
    if not numbers:
        return "无"
    spans: List[str] = []
    start = prev = numbers[0]
    for number in numbers[1:]:
        if number == prev + 1:
            prev = number
            continue
        spans.append(f"S{start}" if start == prev else f"S{start}-S{prev}")
        start = prev = number
    spans.append(f"S{start}" if start == prev else f"S{start}-S{prev}")
    return ", ".join(spans)


def choose_module_title(items: List[Dict[str, Any]], module_index: int) -> str:
    term_counts: Dict[str, int] = {}
    for item in items:
        for term in item["terms"][:5]:
            term_counts[term] = term_counts.get(term, 0) + 1
    ranked_terms = sorted(term_counts.items(), key=lambda pair: (-pair[1], len(pair[0]), pair[0]))
    if ranked_terms:
        return "、".join(term for term, _ in ranked_terms[:2])
    first_title = plain(items[0].get("title"))
    if first_title:
        return clamp_text(first_title, 28)
    return f"模块 {module_index}"


def collect_module_items(items: List[Dict[str, Any]], field: str, limit: int = 3) -> List[str]:
    collected: List[str] = []
    for item in items:
        for value in as_list(note_field(item["note"], field)):
            text = clamp_text(value, 140)
            if text:
                collected.append(f"S{item['number']}: {text}")
    return unique_preserve(collected)[:limit]


def collect_module_questions(items: List[Dict[str, Any]], limit: int = 3) -> List[str]:
    questions: List[str] = []
    for item in items:
        for question in as_list(note_field(item["note"], "likely_questions")):
            if isinstance(question, dict):
                text = question.get("question") or question.get("front") or ""
            else:
                text = question
            text = clamp_text(text, 120)
            if text:
                questions.append(f"S{item['number']}: {text}")
    return unique_preserve(questions)[:limit]


def focus_reason(item: Dict[str, Any]) -> str:
    reasons: List[str] = []
    tier = item["slide"].get("review_tier")
    if tier in REVIEW_TIER_LABELS:
        reasons.append(REVIEW_TIER_LABELS[tier])
    if slide_has_formula(item["slide"]) or note_field(item["note"], "formula_explanations", "formulas"):
        reasons.append("公式/推导")
    if slide_has_extracted_visual(item["slide"]):
        reasons.append("图表/素材")
    difficulty = slide_difficulty_label(item["note"], item["slide"])
    if "困难" in difficulty or difficulty.lower() in {"hard", "difficult"}:
        reasons.append("高难度")
    if note_field(item["note"], "exam_focus"):
        reasons.append("有明确考点")
    return "、".join(reasons[:3]) or ROLE_LABELS.get(item["role"], item["role"])


def summarize_learning_module(items: List[Dict[str, Any]], module_index: int) -> Dict[str, Any]:
    numbers = [item["number"] for item in items]
    roles = unique_preserve([ROLE_LABELS.get(item["role"], item["role"]) for item in items])
    all_terms = unique_preserve([term for item in items for term in item["terms"]])
    focus_items = [
        item
        for item in items
        if item["slide"].get("review_tier") in {"deep", "quick"}
        or slide_has_formula(item["slide"])
        or slide_has_extracted_visual(item["slide"])
        or "困难" in slide_difficulty_label(item["note"], item["slide"])
        or note_field(item["note"], "exam_focus")
    ]
    if not focus_items and items:
        focus_items = [items[0]]
    title = choose_module_title(items, module_index)
    topic = "、".join(all_terms[:4]) or title
    goal = f"理解 {topic}，能说明这些页面如何从{roles[0] if roles else '概念'}推进到考点应用。"
    exam_focuses = collect_module_items(items, "exam_focus", limit=2)
    if exam_focuses:
        goal = clamp_text(exam_focuses[0], 160)
    checkpoints = collect_module_questions(items, limit=3)
    if not checkpoints:
        checkpoints = [f"闭卷说出 {topic} 的定义、适用条件和一个典型考法。"]
    traps = collect_module_items(items, "common_mistakes", limit=3)
    if not traps:
        traps = [f"不要只记 {topic} 的结论；要能说出条件、变量或图表读法。"]
    prerequisites = collect_module_items(items, "prerequisites", limit=3)
    return {
        "index": module_index,
        "title": title,
        "slides": numbers,
        "roles": roles,
        "terms": all_terms[:6],
        "minutes": sum(parse_review_minutes(item["note"], item["slide"]) for item in items),
        "goal": goal,
        "focus": [(item["number"], focus_reason(item)) for item in focus_items[:4]],
        "checkpoints": checkpoints,
        "traps": traps,
        "prerequisites": prerequisites,
    }


def build_learning_modules(extraction: Dict[str, Any], notes: Dict[int, Dict[str, Any]]) -> List[Dict[str, Any]]:
    prepared: List[Dict[str, Any]] = []
    for slide in study_slides(extraction):
        number = int(slide.get("number", 0) or 0)
        note = notes.get(number, {})
        prepared.append(
            {
                "number": number,
                "title": note_field(note, "title") or slide.get("title") or f"Slide {number}",
                "terms": extract_topic_terms(note, slide),
                "role": infer_slide_role(note, slide),
                "note": note,
                "slide": slide,
            }
        )
    modules: List[Dict[str, Any]] = []
    current: List[Dict[str, Any]] = []
    current_terms: List[str] = []
    for item in prepared:
        start_new = False
        if current:
            overlap = term_overlap_score(current_terms, item["terms"])
            if is_section_break_title(item["title"]) and len(current) >= 2:
                start_new = True
            elif len(current) >= 7:
                start_new = True
            elif len(current) >= 4 and item["role"] in {"concept", "formula"} and overlap < 0.08:
                start_new = True
        if start_new:
            modules.append(summarize_learning_module(current, len(modules) + 1))
            current = []
            current_terms = []
        current.append(item)
        current_terms = unique_preserve(current_terms + item["terms"])
    if current:
        modules.append(summarize_learning_module(current, len(modules) + 1))
    return modules


def write_learning_path(extraction: Dict[str, Any], notes: Dict[int, Dict[str, Any]], output: Path) -> None:
    modules = build_learning_modules(extraction, notes)
    total_minutes = sum(int(module.get("minutes", 0) or 0) for module in modules)
    total_time = f"{total_minutes} 分钟" if total_minutes else "未估算"
    lines = [
        "# 学习路径",
        "",
        "这份文件负责回答“先学什么、为什么、学到什么程度再往后走”。它不重复完整讲义，只给复习顺序和通关标准。",
        "",
        "## 先学路径",
        "",
    ]
    if not modules:
        lines.append("- 未检测到页面。")
    for module in modules:
        terms = "、".join(module["terms"][:4]) or module["title"]
        lines.append(f"{module['index']}. **{module['title']}** ({slide_span(module['slides'])}, 约 {module['minutes']} 分钟): {terms}")
    lines += [
        "",
        f"- 模块数: {len(modules)}",
        f"- 预计首轮理解时间: {total_time}",
        "- 用法: 先按下面每个模块过一遍，再打开完整讲义补细节，最后做主动回忆题。",
        "",
    ]
    for module in modules:
        focus = "；".join(f"S{number}({reason})" for number, reason in module["focus"]) or slide_span(module["slides"])
        role_text = " -> ".join(module["roles"]) or "概念"
        lines += [
            f"## {module['index']}. {module['title']}",
            "",
            f"- **覆盖页面**: {slide_span(module['slides'])}",
            f"- **课堂角色**: {role_text}",
            f"- **学习目标**: {module['goal']}",
        ]
        if module["prerequisites"]:
            lines.append(f"- **先补前置**: {'；'.join(module['prerequisites'])}")
        lines += [
            f"- **必看页**: {focus}",
            "- **学习动作**:",
            "  1. 先看本模块标题和必看页截图，判断它在讲定义、公式、方法还是例题。",
            "  2. 再读 `01_复习讲义.docx` 对应页面，只摘条件、变量、图表结论和常见错误。",
            "  3. 合上讲义，用下面的通关题闭卷回答。",
            "- **自测通过线**:",
        ]
        for checkpoint in module["checkpoints"]:
            lines.append(f"  - {checkpoint}")
        lines.append("- **常见卡点**:")
        for trap in module["traps"]:
            lines.append(f"  - {trap}")
        lines += [
            "- **进入下一模块前**: 能用自己的话讲清本模块解决什么问题、哪些条件不能漏、遇到题目先看哪个信号。",
            "",
        ]
    lines += [
        "## 复习分流",
        "",
        "- 时间少: 只看每个模块的必看页，再做主动回忆题。",
        "- 公式多: 先走本路径，再集中打开 `05_公式速查.md` 手算例题。",
        "- 图表多: 每个图都要说出轴、箭头、区域或表头分别代表什么。",
        "- 考前最后一遍: 只看本路径的自测通过线、`03_一页纸总览.md` 和错题本。",
        "",
    ]
    output.write_text("\n".join(lines), encoding="utf-8")


def write_review_index(extraction: Dict[str, Any], output: Path) -> None:
    lines = [
        "# 阅读取舍索引",
        "",
        "这份文件先回答“哪些必须读、哪些扫一眼、哪些暂时不用看”。默认按信息增量压缩 PPT，避免把重复页写成长文。",
        "",
        f"- {slide_filter_summary(extraction)}",
        f"- {review_plan_summary(extraction)}",
        "",
        "## 1. 必读深讲",
        "",
        "这些页通常含公式、定义、例题、图表或期末高频考法，应该认真读讲义并闭卷复述。",
        "",
    ]
    deep = deep_review_slides(extraction)
    if not deep:
        lines.append("- 暂无。")
    for slide in deep:
        reasons = "、".join(as_list(slide.get("importance_reasons"))) or "高价值学习页"
        lines.append(f"- S{slide.get('number')}: {slide.get('title', '')} - {reasons}")
    lines += [
        "",
        "## 2. 快速扫读",
        "",
        "这些页只抓新增结论、考试信号和容易错的条件；不要逐字重写 PPT。",
        "",
    ]
    quick = quick_review_slides(extraction)
    if not quick:
        lines.append("- 暂无。")
    for slide in quick:
        reasons = "、".join(as_list(slide.get("importance_reasons"))) or "中等信息量"
        duplicate = f"，与 S{slide.get('duplicate_of')} 相似" if slide.get("duplicate_of") else ""
        lines.append(f"- S{slide.get('number')}: {slide.get('title', '')} - {reasons}{duplicate}")
    lines += [
        "",
        "## 3. 参考/重复",
        "",
        "这些页默认不写逐页讲解；只有当你在必读页看不懂上下文时再回看。",
        "",
    ]
    reference = reference_review_slides(extraction)
    if not reference:
        lines.append("- 暂无。")
    for slide in reference:
        reasons = "、".join(as_list(slide.get("importance_reasons"))) or "低增量"
        duplicate = f"，与 S{slide.get('duplicate_of')} 相似" if slide.get("duplicate_of") else ""
        lines.append(f"- S{slide.get('number')}: {slide.get('title', '')} - {reasons}{duplicate}")
    lines += [
        "",
        "## 使用规则",
        "",
        "1. 第一遍只读必读深讲页。",
        "2. 第二遍扫快速页，发现不会的再回到必读页。",
        "3. 参考/重复页不主动读，除非题目或公式要求核对原图。",
        "4. 做题错了，再用错题反馈路径把参考页拉回来。",
        "",
    ]
    output.write_text("\n".join(lines), encoding="utf-8")


def csv_rows_to_text(rows: List[List[str]]) -> str:
    buffer = io.StringIO()
    writer = csv.writer(buffer)
    writer.writerows(rows)
    return buffer.getvalue()


def add_flashcard(rows: List[List[str]], front: str, back: str, tags: str, slide_number: int) -> None:
    front = plain(front)
    back = plain(back)
    if not front or not back:
        return
    rows.append([front, back, tags, str(slide_number)])


def study_days_until(exam_date: Optional[str]) -> Optional[int]:
    if not exam_date:
        return None
    try:
        target = Date.fromisoformat(exam_date)
    except ValueError:
        return None
    today = Date.today()
    return max((target - today).days, 0)


def write_formula_sheet(extraction: Dict[str, Any], notes: Dict[int, Dict[str, Any]], output: Path) -> None:
    lines = ["# Formula Sheet", ""]
    has_formula = False
    for slide in study_slides(extraction):
        number = int(slide.get("number", 0) or 0)
        note = notes.get(number, {})
        formulas = as_list(note_field(note, "formula_explanations", "formulas"))
        candidates = slide.get("formula_candidates") or []
        omml = slide.get("formulas") or []
        if not formulas and not candidates and not omml:
            continue
        has_formula = True
        lines += [f"## Slide {number}: {note_field(note, 'title') or slide.get('title', '')}", ""]
        for item in formulas:
            if isinstance(item, dict):
                formula = item.get("formula", "未填写")
                lines.append(f"- Formula: `{formula}`")
                display = latex_to_readable_formula(formula)
                if display and display != formula:
                    lines.append(f"  - Display: `{display}`")
                if item.get("meaning"):
                    lines.append(f"  - Meaning: {item.get('meaning')}")
                if item.get("conditions"):
                    lines.append(f"  - Conditions: {item.get('conditions')}")
                if item.get("example"):
                    lines.append(f"  - Example: {item.get('example')}")
            else:
                lines.append(f"- {item}")
        for item in omml:
            lines.append(f"- Extracted Office Math: `{item.get('text') or '[inspect screenshot]'}`")
        for candidate in candidates:
            lines.append(f"- Candidate: `{candidate}`")
        lines.append("")
    if not has_formula:
        lines.append("No formulas were extracted. Still inspect slide screenshots for image-only formulas.")
    output.write_text("\n".join(lines), encoding="utf-8")


def write_active_recall(extraction: Dict[str, Any], notes: Dict[int, Dict[str, Any]], md_path: Path, json_path: Path) -> None:
    items: List[Dict[str, Any]] = []
    lines = ["# Active Recall Questions", ""]
    for slide in note_required_slides(extraction):
        number = int(slide.get("number", 0) or 0)
        note = notes.get(number, {})
        title = note_field(note, "title") or slide.get("title", "")
        questions = as_list(note_field(note, "likely_questions"))
        if not questions:
            questions = [
                f"第 {number} 页的核心概念是什么？",
                f"这页内容在期末考试中可能怎么考？",
            ]
            if slide_has_formula(slide):
                questions.append("本页公式中每个变量分别代表什么？请举一个使用例子。")
            if slide_has_visual(slide):
                questions.append("本页图表/图片想表达的结论是什么？")
        lines += [f"## Slide {number}: {title}", ""]
        for index, question in enumerate(questions, start=1):
            answer = ""
            if isinstance(question, dict):
                answer = plain(question.get("answer") or question.get("back") or "")
                question_text = plain(question.get("question") or question.get("front") or question)
            else:
                question_text = plain(question)
            if not answer:
                answer = plain(note_field(note, "what_it_says", "summary")) or "查看本页详细笔记后闭卷复述。"
            items.append({"slide": number, "title": title, "question": question_text, "answer_hint": answer})
            lines.append(f"{index}. {question_text}")
            lines.append(f"   - Answer hint: {answer}")
        lines.append("")
    json_path.write_text(json.dumps({"questions": items}, ensure_ascii=False, indent=2), encoding="utf-8")
    md_path.write_text("\n".join(lines), encoding="utf-8")


def practice_question_from_dict(raw: Dict[str, Any], slide_number: int, title: str) -> Dict[str, Any]:
    question = plain(raw.get("question") or raw.get("prompt") or raw.get("front") or "")
    answer = plain(raw.get("answer") or raw.get("back") or raw.get("answer_hint") or "")
    solution = plain(raw.get("solution") or raw.get("explanation") or raw.get("解题思路") or "")
    source = plain(raw.get("source") or raw.get("source_title") or "")
    source_url = plain(raw.get("source_url") or raw.get("url") or "")
    difficulty = plain(raw.get("difficulty") or "基础")
    if not solution and answer:
        solution = "先定位题目考查的定义、条件或公式，再逐步代入；最后用答案检查单位、条件和结论是否一致。"
    return {
        "slide": slide_number,
        "title": title,
        "difficulty": difficulty,
        "question": question,
        "answer": answer or "回看本页讲义后，用自己的话写出完整答案。",
        "solution": solution or "先写出本页核心概念，再说明适用条件，最后给出结论。",
        "source": source or "PPT 内容原创生成",
        "source_url": source_url,
    }


def load_practice_bank(path: Optional[Path]) -> List[Dict[str, Any]]:
    if not path:
        return []
    raw = json.loads(path.read_text(encoding="utf-8"))
    items = raw if isinstance(raw, list) else raw.get("questions", [])
    return [item for item in items if isinstance(item, dict) and plain(item.get("question"))]


def practice_bank_matches(bank_item: Dict[str, Any], terms: List[str]) -> bool:
    haystack = plain(
        [
            bank_item.get("terms") or [],
            bank_item.get("tags") or [],
            bank_item.get("topic") or "",
            bank_item.get("question") or "",
        ]
    ).lower()
    return any(term.lower() in haystack for term in terms if len(term) >= 2)


def build_slide_practice_questions(
    slide: Dict[str, Any],
    note: Dict[str, Any],
    max_items: int = 3,
    practice_bank: Optional[List[Dict[str, Any]]] = None,
) -> List[Dict[str, Any]]:
    number = int(slide.get("number", 0) or 0)
    title = note_field(note, "title") or slide.get("title", f"Slide {number}")
    items: List[Dict[str, Any]] = []
    terms = extract_topic_terms(note, slide, limit=10)
    for raw in practice_bank or []:
        if practice_bank_matches(raw, terms):
            item = practice_question_from_dict(raw, number, title)
            item["source"] = item["source"] if item["source"] != "PPT 内容原创生成" else "开放题源改编"
            items.append(item)
        if len(items) >= max_items:
            return items[:max_items]

    for raw in as_list(note_field(note, "practice_questions", "exercises")):
        if isinstance(raw, dict):
            item = practice_question_from_dict(raw, number, title)
        else:
            item = {
                "slide": number,
                "title": title,
                "difficulty": "基础",
                "question": plain(raw),
                "answer": plain(note_field(note, "what_it_says", "summary")) or "回看本页讲义后回答。",
                "solution": "先定位题目中的关键词，再回到本页定义、公式或图表结论，最后写出完整推理。",
                "source": "PPT 内容原创生成",
                "source_url": "",
            }
        if item["question"]:
            items.append(item)
    if len(items) >= max_items:
        return items[:max_items]

    for formula in as_list(note_field(note, "formula_explanations", "formulas")):
        if not isinstance(formula, dict):
            continue
        formula_text = formula.get("formula", "")
        readable = latex_to_readable_formula(formula_text)
        question = f"解释并套用公式 {readable or formula_text}：每个变量代表什么？这个公式适用在什么条件下？"
        if formula.get("example"):
            question += f" 仿照例题完成一次计算或判断：{plain(formula.get('example'))}"
        items.append(
            {
                "slide": number,
                "title": title,
                "difficulty": slide_difficulty_label(note, slide),
                "question": question,
                "answer": plain(formula.get("meaning")) or "写出变量含义、公式目标和结果解释。",
                "solution": plain(formula.get("conditions")) or "先确认公式条件，再代入变量，最后解释结果含义。",
                "source": "PPT 公式原创改编",
                "source_url": "",
            }
        )
        if len(items) >= max_items:
            return items[:max_items]

    for raw_question in as_list(note_field(note, "likely_questions")):
        if isinstance(raw_question, dict):
            question = plain(raw_question.get("question") or raw_question.get("front") or "")
            answer = plain(raw_question.get("answer") or raw_question.get("back") or "")
        else:
            question = plain(raw_question)
            answer = ""
        if not question:
            continue
        items.append(
            {
                "slide": number,
                "title": title,
                "difficulty": slide_difficulty_label(note, slide),
                "question": question,
                "answer": answer or plain(note_field(note, "what_it_says", "summary")) or "回看本页讲义后回答。",
                "solution": "先判断题目考查的是概念、公式、图表还是常见错误；再按本页讲义中的条件和步骤作答。",
                "source": "PPT 考点原创改编",
                "source_url": "",
            }
        )
        if len(items) >= max_items:
            return items[:max_items]

    fallback_question = f"第 {number} 页 `{title}` 的核心考点是什么？请写出一个容易错的地方。"
    items.append(
        {
            "slide": number,
            "title": title,
            "difficulty": slide_difficulty_label(note, slide),
            "question": fallback_question,
            "answer": plain(note_field(note, "key_takeaways")) or plain(note_field(note, "exam_focus")) or "写出本页核心结论。",
            "solution": "先用一句话说出本页解决的问题，再列出关键词、条件和常见错误。",
            "source": "PPT 内容原创生成",
            "source_url": "",
        }
    )
    return items[:max_items]


def write_practice_questions(
    extraction: Dict[str, Any],
    notes: Dict[int, Dict[str, Any]],
    md_path: Path,
    json_path: Path,
    practice_bank: Optional[List[Dict[str, Any]]] = None,
) -> None:
    items: List[Dict[str, Any]] = []
    for slide in note_required_slides(extraction):
        number = int(slide.get("number", 0) or 0)
        max_items = 2 if slide.get("review_tier") == "deep" else 1
        items.extend(build_slide_practice_questions(slide, notes.get(number, {}), max_items=max_items, practice_bank=practice_bank))
    lines = [
        "# 小题练习",
        "",
        "这些题目用于把 PPT 内容变成可做的基础练习。默认题目为根据课件原创生成或改编；如果人工/联网补充题源，必须保留来源链接并避免整段复制题库内容。",
        "",
    ]
    for index, item in enumerate(items, start=1):
        lines += [
            f"## {index}. Slide {item['slide']}: {item['title']}",
            "",
            f"- 难度: {item['difficulty']}",
            f"- 来源: {item['source']}{' - ' + item['source_url'] if item.get('source_url') else ''}",
            "",
            f"**题目:** {item['question']}",
            "",
            f"**答案:** {item['answer']}",
            "",
            f"**解题思路:** {item['solution']}",
            "",
        ]
    json_path.write_text(json.dumps({"practice_questions": items}, ensure_ascii=False, indent=2), encoding="utf-8")
    md_path.write_text("\n".join(lines), encoding="utf-8")


def write_flashcards(extraction: Dict[str, Any], notes: Dict[int, Dict[str, Any]], csv_path: Path, md_path: Path) -> None:
    rows = [["Front", "Back", "Tags", "SourceSlide"]]
    md_lines = ["# Flashcards", ""]
    for slide in note_required_slides(extraction):
        number = int(slide.get("number", 0) or 0)
        note = notes.get(number, {})
        title = note_field(note, "title") or slide.get("title", f"Slide {number}")
        tags = slide_tags(note, slide)
        add_flashcard(rows, f"第 {number} 页的作用是什么？", note_field(note, "purpose", "page_purpose"), tags, number)
        add_flashcard(rows, f"{title}: 必须记住的核心点有哪些？", note_field(note, "key_takeaways") or note_field(note, "what_it_says", "summary"), tags, number)
        add_flashcard(rows, f"{title}: 常见错误是什么？", note_field(note, "common_mistakes"), tags, number)
        for formula in as_list(note_field(note, "formula_explanations", "formulas")):
            if isinstance(formula, dict):
                front = f"解释公式：{formula.get('formula', '')}"
                back = f"{formula.get('meaning', '')} {formula.get('conditions', '')} {formula.get('example', '')}"
                add_flashcard(rows, front, back, tags + " formula", number)
        for question in as_list(note_field(note, "likely_questions")):
            if isinstance(question, dict):
                add_flashcard(rows, question.get("question") or question.get("front") or "", question.get("answer") or question.get("back") or note_field(note, "what_it_says"), tags + " recall", number)
            else:
                add_flashcard(rows, question, note_field(note, "what_it_says", "summary"), tags + " recall", number)

    csv_path.write_text(csv_rows_to_text(rows), encoding="utf-8")
    for front, back, tags, slide_number in rows[1:]:
        md_lines += [f"## Slide {slide_number}", "", f"**Q:** {front}", "", f"**A:** {back}", "", f"`{tags}`", ""]
    md_path.write_text("\n".join(md_lines), encoding="utf-8")


def write_cram_plan(
    extraction: Dict[str, Any],
    notes: Dict[int, Dict[str, Any]],
    output: Path,
    exam_date: Optional[str],
    daily_minutes: int,
    target_score: Optional[str],
) -> None:
    days = study_days_until(exam_date)
    slide_count = len(study_slides(extraction))
    minutes = max(daily_minutes, 20)
    target = target_score or "按课程目标自定"
    lines = [
        "# Final Exam Cram Plan",
        "",
        f"- Exam date: {exam_date or '未设置'}",
        f"- Days left: {days if days is not None else '未设置'}",
        f"- Daily minutes: {minutes}",
        f"- Target score: {target}",
        f"- Study slides: {slide_count}",
        f"- {slide_filter_summary(extraction)}",
        "",
        "## Daily Loop",
        "",
        "1. 先看 `learning_path.md`，按模块确定今天学哪几页。",
        "2. 用 `active_recall_questions.md` 闭卷回答，答不出就回看对应页。",
        "3. 做 `practice_questions.md`，先自己写，再看答案和解题思路。",
        "4. 复习 `formula_sheet.md`，每个公式至少手算一个例子。",
        "5. 用 `mistake_log_template.md` 记录错因，而不是只记录答案。",
        "6. 睡前用 `flashcards_anki.csv` 或 `flashcards.md` 快速回忆。",
        "",
    ]
    if days is None or days >= 7:
        phases = [
            ("Day 1-2", "通读所有学习页，标记不会的公式、图表和定义。"),
            ("Day 3-4", "按章节做主动回忆，补齐 `common_mistakes` 和错题本。"),
            ("Day 5", "只做公式、算法、图表解释题。"),
            ("Day 6", "模拟考试：闭卷回答所有 high-yield 问题。"),
            ("Day 7 / Exam Eve", "只看一页纸总结、公式纸和错题本。"),
        ]
    elif days >= 3:
        phases = [
            ("Day 1", "快速通读，先掌握每页 `exam_focus` 和 `key_takeaways`。"),
            ("Day 2", "刷主动回忆题和公式例题，所有答不出的加入错题本。"),
            ("Day 3 / Exam Eve", "只看错题本、公式纸、一页纸总结。"),
        ]
    else:
        phases = [
            ("Today", "先背 `one_page_review.md`，再做所有 active recall。"),
            ("Exam Eve", "公式纸和错题本优先，不再大面积重读 PPT。"),
        ]
    lines += ["## Suggested Schedule", ""]
    for phase, action in phases:
        lines.append(f"- **{phase}**: {action}")
    lines += ["", "## Slide Priority", ""]
    for slide in study_slides(extraction):
        number = int(slide.get("number", 0) or 0)
        note = notes.get(number, {})
        difficulty = note_field(note, "difficulty") or ("困难" if slide_has_formula(slide) else "中等" if slide_has_visual(slide) else "基础")
        focus = note_field(note, "exam_focus") or note_field(note, "purpose") or slide.get("title", "")
        est = note_field(note, "estimated_review_minutes") or ("12" if difficulty == "困难" else "8" if difficulty == "中等" else "5")
        tier = REVIEW_TIER_LABELS.get(slide.get("review_tier", "deep"), slide.get("review_tier", "deep"))
        lines.append(f"- Slide {number} [{tier}, {difficulty}, {est} min]: {focus}")
    output.write_text("\n".join(lines), encoding="utf-8")


def write_one_page_review(extraction: Dict[str, Any], notes: Dict[int, Dict[str, Any]], output: Path) -> None:
    lines = ["# One Page Review", "", "## High-Yield Points", ""]
    for slide in note_required_slides(extraction):
        number = int(slide.get("number", 0) or 0)
        note = notes.get(number, {})
        takeaways = as_list(note_field(note, "key_takeaways")) or [note_field(note, "exam_focus") or note_field(note, "what_it_says", "summary") or slide.get("title", "")]
        limit = 3 if slide.get("review_tier") == "deep" else 1
        for item in takeaways[:limit]:
            if plain(item):
                lines.append(f"- S{number}: {plain(item)}")
    lines += ["", "## Must-Check Mistakes", ""]
    has_mistake = False
    for slide in note_required_slides(extraction):
        number = int(slide.get("number", 0) or 0)
        for item in as_list(note_field(notes.get(number, {}), "common_mistakes")):
            if plain(item):
                has_mistake = True
                lines.append(f"- S{number}: {plain(item)}")
    if not has_mistake:
        lines.append("- 待补写：把平时做题中反复错的点放到这里。")
    output.write_text("\n".join(lines), encoding="utf-8")


def mermaid_node_id(number: int) -> str:
    return f"S{number}"


def mermaid_label(text: str) -> str:
    safe = plain(text).replace('"', "'")
    return safe[:60] or "Untitled"


def write_concept_map(extraction: Dict[str, Any], notes: Dict[int, Dict[str, Any]], output: Path) -> None:
    lines = ["flowchart TD"]
    previous = None
    for slide in study_slides(extraction):
        number = int(slide.get("number", 0) or 0)
        note = notes.get(number, {})
        label = mermaid_label(note_field(note, "title") or slide.get("title", f"Slide {number}"))
        node = mermaid_node_id(number)
        lines.append(f'  {node}["S{number}: {label}"]')
        if previous:
            lines.append(f"  {previous} --> {node}")
        if slide_has_formula(slide):
            formula_node = f"{node}_F"
            lines.append(f'  {formula_node}["公式/推导"]')
            lines.append(f"  {node} --> {formula_node}")
        if slide_has_visual(slide):
            visual_node = f"{node}_V"
            lines.append(f'  {visual_node}["图表/视觉理解"]')
            lines.append(f"  {node} --> {visual_node}")
        previous = node
    output.write_text("\n".join(lines) + "\n", encoding="utf-8")


def write_mistake_log(output: Path) -> None:
    output.write_text(
        "\n".join(
            [
                "# Mistake Log",
                "",
                "| Date | Slide | Question | Wrong Answer | Root Cause | Correct Method | Next Review |",
                "| --- | --- | --- | --- | --- | --- | --- |",
                "|  |  |  |  | 概念不清 / 公式不会用 / 审题错误 / 计算错误 / 图表没看懂 |  |  |",
                "",
                "Use this as an error notebook. The goal is to identify repeatable causes, not just copy correct answers.",
            ]
        ),
        encoding="utf-8",
    )


def load_wrong_answers(path: Optional[Path]) -> Dict[str, Any]:
    if not path:
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def write_adaptive_review(
    extraction: Dict[str, Any],
    notes: Dict[int, Dict[str, Any]],
    output: Path,
    template_path: Path,
    wrong_answers: Optional[Dict[str, Any]] = None,
) -> None:
    practice_items: List[Dict[str, Any]] = []
    for slide in note_required_slides(extraction):
        number = int(slide.get("number", 0) or 0)
        practice_items.extend(build_slide_practice_questions(slide, notes.get(number, {}), max_items=1))
    template = {
        "missed": [
            {
                "slide": item.get("slide"),
                "question": item.get("question"),
                "your_answer": "",
                "root_cause": "概念不清 / 公式不会用 / 审题错误 / 计算错误 / 图表没看懂",
                "note": "",
            }
            for item in practice_items[:20]
        ]
    }
    template_path.write_text(json.dumps(template, ensure_ascii=False, indent=2), encoding="utf-8")

    missed = []
    if wrong_answers:
        raw_missed = wrong_answers.get("missed", wrong_answers if isinstance(wrong_answers, list) else [])
        missed = [item for item in raw_missed if isinstance(item, dict)]
    lines = [
        "# 错题反馈路径",
        "",
        "这个文件把错题变成二次学习路线。没有传入错题时，它会给出高风险页面和错题记录模板；传入 `--wrong-answers-json` 后会按错因重新排序复习页。",
        "",
        f"- 错题输入模板: `{template_path.name}`",
        "",
    ]
    if missed:
        slide_counts: Dict[int, int] = {}
        causes: Dict[str, int] = {}
        for item in missed:
            try:
                number = int(item.get("slide") or item.get("page") or 0)
            except Exception:
                number = 0
            if number:
                slide_counts[number] = slide_counts.get(number, 0) + 1
            cause = plain(item.get("root_cause")) or "未分类"
            causes[cause] = causes.get(cause, 0) + 1
        lines += ["## 优先重学", ""]
        for number, count in sorted(slide_counts.items(), key=lambda pair: (-pair[1], pair[0]))[:10]:
            slide = next((s for s in extraction.get("slides", []) if int(s.get("number", 0) or 0) == number), {})
            note = notes.get(number, {})
            focus = plain(note_field(note, "exam_focus")) or plain(note_field(note, "key_takeaways")) or slide.get("title", "")
            lines.append(f"- S{number}: 错 {count} 次。先重看 `{focus}`，再做对应小题。")
        lines += ["", "## 错因统计", ""]
        for cause, count in sorted(causes.items(), key=lambda pair: (-pair[1], pair[0])):
            lines.append(f"- {cause}: {count}")
        lines += ["", "## 明日复习动作", ""]
        lines += [
            "1. 只重看上面列出的页面，不要重新通读整份 PPT。",
            "2. 每个错因写一句“以后看到什么信号就用什么方法”。",
            "3. 重新做 `07_小题练习.md` 中对应题目，错题写进错题本。",
        ]
    else:
        lines += ["## 高风险页面", ""]
        risky = []
        for slide in note_required_slides(extraction):
            number = int(slide.get("number", 0) or 0)
            note = notes.get(number, {})
            reasons = []
            if slide_has_formula(slide) or note_field(note, "formula_explanations", "formulas"):
                reasons.append("公式")
            if slide_has_visual(slide):
                reasons.append("图表")
            if plain(note_field(note, "difficulty")) in {"困难", "hard", "difficult"}:
                reasons.append("高难度")
            if note_field(note, "common_mistakes"):
                reasons.append("有常见错误")
            if reasons:
                risky.append((number, "、".join(reasons), plain(note_field(note, "exam_focus")) or slide.get("title", "")))
        if not risky:
            lines.append("- 暂未发现明显高风险页。先做 `07_小题练习.md`，再把错题填入模板。")
        for number, reason, focus in risky[:12]:
            lines.append(f"- S{number} [{reason}]: {focus}")
        lines += [
            "",
            "## 使用方式",
            "",
            "1. 做完小题后，把错题填入 `wrong_answer_template.json`。",
            "2. 重新运行并传入 `--wrong-answers-json wrong_answer_template.json`。",
            "3. 新的错题反馈路径会按错因和错题页重新排序。",
        ]
    output.write_text("\n".join(lines), encoding="utf-8")


def file_uri_or_empty(path_text: str) -> str:
    if not path_text:
        return ""
    path = Path(path_text)
    if path.exists():
        try:
            return path.resolve().as_uri()
        except ValueError:
            return ""
    return ""


def write_study_html(extraction: Dict[str, Any], notes: Dict[int, Dict[str, Any]], output: Path) -> None:
    modules = build_learning_modules(extraction, notes)
    practice_items: List[Dict[str, Any]] = []
    for slide in note_required_slides(extraction):
        number = int(slide.get("number", 0) or 0)
        practice_items.extend(build_slide_practice_questions(slide, notes.get(number, {}), max_items=1))
    nav_items = "\n".join(
        f'<a href="#module-{module["index"]}">模块 {module["index"]}: {html_escape(module["title"])}</a>' for module in modules
    )
    module_cards = []
    for module in modules:
        module_cards.append(
            f"""
<section class="panel" id="module-{module['index']}">
  <h2>{module['index']}. {html_escape(module['title'])}</h2>
  <p><strong>覆盖页面:</strong> {html_escape(slide_span(module['slides']))}</p>
  <p><strong>学习目标:</strong> {html_escape(module['goal'])}</p>
  <p><strong>必看页:</strong> {html_escape('；'.join(f"S{n}({r})" for n, r in module['focus']))}</p>
  <ul>{''.join(f"<li>{html_escape(item)}</li>" for item in module['checkpoints'])}</ul>
</section>
""".strip()
        )
    slide_cards = []
    for slide in note_required_slides(extraction):
        number = int(slide.get("number", 0) or 0)
        note = notes.get(number, {})
        image_uri = file_uri_or_empty(slide.get("screenshot", ""))
        formula_html = ""
        for formula in as_list(note_field(note, "formula_explanations", "formulas")):
            if isinstance(formula, dict) and formula.get("formula"):
                formula_html += f'<div class="formula">{html_escape(latex_to_readable_formula(formula.get("formula")) or formula.get("formula"))}</div>'
        slide_cards.append(
            f"""
<article class="slide-card">
  <h3>S{number}: {html_escape(plain(note_field(note, 'title') or slide.get('title', '')))}</h3>
  <p><strong>阅读层级:</strong> {html_escape(REVIEW_TIER_LABELS.get(slide.get('review_tier', 'deep'), slide.get('review_tier', 'deep')))}</p>
  {'<img src="' + html_escape(image_uri) + '" alt="slide screenshot">' if image_uri else ''}
  <p><strong>考点:</strong> {html_escape(plain(note_field(note, 'exam_focus')) or '待补写')}</p>
  <p><strong>核心:</strong> {html_escape(plain(note_field(note, 'key_takeaways')) or plain(note_field(note, 'what_it_says', 'summary')) or '待补写')}</p>
  {formula_html}
</article>
""".strip()
        )
    practice_cards = "\n".join(
        f"""
<details>
  <summary>S{item['slide']}: {html_escape(item['question'])}</summary>
  <p><strong>答案:</strong> {html_escape(item['answer'])}</p>
  <p><strong>思路:</strong> {html_escape(item['solution'])}</p>
</details>
""".strip()
        for item in practice_items
    )
    html = f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>PPT 学习页面</title>
  <style>
    body {{ margin: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; color: #172033; background: #f6f7fb; }}
    header {{ padding: 24px 32px; background: #172033; color: white; }}
    main {{ display: grid; grid-template-columns: 260px 1fr; gap: 24px; padding: 24px; }}
    nav {{ position: sticky; top: 16px; align-self: start; background: white; border: 1px solid #dde3ee; padding: 16px; }}
    nav a {{ display: block; color: #1d4ed8; text-decoration: none; margin: 8px 0; }}
    .panel, .slide-card, details {{ background: white; border: 1px solid #dde3ee; padding: 18px; margin-bottom: 16px; }}
    h1, h2, h3 {{ margin-top: 0; }}
    img {{ max-width: 100%; border: 1px solid #e5e7eb; }}
    .formula {{ font-family: "Cambria Math", "Times New Roman", serif; background: #ecfdf5; color: #064e3b; padding: 10px; text-align: center; font-size: 1.15rem; margin: 8px 0; }}
    @media (max-width: 800px) {{ main {{ grid-template-columns: 1fr; }} nav {{ position: static; }} }}
  </style>
</head>
<body>
  <header>
    <h1>PPT 学习页面</h1>
    <p>{html_escape(str(extraction.get('source', '')))}</p>
  </header>
  <main>
    <nav>
      <strong>学习路径</strong>
      {nav_items}
      <a href="#slides">逐页讲义</a>
      <a href="#practice">小题练习</a>
    </nav>
    <div>
      {''.join(module_cards)}
      <section class="panel" id="slides"><h2>逐页讲义</h2>{''.join(slide_cards)}</section>
      <section class="panel" id="practice"><h2>小题练习</h2>{practice_cards}</section>
    </div>
  </main>
</body>
</html>
"""
    output.write_text(html, encoding="utf-8")


def write_study_dashboard(extraction: Dict[str, Any], notes: Dict[int, Dict[str, Any]], output: Path, study_dir: Path) -> None:
    slides = study_slides(extraction)
    formula_slides = [slide.get("number") for slide in slides if slide_has_formula(slide)]
    visual_slides = [slide.get("number") for slide in slides if slide_has_visual(slide)]
    filled_notes = len(notes)
    total = len(slides)
    lines = [
        "# Study Pack Dashboard",
        "",
        f"- Source: {extraction.get('source', '')}",
        f"- Study slides: {total}",
        f"- {slide_filter_summary(extraction)}",
        f"- {review_plan_summary(extraction)}",
        f"- Skipped navigation slides: {extraction.get('skipped_navigation_slide_count', 0)}",
        f"- Slides with filled notes: {filled_notes}",
        f"- Formula-heavy slides: {formula_slides or 'none detected'}",
        f"- Visual-heavy slides: {visual_slides or 'none detected'}",
        "",
        "## Files",
        "",
        f"- [{(study_dir / 'review_index.md').name}](review_index.md)",
        f"- [{(study_dir / 'learning_path.md').name}](learning_path.md)",
        f"- [{(study_dir / 'exam_cram_plan.md').name}](exam_cram_plan.md)",
        f"- [{(study_dir / 'one_page_review.md').name}](one_page_review.md)",
        f"- [{(study_dir / 'active_recall_questions.md').name}](active_recall_questions.md)",
        f"- [{(study_dir / 'practice_questions.md').name}](practice_questions.md)",
        f"- [{(study_dir / 'study_index.html').name}](study_index.html)",
        f"- [{(study_dir / 'adaptive_review.md').name}](adaptive_review.md)",
        f"- [{(study_dir / 'formula_sheet.md').name}](formula_sheet.md)",
        f"- [{(study_dir / 'flashcards_anki.csv').name}](flashcards_anki.csv)",
        f"- [{(study_dir / 'flashcards.md').name}](flashcards.md)",
        f"- [{(study_dir / 'mistake_log_template.md').name}](mistake_log_template.md)",
        f"- [{(study_dir / 'concept_map.mmd').name}](concept_map.mmd)",
        "",
        "## How To Use",
        "",
        "1. Start with `review_index.md` to decide what is must-read, quick-scan, or reference only.",
        "2. Use `learning_path.md` to follow the chapter-style order.",
        "3. Read the handout pages for the current module only.",
        "4. Answer active recall questions without opening the slides.",
        "5. Do `practice_questions.md`, then read the answer and solution steps.",
        "6. Open `study_index.html` for a guided local review page.",
        "7. Import `flashcards_anki.csv` into Anki or review `flashcards.md` manually.",
        "8. Rework every formula from `formula_sheet.md` with a small example.",
        "9. Put every wrong answer into the mistake log and revisit it with `adaptive_review.md`.",
    ]
    output.write_text("\n".join(lines), encoding="utf-8")


def write_study_pack(
    extraction: Dict[str, Any],
    notes: Dict[int, Dict[str, Any]],
    study_dir: Path,
    exam_date: Optional[str],
    daily_minutes: int,
    target_score: Optional[str],
    wrong_answers: Optional[Dict[str, Any]] = None,
    practice_bank: Optional[List[Dict[str, Any]]] = None,
) -> None:
    study_dir.mkdir(parents=True, exist_ok=True)
    write_review_index(extraction, study_dir / "review_index.md")
    write_learning_path(extraction, notes, study_dir / "learning_path.md")
    write_study_dashboard(extraction, notes, study_dir / "README.md", study_dir)
    write_cram_plan(extraction, notes, study_dir / "exam_cram_plan.md", exam_date, daily_minutes, target_score)
    write_one_page_review(extraction, notes, study_dir / "one_page_review.md")
    write_formula_sheet(extraction, notes, study_dir / "formula_sheet.md")
    write_active_recall(extraction, notes, study_dir / "active_recall_questions.md", study_dir / "active_recall_questions.json")
    write_practice_questions(extraction, notes, study_dir / "practice_questions.md", study_dir / "practice_questions.json", practice_bank=practice_bank)
    write_adaptive_review(extraction, notes, study_dir / "adaptive_review.md", study_dir / "wrong_answer_template.json", wrong_answers=wrong_answers)
    write_study_html(extraction, notes, study_dir / "study_index.html")
    write_flashcards(extraction, notes, study_dir / "flashcards_anki.csv", study_dir / "flashcards.md")
    write_mistake_log(study_dir / "mistake_log_template.md")
    write_concept_map(extraction, notes, study_dir / "concept_map.mmd")


def copy_deliverable(src: Path, dst: Path) -> Optional[Path]:
    if not src.exists():
        return None
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)
    return dst


def relative_link(path: Path, base: Path) -> str:
    try:
        return path.relative_to(base).as_posix()
    except ValueError:
        return str(path)


def write_start_here(
    deliverables_dir: Path,
    source: Path,
    output_docx: Optional[Path],
    notes_markdown: Path,
    study_pack_dir: Optional[Path],
    workdir: Path,
    report: Dict[str, Any],
    notes_provided: bool,
) -> Dict[str, Path]:
    deliverables_dir.mkdir(parents=True, exist_ok=True)
    files: Dict[str, Path] = {}

    if output_docx and output_docx.exists():
        copied = copy_deliverable(output_docx, deliverables_dir / "01_复习讲义.docx")
        if copied:
            files["handout"] = copied
    if notes_markdown.exists():
        copied = copy_deliverable(notes_markdown, deliverables_dir / "02_完整讲义.md")
        if copied:
            files["markdown"] = copied

    if study_pack_dir and study_pack_dir.exists():
        mapping = [
            ("path", "learning_path.md", "00_学习路径.md"),
            ("overview", "one_page_review.md", "03_一页纸总览.md"),
            ("recall", "active_recall_questions.md", "04_主动回忆题.md"),
            ("formula", "formula_sheet.md", "05_公式速查.md"),
            ("mistakes", "mistake_log_template.md", "06_错题本模板.md"),
            ("practice", "practice_questions.md", "07_小题练习.md"),
            ("html", "study_index.html", "08_学习页面.html"),
            ("adaptive", "adaptive_review.md", "09_错题反馈路径.md"),
            ("wrong_template", "wrong_answer_template.json", "可选_错题输入模板.json"),
            ("anki", "flashcards_anki.csv", "可选_Anki卡片.csv"),
            ("plan", "exam_cram_plan.md", "可选_冲刺计划.md"),
            ("review_index", "review_index.md", "10_阅读取舍.md"),
        ]
        for key, src_name, dst_name in mapping:
            copied = copy_deliverable(study_pack_dir / src_name, deliverables_dir / dst_name)
            if copied:
                files[key] = copied

    quality_src = workdir / "quality_report.md"
    copied_quality = copy_deliverable(quality_src, deliverables_dir / "质量检查.md")
    if copied_quality:
        files["quality"] = copied_quality

    lines = [
        "# START HERE",
        "",
        "这是给学生看的入口页。不要从 `work/`、`extraction.json` 或截图文件夹开始看；那些是调试材料。",
        "",
        f"- 来源: `{source}`",
        f"- 质量分: {report.get('score', 0)}%",
        f"- 需要写讲解的页: {report.get('slide_count', 0)} / 学习页: {report.get('study_slide_count', report.get('slide_count', 0))} / 原始页: {report.get('total_slide_count', report.get('slide_count', 0))}",
        f"- 阅读分层: 必读深讲 {report.get('deep_slide_count', 0)} 页，快速扫读 {report.get('quick_slide_count', 0)} 页，参考/重复 {report.get('reference_slide_count', 0)} 页",
        f"- 已压缩标题/目录/章节过渡页: {report.get('skipped_navigation_slide_count', 0)}",
        "",
    ]
    skipped = report.get("skipped_navigation_slides") or []
    if skipped:
        preview = "、".join(f"S{item.get('number')}" for item in skipped[:12])
        lines += [
            "## 已自动压缩的页面",
            "",
            f"- {preview}{' 等' if len(skipped) > 12 else ''}",
            "- 这些页仍保留在调试目录的 `extraction.json` 里，但不会占用完整讲义、题目和学习路径篇幅。",
            "- 如果确实需要逐页全量输出，重新运行时加 `--content-filter all`。",
            "",
        ]
    if not notes_provided:
        lines += [
            "## 当前状态",
            "",
            "现在只是初稿骨架：已经完成截图、文字和素材提取，但还没有填入高质量逐页讲解。",
            "",
            "下一步：填写 `notes_template.json`，再重新运行并传入 `--notes-json`。",
            "",
        ]
    lines += [
        "## 先看这 5 个",
        "",
    ]
    priority = [
        ("1. 阅读取舍", "review_index", "先看哪些页必读、哪些页只扫、哪些页暂时不用看。"),
        ("2. 学习路径", "path", "按老师讲课顺序看，知道每一章先学什么、为什么学、学到什么程度。"),
        ("3. 复习讲义", "handout", "按学习路径指定的页面阅读，包含截图、考点、深度讲解、公式例题和常见错误。"),
        ("4. 主动回忆题", "recall", "闭卷答题。答不出来再回看讲义。"),
        ("5. 小题练习", "practice", "做基础题，再看答案和解题思路。"),
    ]
    for title, key, desc in priority:
        if key in files:
            lines.append(f"- **{title}**: [{files[key].name}]({relative_link(files[key], deliverables_dir)}) - {desc}")
        else:
            lines.append(f"- **{title}**: 未生成 - {desc}")
    lines += [
        "",
        "## 可选材料",
        "",
    ]
    optional = [
        ("一页纸总览", "overview", "考前快速过全局重点。"),
        ("学习页面", "html", "在浏览器里按模块复习，题目答案可折叠。"),
        ("错题反馈路径", "adaptive", "做完题后按错因重新安排复习。"),
        ("错题输入模板", "wrong_template", "把错题填进去后可用 --wrong-answers-json 生成二次学习路径。"),
        ("Anki 卡片", "anki", "导入 Anki 做间隔复习。"),
        ("错题本模板", "mistakes", "把不会的题、错因和复盘写进去。"),
        ("冲刺计划", "plan", "临近期末时按天安排。"),
        ("完整 Markdown", "markdown", "适合在编辑器里搜索和二次整理。"),
        ("质量检查", "quality", "看哪些页解释还不够深。"),
    ]
    for title, key, desc in optional:
        if key in files:
            lines.append(f"- [{title}]({relative_link(files[key], deliverables_dir)}) - {desc}")
    lines += [
        "",
        "## 不建议先看",
        "",
        f"- `{workdir}`: 调试目录，里面是截图、提取 JSON、prompt pack 和中间文件。",
        "- `extraction.json`: 给工具和开发者看的结构化素材，不是学生复习入口。",
        "- 单张截图文件夹: 只有在核对公式或图表时再打开。",
        "",
        "## 推荐复习顺序",
        "",
        "1. 先读 `10_阅读取舍.md`，确认必读、扫读和参考页。",
        "2. 读 `00_学习路径.md`，按模块决定今天看哪些页。",
        "3. 打开 `01_复习讲义.docx`，只读当前模块对应页面。",
        "4. 合上讲义，做 `04_主动回忆题.md`。",
        "5. 做 `07_小题练习.md`，先写答案，再看解题思路。",
        "6. 打开 `08_学习页面.html` 做折叠式复习。",
        "7. 错题写入 `06_错题本模板.md` 或 `可选_错题输入模板.json`。",
        "8. 考前只看 `03_一页纸总览.md`、`05_公式速查.md`、`09_错题反馈路径.md` 和错题本。",
        "",
    ]
    start = deliverables_dir / "START_HERE.md"
    start.write_text("\n".join(lines), encoding="utf-8")
    files["start"] = start
    return files


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Extract PPT/PDF slide content and build a Word study-notes document.")
    parser.add_argument("source", help="Input .pptx, .ppt, or slide .pdf")
    parser.add_argument("-o", "--output", help="Output .docx path")
    parser.add_argument("-w", "--workdir", help="Directory for screenshots, extracted images, and JSON files")
    parser.add_argument("--notes-json", help="Filled notes JSON to include in the final Word document")
    parser.add_argument("--ocr-json", help="Optional OCR/math-recognition JSON to merge into extracted slide text and formula candidates")
    parser.add_argument("--practice-bank-json", help="Optional open-education practice bank JSON with questions, answers, solutions, terms, and source URLs")
    parser.add_argument("--wrong-answers-json", help="Optional wrong-answer JSON used to generate adaptive review guidance")
    parser.add_argument("--notes-markdown", help="Output final study notes as Markdown; defaults to workdir/final_notes.md")
    parser.add_argument("--prompt-pack", help="Output LLM/VLM prompt pack; defaults to workdir/prompt_pack.md")
    parser.add_argument("--study-pack-dir", help="Directory for final-exam review pack; defaults to workdir/study_pack")
    parser.add_argument("--deliverables-dir", help="Clean student-facing output folder; defaults to output-name_deliverables")
    parser.add_argument("--no-study-pack", action="store_true", help="Skip final-exam review pack generation")
    parser.add_argument("--exam-date", help="Exam date in YYYY-MM-DD format for cram-plan generation")
    parser.add_argument("--daily-minutes", type=int, default=90, help="Daily study minutes for cram-plan generation, default: 90")
    parser.add_argument("--target-score", help="Target score label for cram-plan generation, for example: 90+")
    parser.add_argument("--language", default="zh", help="Notes template language label, default: zh")
    parser.add_argument("--dpi", type=int, default=180, help="Slide screenshot render DPI, default: 180")
    parser.add_argument("--slides", help="Only include selected slides, for example: 1,3-5,9")
    parser.add_argument("--max-slides", type=int, help="Only include the first N extracted slides")
    parser.add_argument("--content-filter", choices=["study", "all"], default="study", help="Default 'study' compacts title, agenda, and section divider slides; use 'all' for full per-slide output.")
    parser.add_argument("--review-depth", choices=["compressed", "balanced", "complete"], default="compressed", help="Reading-density mode. compressed is default for exam review; complete keeps old full per-slide behavior.")
    parser.add_argument("--max-deep-slides", type=int, help="Optional cap for must-read deep slides in compressed/balanced review.")
    parser.add_argument("--dedupe-threshold", type=float, default=0.86, help="Token-overlap threshold for marking repeated slides, default: 0.86")
    parser.add_argument("--no-render", action="store_true", help="Skip PDF conversion and screenshot rendering")
    parser.add_argument("--template-only", action="store_true", help="Create extraction files and notes template without building DOCX")
    parser.add_argument("--fail-under", type=float, help="Exit with status 1 if the notes quality score is below this percentage")
    parser.add_argument("--study-mode", choices=["notes", "final"], default="notes", help="Quality-report mode. Use 'final' to require exam-review fields.")
    parser.add_argument("--layout", choices=["study", "audit"], default="study", help="DOCX layout. 'study' is polished review handout; 'audit' includes full raw extraction.")
    parser.add_argument("--output-profile", choices=["teacher", "complete", "debug"], default="teacher", help="How much output to surface. teacher shows a clean START_HERE folder; complete/debug print all artifact paths.")
    parser.add_argument("--version", action="version", version=f"%(prog)s {VERSION}")
    args = parser.parse_args(argv)

    source = Path(args.source).expanduser().resolve()
    if not source.exists():
        print(f"Source file not found: {source}", file=sys.stderr)
        return 2

    output = Path(args.output).expanduser().resolve() if args.output else source.with_name(f"{source.stem}_PPT学习笔记.docx")
    workdir = Path(args.workdir).expanduser().resolve() if args.workdir else output.with_suffix("").with_name(output.stem + "_work")
    workdir.mkdir(parents=True, exist_ok=True)

    extraction = extract_source(source, workdir, args.dpi, args.no_render)
    try:
        filter_extraction_slides(extraction, args.slides, args.max_slides)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 2
    apply_ocr_json(extraction, Path(args.ocr_json).expanduser().resolve() if args.ocr_json else None)
    annotate_slide_kinds(extraction, args.content_filter)
    try:
        annotate_review_plan(extraction, args.review_depth, args.max_deep_slides, args.dedupe_threshold)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 2

    extraction_path = workdir / "extraction.json"
    extraction_path.write_text(json.dumps(extraction, ensure_ascii=False, indent=2), encoding="utf-8")

    template = make_notes_template(extraction, args.language)
    template_path = workdir / "notes_template.json"
    if not template_path.exists():
        template_path.write_text(json.dumps(template, ensure_ascii=False, indent=2), encoding="utf-8")

    markdown_path = workdir / "extraction.md"
    write_markdown(extraction, markdown_path)

    notes = load_notes(Path(args.notes_json).expanduser().resolve() if args.notes_json else None)
    practice_bank = load_practice_bank(Path(args.practice_bank_json).expanduser().resolve() if args.practice_bank_json else None)
    wrong_answers = load_wrong_answers(Path(args.wrong_answers_json).expanduser().resolve() if args.wrong_answers_json else None)
    notes_markdown_path = Path(args.notes_markdown).expanduser().resolve() if args.notes_markdown else workdir / "final_notes.md"
    prompt_pack_path = Path(args.prompt_pack).expanduser().resolve() if args.prompt_pack else workdir / "prompt_pack.md"
    write_notes_markdown(extraction, notes, notes_markdown_path)
    write_prompt_pack(extraction, prompt_pack_path, args.language)
    study_pack_dir = Path(args.study_pack_dir).expanduser().resolve() if args.study_pack_dir else workdir / "study_pack"
    if not args.no_study_pack:
        write_study_pack(
            extraction,
            notes,
            study_pack_dir,
            args.exam_date,
            args.daily_minutes,
            args.target_score,
            wrong_answers=wrong_answers,
            practice_bank=practice_bank,
        )

    report = build_quality_report(extraction, notes, args.fail_under, args.study_mode)
    quality_json_path = workdir / "quality_report.json"
    quality_md_path = workdir / "quality_report.md"
    write_quality_report(report, quality_json_path, quality_md_path)

    built_docx: Optional[Path] = None
    if not args.template_only:
        build_docx(extraction, notes, output, layout=args.layout)
        built_docx = output
        print(f"DOCX: {output}")
    else:
        print("DOCX: skipped (--template-only)")

    deliverables_dir = Path(args.deliverables_dir).expanduser().resolve() if args.deliverables_dir else output.with_suffix("").with_name(output.stem + "_deliverables")
    if args.output_profile in {"teacher", "complete"}:
        delivered = write_start_here(
            deliverables_dir,
            source,
            built_docx,
            notes_markdown_path,
            None if args.no_study_pack else study_pack_dir,
            workdir,
            report,
            bool(args.notes_json),
        )
        print(f"START HERE: {delivered['start']}")

    if args.output_profile in {"complete", "debug"}:
        print(f"Extraction JSON: {extraction_path}")
        print(f"Notes template: {template_path}")
        print(f"Markdown extraction: {markdown_path}")
        print(f"Final notes Markdown: {notes_markdown_path}")
        print(f"Prompt pack: {prompt_pack_path}")
        if not args.no_study_pack:
            print(f"Study pack: {study_pack_dir}")
        print(f"Quality report JSON: {quality_json_path}")
        print(f"Quality report Markdown: {quality_md_path}")
    else:
        print(f"Clean deliverables: {deliverables_dir}")
        print(f"Debug workdir: {workdir}")

    print(f"Quality score: {report['score']}%")
    print(f"Content filter: {slide_filter_summary(extraction)}")
    print(f"Review plan: {review_plan_summary(extraction)}")
    if extraction.get("warnings"):
        print("Warnings:")
        for warning in extraction["warnings"]:
            print(f"- {warning}")
    if not args.notes_json:
        print("Next: fill notes_template.json, then rerun with --notes-json to produce completed explanatory notes.")
    if args.fail_under is not None and not report["passed"]:
        print(f"Quality score {report['score']}% is below --fail-under {args.fail_under}%.", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
