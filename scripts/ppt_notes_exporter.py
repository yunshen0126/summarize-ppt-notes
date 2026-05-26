#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import hashlib
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
}

VERSION = "0.4.0"


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


def make_notes_template(extraction: Dict[str, Any], language: str) -> Dict[str, Any]:
    slides = []
    for slide in extraction.get("slides", []):
        slides.append(
            {
                "number": slide.get("number"),
                "title": slide.get("title", ""),
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
                "common_mistakes": [],
                "prerequisites": [],
                "difficulty": "",
                "estimated_review_minutes": "",
                "tags": [],
            }
        )
    return {"language": language, "source": extraction.get("source", ""), "slides": slides}


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


def field_has_content(note: Dict[str, Any], min_chars: int, *keys: str) -> bool:
    return len(meaningful_text(note_field(note, *keys))) >= min_chars


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
    for slide in extraction.get("slides", []):
        number = int(slide.get("number", 0) or 0)
        note = notes.get(number, {})
        checks = [
            ("purpose", "这一页是干什么用的", field_has_content(note, 12, "purpose", "page_purpose")),
            ("what_it_says", "这一页讲了什么", field_has_content(note, 24, "what_it_says", "summary")),
            ("detailed_explanation", "复杂内容详解", field_has_content(note, 32, "detailed_explanation", "complex_explanation")),
        ]
        if slide_has_visual(slide):
            checks.append(("visual_explanation", "图片/图表/表格说明", field_has_content(note, 20, "visual_explanation", "image_explanation")))
        if slide_has_formula(slide):
            checks.append(("formula_explanations", "公式说明", bool(note_field(note, "formula_explanations", "formulas"))))
            checks.append(("worked_examples", "公式或方法例题", bool(note_field(note, "worked_examples", "examples"))))
        if study_mode == "final":
            checks.extend(
                [
                    ("exam_focus", "期末考点定位", field_has_content(note, 12, "exam_focus")),
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
            }
        )

    score = round((earned_points / total_points) * 100, 2) if total_points else 0.0
    passed = fail_under is None or score >= fail_under
    return {
        "tool": "summarize-ppt-notes",
        "version": VERSION,
        "source": extraction.get("source", ""),
        "slide_count": extraction.get("slide_count", 0),
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
        f"- Slides: {report.get('slide_count', 0)}",
        f"- Passed: {report.get('passed')}",
        "",
        "## Slide Checks",
        "",
    ]
    for slide in report.get("slides", []):
        missing = slide.get("missing") or []
        status = "PASS" if not missing else "TODO"
        lines.append(f"### Slide {slide.get('number')}: {status} ({slide.get('score')}%)")
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


class DocxBuilder:
    def __init__(self, title: str) -> None:
        self.title = title
        self.body: List[str] = []
        self.image_rels: List[Tuple[str, str, Path]] = []
        self.image_id = 1

    def add_heading(self, text: str, level: int = 1) -> None:
        style = "Heading1" if level == 1 else "Heading2"
        self.add_paragraph(text, style=style)

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

    def add_bullet(self, text: Any) -> None:
        self.add_paragraph(f"- {text}", style="ListBullet")

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
      <w:pgMar w:top="850" w:right="850" w:bottom="850" w:left="850" w:header="708" w:footer="708" w:gutter="0"/>
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
    <w:rPr><w:rFonts w:ascii="Microsoft YaHei" w:eastAsia="Microsoft YaHei" w:hAnsi="Microsoft YaHei"/><w:sz w:val="21"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr><w:spacing w:before="360" w:after="180"/><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:b/><w:rFonts w:ascii="Microsoft YaHei" w:eastAsia="Microsoft YaHei" w:hAnsi="Microsoft YaHei"/><w:sz w:val="34"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="1"/></w:pPr>
    <w:rPr><w:b/><w:rFonts w:ascii="Microsoft YaHei" w:eastAsia="Microsoft YaHei" w:hAnsi="Microsoft YaHei"/><w:sz w:val="26"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListBullet">
    <w:name w:val="List Bullet"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:ind w:left="360" w:hanging="180"/></w:pPr>
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
    doc.add_paragraph(label, bold=True)
    if isinstance(value, list):
        if value:
            for item in value:
                doc.add_bullet(item if not isinstance(item, dict) else json.dumps(item, ensure_ascii=False))
        else:
            doc.add_paragraph(placeholder)
    elif value:
        doc.add_paragraph(value)
    else:
        doc.add_paragraph(placeholder)


def add_formula_notes(doc: DocxBuilder, formulas: Any) -> None:
    doc.add_paragraph("公式说明和例题", bold=True)
    if not formulas:
        doc.add_paragraph("无或待补写。")
        return
    if isinstance(formulas, str):
        doc.add_paragraph(formulas)
        return
    if not isinstance(formulas, list):
        doc.add_paragraph(json.dumps(formulas, ensure_ascii=False))
        return
    for item in formulas:
        if isinstance(item, dict):
            formula = item.get("formula", "")
            meaning = item.get("meaning", "")
            conditions = item.get("conditions", "")
            example = item.get("example", "")
            doc.add_bullet(f"公式：{formula}" if formula else "公式：未填写")
            if meaning:
                doc.add_paragraph(f"含义：{meaning}")
            if conditions:
                doc.add_paragraph(f"条件：{conditions}")
            if example:
                doc.add_paragraph(f"例题：{example}")
        else:
            doc.add_bullet(item)


def build_docx(extraction: Dict[str, Any], notes: Dict[int, Dict[str, Any]], output: Path) -> None:
    title = f"PPT学习笔记 - {Path(extraction.get('source', 'slides')).stem}"
    doc = DocxBuilder(title)
    doc.add_heading(title, 1)
    doc.add_paragraph(f"来源文件：{extraction.get('source', '')}")
    doc.add_paragraph(f"生成时间：{extraction.get('generated_at', '')}")
    doc.add_paragraph(f"页数：{extraction.get('slide_count', 0)}")

    warnings = extraction.get("warnings") or []
    if warnings:
        add_list_section(doc, "处理提示", warnings)

    for slide in extraction.get("slides", []):
        number = int(slide.get("number", 0) or 0)
        note = notes.get(number, {})
        title_text = note_field(note, "title") or slide.get("title") or f"Slide {number}"
        doc.add_heading(f"第 {number} 页：{title_text}", 1)

        screenshot = slide.get("screenshot", "")
        doc.add_heading("页面截图", 2)
        if not doc.add_image(screenshot, max_width_inches=6.7):
            doc.add_paragraph("未生成页面截图。")

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

        formulas = slide.get("formulas") or []
        candidates = slide.get("formula_candidates") or []
        doc.add_heading("公式与候选表达式", 2)
        if not formulas and not candidates:
            doc.add_paragraph("未从结构化文本中提取到公式；仍需根据页面截图确认是否存在图片公式。")
        for formula in formulas:
            text = formula.get("text") or "[OMML formula]"
            doc.add_bullet(text)
        for candidate in candidates:
            doc.add_bullet(candidate)

        doc.add_heading("讲解笔记", 2)
        add_note_text(doc, "这一页是干什么用的", note_field(note, "purpose", "page_purpose"), "待补写：说明本页在整套 PPT 中的作用。")
        add_note_text(doc, "这一页讲了什么", note_field(note, "what_it_says", "summary"), "待补写：完整解释本页内容。")
        add_note_text(doc, "复杂内容详解", note_field(note, "detailed_explanation", "complex_explanation"), "待补写：对复杂概念、推导、图表或算法做展开说明。")
        add_note_text(doc, "图片/图表说明", note_field(note, "visual_explanation", "image_explanation"), "待补写：说明图片、图表、流程图、架构图等视觉元素。")
        add_formula_notes(doc, note_field(note, "formula_explanations", "formulas"))
        add_note_text(doc, "补充例题/案例", note_field(note, "worked_examples", "examples"), "无或待补写。")
        add_note_text(doc, "期末考点定位", note_field(note, "exam_focus"), "待补写：说明本页在期末考试中可能怎么考。")
        add_note_text(doc, "核心记忆点", note_field(note, "key_takeaways"), "待补写：列出必须背会或能复述的要点。")
        add_note_text(doc, "记忆钩子", note_field(note, "memory_hooks"), "无或待补写。")
        add_note_text(doc, "可能考法/自测题", note_field(note, "likely_questions"), "待补写：生成主动回忆题或考试风格问题。")
        add_note_text(doc, "常见错误", note_field(note, "common_mistakes"), "待补写：说明学生容易错在哪里。")
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
        f"- 页数: {extraction.get('slide_count', 0)}",
        f"- 工具版本: {extraction.get('version', VERSION)}",
        "",
    ]
    warnings = extraction.get("warnings") or []
    if warnings:
        lines += ["## 处理提示", ""]
        lines += [f"- {warning}" for warning in warnings]
        lines.append("")

    for slide in extraction.get("slides", []):
        number = int(slide.get("number", 0) or 0)
        note = notes.get(number, {})
        slide_title = note_field(note, "title") or slide.get("title") or f"Slide {number}"
        lines += [f"## 第 {number} 页：{slide_title}", ""]
        if slide.get("screenshot"):
            lines += [f"![第 {number} 页截图]({slide.get('screenshot')})", ""]

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
        "## Required Output",
        "",
        "- Return valid JSON matching `references/note-schema.md`.",
        "- Keep one object per slide.",
        "- Explain each slide's purpose, content, complex ideas, visual elements, formulas, and examples.",
        "- Add final-exam fields: exam_focus, key_takeaways, memory_hooks, likely_questions, common_mistakes, prerequisites, difficulty, estimated_review_minutes, and tags.",
        "- likely_questions should include active-recall questions and at least one exam-style question for important formulas or algorithms.",
        "- Mark uncertain visual or formula recognition as `需核对`.",
        f"- Output language: {language}.",
        "",
        "## Source",
        "",
        f"- File: {extraction.get('source', '')}",
        f"- SHA256: {extraction.get('source_sha256', '')}",
        f"- Slides: {extraction.get('slide_count', 0)}",
        "",
    ]
    for slide in extraction.get("slides", []):
        lines += [
            f"## Slide {slide.get('number')}: {slide.get('title', '')}",
            "",
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
    for slide in extraction.get("slides", []):
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
                lines.append(f"- Formula: `{item.get('formula', '未填写')}`")
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
    for slide in extraction.get("slides", []):
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


def write_flashcards(extraction: Dict[str, Any], notes: Dict[int, Dict[str, Any]], csv_path: Path, md_path: Path) -> None:
    rows = [["Front", "Back", "Tags", "SourceSlide"]]
    md_lines = ["# Flashcards", ""]
    for slide in extraction.get("slides", []):
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
    slide_count = len(extraction.get("slides", []))
    minutes = max(daily_minutes, 20)
    target = target_score or "按课程目标自定"
    lines = [
        "# Final Exam Cram Plan",
        "",
        f"- Exam date: {exam_date or '未设置'}",
        f"- Days left: {days if days is not None else '未设置'}",
        f"- Daily minutes: {minutes}",
        f"- Target score: {target}",
        f"- Slides: {slide_count}",
        "",
        "## Daily Loop",
        "",
        "1. 先看 `one_page_review.md` 建立全局框架。",
        "2. 用 `active_recall_questions.md` 闭卷回答，答不出就回看对应页。",
        "3. 复习 `formula_sheet.md`，每个公式至少手算一个例子。",
        "4. 用 `mistake_log_template.md` 记录错因，而不是只记录答案。",
        "5. 睡前用 `flashcards_anki.csv` 或 `flashcards.md` 快速回忆。",
        "",
    ]
    if days is None or days >= 7:
        phases = [
            ("Day 1-2", "通读所有页面，标记不会的公式、图表和定义。"),
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
    for slide in extraction.get("slides", []):
        number = int(slide.get("number", 0) or 0)
        note = notes.get(number, {})
        difficulty = note_field(note, "difficulty") or ("困难" if slide_has_formula(slide) else "中等" if slide_has_visual(slide) else "基础")
        focus = note_field(note, "exam_focus") or note_field(note, "purpose") or slide.get("title", "")
        est = note_field(note, "estimated_review_minutes") or ("12" if difficulty == "困难" else "8" if difficulty == "中等" else "5")
        lines.append(f"- Slide {number} [{difficulty}, {est} min]: {focus}")
    output.write_text("\n".join(lines), encoding="utf-8")


def write_one_page_review(extraction: Dict[str, Any], notes: Dict[int, Dict[str, Any]], output: Path) -> None:
    lines = ["# One Page Review", "", "## High-Yield Points", ""]
    for slide in extraction.get("slides", []):
        number = int(slide.get("number", 0) or 0)
        note = notes.get(number, {})
        takeaways = as_list(note_field(note, "key_takeaways")) or [note_field(note, "exam_focus") or note_field(note, "what_it_says", "summary") or slide.get("title", "")]
        for item in takeaways[:3]:
            if plain(item):
                lines.append(f"- S{number}: {plain(item)}")
    lines += ["", "## Must-Check Mistakes", ""]
    has_mistake = False
    for slide in extraction.get("slides", []):
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
    for slide in extraction.get("slides", []):
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


def write_study_dashboard(extraction: Dict[str, Any], notes: Dict[int, Dict[str, Any]], output: Path, study_dir: Path) -> None:
    formula_slides = [slide.get("number") for slide in extraction.get("slides", []) if slide_has_formula(slide)]
    visual_slides = [slide.get("number") for slide in extraction.get("slides", []) if slide_has_visual(slide)]
    filled_notes = len(notes)
    total = len(extraction.get("slides", []))
    lines = [
        "# Study Pack Dashboard",
        "",
        f"- Source: {extraction.get('source', '')}",
        f"- Slides: {total}",
        f"- Slides with filled notes: {filled_notes}",
        f"- Formula-heavy slides: {formula_slides or 'none detected'}",
        f"- Visual-heavy slides: {visual_slides or 'none detected'}",
        "",
        "## Files",
        "",
        f"- [{(study_dir / 'exam_cram_plan.md').name}](exam_cram_plan.md)",
        f"- [{(study_dir / 'one_page_review.md').name}](one_page_review.md)",
        f"- [{(study_dir / 'active_recall_questions.md').name}](active_recall_questions.md)",
        f"- [{(study_dir / 'formula_sheet.md').name}](formula_sheet.md)",
        f"- [{(study_dir / 'flashcards_anki.csv').name}](flashcards_anki.csv)",
        f"- [{(study_dir / 'flashcards.md').name}](flashcards.md)",
        f"- [{(study_dir / 'mistake_log_template.md').name}](mistake_log_template.md)",
        f"- [{(study_dir / 'concept_map.mmd').name}](concept_map.mmd)",
        "",
        "## How To Use",
        "",
        "1. Start with one-page review.",
        "2. Answer active recall questions without opening the slides.",
        "3. Import `flashcards_anki.csv` into Anki or review `flashcards.md` manually.",
        "4. Rework every formula from `formula_sheet.md` with a small example.",
        "5. Put every wrong answer into the mistake log and revisit it the next day.",
    ]
    output.write_text("\n".join(lines), encoding="utf-8")


def write_study_pack(
    extraction: Dict[str, Any],
    notes: Dict[int, Dict[str, Any]],
    study_dir: Path,
    exam_date: Optional[str],
    daily_minutes: int,
    target_score: Optional[str],
) -> None:
    study_dir.mkdir(parents=True, exist_ok=True)
    write_study_dashboard(extraction, notes, study_dir / "README.md", study_dir)
    write_cram_plan(extraction, notes, study_dir / "exam_cram_plan.md", exam_date, daily_minutes, target_score)
    write_one_page_review(extraction, notes, study_dir / "one_page_review.md")
    write_formula_sheet(extraction, notes, study_dir / "formula_sheet.md")
    write_active_recall(extraction, notes, study_dir / "active_recall_questions.md", study_dir / "active_recall_questions.json")
    write_flashcards(extraction, notes, study_dir / "flashcards_anki.csv", study_dir / "flashcards.md")
    write_mistake_log(study_dir / "mistake_log_template.md")
    write_concept_map(extraction, notes, study_dir / "concept_map.mmd")


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Extract PPT/PDF slide content and build a Word study-notes document.")
    parser.add_argument("source", help="Input .pptx, .ppt, or slide .pdf")
    parser.add_argument("-o", "--output", help="Output .docx path")
    parser.add_argument("-w", "--workdir", help="Directory for screenshots, extracted images, and JSON files")
    parser.add_argument("--notes-json", help="Filled notes JSON to include in the final Word document")
    parser.add_argument("--notes-markdown", help="Output final study notes as Markdown; defaults to workdir/final_notes.md")
    parser.add_argument("--prompt-pack", help="Output LLM/VLM prompt pack; defaults to workdir/prompt_pack.md")
    parser.add_argument("--study-pack-dir", help="Directory for final-exam review pack; defaults to workdir/study_pack")
    parser.add_argument("--no-study-pack", action="store_true", help="Skip final-exam review pack generation")
    parser.add_argument("--exam-date", help="Exam date in YYYY-MM-DD format for cram-plan generation")
    parser.add_argument("--daily-minutes", type=int, default=90, help="Daily study minutes for cram-plan generation, default: 90")
    parser.add_argument("--target-score", help="Target score label for cram-plan generation, for example: 90+")
    parser.add_argument("--language", default="zh", help="Notes template language label, default: zh")
    parser.add_argument("--dpi", type=int, default=180, help="Slide screenshot render DPI, default: 180")
    parser.add_argument("--slides", help="Only include selected slides, for example: 1,3-5,9")
    parser.add_argument("--max-slides", type=int, help="Only include the first N extracted slides")
    parser.add_argument("--no-render", action="store_true", help="Skip PDF conversion and screenshot rendering")
    parser.add_argument("--template-only", action="store_true", help="Create extraction files and notes template without building DOCX")
    parser.add_argument("--fail-under", type=float, help="Exit with status 1 if the notes quality score is below this percentage")
    parser.add_argument("--study-mode", choices=["notes", "final"], default="notes", help="Quality-report mode. Use 'final' to require exam-review fields.")
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

    extraction_path = workdir / "extraction.json"
    extraction_path.write_text(json.dumps(extraction, ensure_ascii=False, indent=2), encoding="utf-8")

    template = make_notes_template(extraction, args.language)
    template_path = workdir / "notes_template.json"
    if not template_path.exists():
        template_path.write_text(json.dumps(template, ensure_ascii=False, indent=2), encoding="utf-8")

    markdown_path = workdir / "extraction.md"
    write_markdown(extraction, markdown_path)

    notes = load_notes(Path(args.notes_json).expanduser().resolve() if args.notes_json else None)
    notes_markdown_path = Path(args.notes_markdown).expanduser().resolve() if args.notes_markdown else workdir / "final_notes.md"
    prompt_pack_path = Path(args.prompt_pack).expanduser().resolve() if args.prompt_pack else workdir / "prompt_pack.md"
    write_notes_markdown(extraction, notes, notes_markdown_path)
    write_prompt_pack(extraction, prompt_pack_path, args.language)
    study_pack_dir = Path(args.study_pack_dir).expanduser().resolve() if args.study_pack_dir else workdir / "study_pack"
    if not args.no_study_pack:
        write_study_pack(extraction, notes, study_pack_dir, args.exam_date, args.daily_minutes, args.target_score)

    report = build_quality_report(extraction, notes, args.fail_under, args.study_mode)
    quality_json_path = workdir / "quality_report.json"
    quality_md_path = workdir / "quality_report.md"
    write_quality_report(report, quality_json_path, quality_md_path)

    if not args.template_only:
        build_docx(extraction, notes, output)
        print(f"DOCX: {output}")
    else:
        print("DOCX: skipped (--template-only)")
    print(f"Extraction JSON: {extraction_path}")
    print(f"Notes template: {template_path}")
    print(f"Markdown extraction: {markdown_path}")
    print(f"Final notes Markdown: {notes_markdown_path}")
    print(f"Prompt pack: {prompt_pack_path}")
    if not args.no_study_pack:
        print(f"Study pack: {study_pack_dir}")
    print(f"Quality report JSON: {quality_json_path}")
    print(f"Quality report Markdown: {quality_md_path}")
    print(f"Quality score: {report['score']}%")
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
