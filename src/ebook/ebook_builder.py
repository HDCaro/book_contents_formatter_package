#!/usr/bin/env python3

# ================================================================================
# EBOOK BUILDER MODULE
# ================================================================================
# PURPOSE:
# Build an EPUB from front matter + book body DOCX files.
#
# DESIGN:
# - Config-based paths and metadata (ebook_builder.config.json)
# - Word COM export DOCX -> filtered HTML
# - Heading extraction reusing the same detector logic as front_matter_builder.py
# - Robust Word preflight cleanup (word killer)
# - Verbose logging in the same spirit as front_matter_builder.py
# ================================================================================

from __future__ import annotations

import argparse
import json
from html import escape
import os
import posixpath
import re
import shutil
import subprocess
import sys
import unicodedata
import uuid
import zipfile
from datetime import datetime
from pathlib import Path

import win32com.client
from bs4 import BeautifulSoup
from ebooklib import ITEM_DOCUMENT, epub


HTML_PARSER = "html.parser"
EPUB_PACKAGE_ROOT = "OEBPS"


def log(message: str) -> None:
    print(message, flush=True)


def kill_running_word_instances() -> None:
    """Terminate running WINWORD processes before COM automation starts."""
    try:
        check = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/NH"],
            capture_output=True,
            text=True,
            check=False,
        )
        output = (check.stdout or "").strip().lower()

        if not output or "no tasks are running" in output:
            log("   [INFO] No existing Word instances found")
            return

        kill = subprocess.run(
            ["taskkill", "/IM", "WINWORD.EXE", "/F"],
            capture_output=True,
            text=True,
            check=False,
        )

        if kill.returncode == 0:
            log("   [OK] Terminated running WINWORD instance(s)")
        else:
            err = (kill.stderr or kill.stdout or "").strip()
            log(f"   [WARN] Could not terminate WINWORD processes: {err}")

    except Exception as exc:
        log(f"   [WARN] Word preflight cleanup warning: {exc}")


def find_project_root() -> Path:
    current = Path(__file__).resolve()
    for parent in [current.parent] + list(current.parents):
        if (parent / "src").exists() and (parent / "data").exists():
            return parent
    raise RuntimeError("Project root not found (expected folders: src and data)")


def resolve_config_path(project_root: Path, configured_path: str) -> Path:
    path = Path(configured_path)
    if path.is_absolute():
        return path
    return project_root / path


def find_file_by_name(project_root: Path, filename: str) -> list[Path]:
    target = filename.lower()
    matches: list[Path] = []
    for root, _, files in os.walk(project_root):
        for name in files:
            if name.lower() == target:
                matches.append(Path(root) / name)
    matches.sort(key=lambda p: (len(p.parts), str(p).lower()))
    return matches


def clean_text(text: str) -> str:
    if not text:
        return ""
    cleaned = "".join(c for c in text if c >= " " or c in ("\n", "\t")).strip()
    cleaned = cleaned.rstrip("/\\").strip()
    return cleaned


def clean_title_text(text: str) -> str:
    if not text:
        return ""

    cleaned = text.replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
    cleaned = re.sub(r"\s+", " ", cleaned)
    cleaned = "".join(c for c in cleaned if c >= " ")
    cleaned = cleaned.rstrip("/\\").strip()
    return cleaned


def get_word():
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    return word


def load_json(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def infer_inputs_from_front_matter_config(project_root: Path, cfg: dict) -> dict:
    options = cfg.get("options", {})
    front_cfg_rel = options.get(
        "front_matter_config",
        "src/front_matter/builders/front_matter_builder.config.json",
    )
    front_cfg_path = resolve_config_path(project_root, front_cfg_rel)
    if not front_cfg_path.exists():
        return {}

    front_cfg = load_json(front_cfg_path)
    inferred: dict = {}

    inferred["book_body_file"] = front_cfg["inputs"]["book_body_file"]

    output_dir = front_cfg["outputs"]["output_dir"]
    output_name = front_cfg["outputs"]["filename"]
    if output_dir and output_name:
        inferred["front_matter_docx"] = str(Path(output_dir) / output_name)

    return inferred


def resolve_input_file(
    project_root: Path,
    inputs: dict,
    key: str,
    *,
    required: bool,
    auto_discover: bool,
) -> Path | None:
    raw_value = str(inputs.get(key, "")).strip()
    if not raw_value:
        if required:
            raise ValueError(f"Missing required input key: {key}")
        return None

    path = resolve_config_path(project_root, raw_value)
    if path.exists():
        return path

    if auto_discover:
        matches = find_file_by_name(project_root, path.name)
        if matches:
            discovered = matches[0]
            log(f"   [OK] Auto-discovered {key}: {discovered}")
            return discovered

    if required:
        raise FileNotFoundError(f"Input file not found for '{key}': {path}")
    return None


def load_builder_config(project_root: Path) -> dict:
    log("\n[CONFIG] Loading ebook builder config...")

    config_path = project_root / "src" / "ebook" / "ebook_builder.config.json"
    if not config_path.exists():
        raise FileNotFoundError(f"Configuration file not found: {config_path}")

    cfg = load_json(config_path)
    inputs = cfg.get("inputs", {})
    outputs = cfg.get("outputs", {})
    metadata = cfg.get("metadata", {})
    options = cfg.get("options", {})

    for required_key in ("output_dir", "filename", "temp_dir"):
        if not str(outputs.get(required_key, "")).strip():
            raise ValueError(f"Missing required outputs key: {required_key}")

    if options.get("fallback_from_front_matter_config", True):
        inferred = infer_inputs_from_front_matter_config(project_root, cfg)
        for key, value in inferred.items():
            if not str(inputs.get(key, "")).strip():
                inputs[key] = value

    auto_discover = bool(inputs.get("auto_discover_missing_inputs", True))

    front_matter_docx = resolve_input_file(
        project_root,
        inputs,
        "front_matter_docx",
        required=False,
        auto_discover=auto_discover,
    )
    book_body_file = resolve_input_file(
        project_root,
        inputs,
        "book_body_file",
        required=True,
        auto_discover=auto_discover,
    )
    cover_image_file = resolve_input_file(
        project_root,
        inputs,
        "cover_image_file",
        required=False,
        auto_discover=auto_discover,
    )
    back_cover_image_file = resolve_input_file(
        project_root,
        inputs,
        "back_cover_image_file",
        required=False,
        auto_discover=auto_discover,
    )
    copyright_page_docx = resolve_input_file(
        project_root,
        inputs,
        "copyright_page_docx",
        required=False,
        auto_discover=auto_discover,
    )

    output_dir = resolve_config_path(project_root, outputs["output_dir"])
    temp_dir = resolve_config_path(project_root, outputs["temp_dir"])
    output_epub = output_dir / outputs["filename"]
    mobi_filename = str(outputs.get("mobi_filename", f"{output_epub.stem}.mobi")).strip() or f"{output_epub.stem}.mobi"
    output_mobi = output_dir / mobi_filename

    output_dir.mkdir(parents=True, exist_ok=True)
    temp_dir.mkdir(parents=True, exist_ok=True)

    return {
        "config_path": config_path,
        "front_matter_docx": front_matter_docx,
        "book_body_file": book_body_file,
        "cover_image_file": cover_image_file,
        "back_cover_image_file": back_cover_image_file,
        "copyright_page_docx": copyright_page_docx,
        "output_dir": output_dir,
        "temp_dir": temp_dir,
        "output_epub": output_epub,
        "output_mobi": output_mobi,
        "keep_html_exports": bool(options.get("keep_html_exports", True)),
        "kindlegen_executable": str(options.get("kindlegen_executable", "")).strip(),
        "title": metadata.get("title", "Untitled"),
        "author": metadata.get("author", "Unknown"),
        "language": metadata.get("language", "es"),
        "identifier": metadata.get("identifier") or str(uuid.uuid4()),
        "publisher": metadata.get("publisher", ""),
        "publication_year": metadata.get("publication_year", ""),
        "isbn": metadata.get("isbn", ""),
        "copyright_holder": metadata.get("copyright_holder", metadata.get("author", "Unknown")),
    }


def export_docx_to_filtered_html(docx_path: Path, output_dir: Path) -> Path:
    if not docx_path.exists():
        raise FileNotFoundError(f"DOCX input not found: {docx_path}")

    output_dir.mkdir(parents=True, exist_ok=True)
    html_path = output_dir / f"{docx_path.stem}.html"

    log(f"\n[WORD] Exporting DOCX -> filtered HTML: {docx_path.name}")
    word = get_word()
    doc = None
    try:
        doc = word.Documents.Open(str(docx_path))
        wd_format_filtered_html = 10
        doc.SaveAs(str(html_path), FileFormat=wd_format_filtered_html)
    finally:
        if doc is not None:
            doc.Close(False)
        word.Quit()

    if not html_path.exists():
        raise RuntimeError(f"Word export failed: {html_path}")

    log(f"   [OK] HTML generated: {html_path}")
    return html_path


def extract_headings(doc_path: Path) -> list[dict]:
    """Same detector logic and verbosity style as front_matter_builder.py."""
    log("\n[HEADINGS] Extracting headings with consecutive Heading 2 collection")

    word = get_word()
    doc = word.Documents.Open(os.path.abspath(str(doc_path)))
    doc.Repaginate()

    headings: list[dict] = []
    total = doc.Paragraphs.Count

    i = 1
    while i <= total:
        try:
            para = doc.Paragraphs(i)
            text = clean_text(para.Range.Text)
            style = para.Style.NameLocal.lower()

            if "heading 1" in style and text:
                log(f"\n   [H1] Found Heading 1: '{text}'")

                if text.lower().startswith("chapter"):
                    log(f"   [SCAN] Processing chapter: '{text}'")

                    chapter = text
                    chapter_title_parts: list[str] = []
                    j = i + 1

                    log(f"   [SCAN] Looking for consecutive Heading 2 from paragraph {j}...")

                    while j <= total:
                        try:
                            p = doc.Paragraphs(j)
                            p_style = p.Style.NameLocal.lower()
                            raw_text = p.Range.Text
                            cleaned_text = clean_title_text(raw_text)

                            log(f"   [P{j}] '{cleaned_text}' (style: {p_style})")

                            if not raw_text or not raw_text.strip():
                                log(f"   [SKIP] Empty paragraph {j}")
                                j += 1
                                continue

                            if "heading 2" in p_style:
                                chapter_title_parts.append(cleaned_text)
                                log(f"   [OK] Collected Heading 2 part: '{cleaned_text}'")
                                j += 1
                                continue

                            if "heading 1" in p_style:
                                log("   [STOP] Hit another Heading 1")
                                break

                            log("   [STOP] Hit non-heading style")
                            break

                        except Exception as exc:
                            log(f"   [WARN] Error processing paragraph {j}: {exc}")
                            break

                    if chapter_title_parts:
                        complete_title = " ".join(chapter_title_parts)
                        full = f"{chapter}: {complete_title}"
                        log(f"   [TITLE] Complete chapter title: '{full}'")
                    else:
                        full = chapter
                        log(f"   [WARN] No Heading 2 found, using chapter only: '{full}'")

                    full = clean_title_text(full)
                    page = para.Range.Information(3)
                    headings.append({"text": full, "page": page})
                    log(f"   [TOC] Added: '{full}' -> page {page}")

                    i = j
                    continue

                page = para.Range.Information(3)
                clean_heading = clean_title_text(text)
                headings.append({"text": clean_heading, "page": page})
                log(f"   [TOC] Single-line heading: '{clean_heading}' -> page {page}")

            elif "heading 2" in style and text:
                page = para.Range.Information(3)
                clean_heading = clean_title_text(para.Range.Text)
                headings.append({"text": clean_heading, "page": page})
                log(f"   [TOC] Standalone Heading 2: '{clean_heading}' -> page {page}")

        except Exception as exc:
            log(f"   [WARN] Error processing paragraph {i}: {exc}")

        i += 1

    doc.Close(False)
    word.Quit()

    log(f"\n   [OK] Total headings extracted: {len(headings)}")
    for idx, heading in enumerate(headings, 1):
        log(f"   {idx}. {heading['text']} -> page {heading['page']}")

    return headings


def read_html_safely(path: Path) -> str:
    for encoding in ("utf-8", "cp1252", "latin-1"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="ignore")


def chapter_label(text: str) -> str:
    cleaned = clean_title_text(text)
    if ":" in cleaned:
        return cleaned.split(":", 1)[0].strip()
    return cleaned


def chapter_title(text: str) -> str:
    return clean_title_text(text)


def split_heading_display_parts(full_title: str) -> tuple[str, str]:
    """Return display label/subtitle from a heading row title.

    Examples:
    - "Chapter 1: Pre His Story" -> ("Chapter 1", "Pre His Story")
    - "INTRODUCTION" -> ("Introduction", "")
    - "RICHARD NILES DISCOGRAPHY BY YEAR" -> (same text, "")
    """
    cleaned = clean_title_text(full_title)
    if not cleaned:
        return "", ""

    if cleaned.lower() == "introduction":
        return "Introduction", ""

    match = re.match(r"^chapter\s+(\d+)\s*(?::\s*(.*))?$", cleaned, flags=re.IGNORECASE)
    if not match:
        return cleaned, ""

    number = match.group(1)
    subtitle = clean_title_text(match.group(2) or "")
    return f"Chapter {number}", subtitle


def normalize_for_match(text: str) -> str:
    return clean_title_text(text).lower()


def find_matching_heading(heading_rows: list[dict], h1_text: str, start_idx: int) -> tuple[str | None, int]:
    target = clean_title_text(h1_text).lower()
    idx = start_idx

    while idx < len(heading_rows):
        candidate_full = clean_title_text(heading_rows[idx].get("text", ""))
        candidate_label = chapter_label(candidate_full)
        if candidate_label.lower() == target:
            return chapter_title(candidate_full), idx + 1
        idx += 1

    return None, start_idx


def anchor_headings_in_full_body(html_content: str, heading_rows: list[dict]) -> tuple[str, list[dict]]:
    """Inject anchor ids in body HTML based on detected headings from Word COM."""
    soup = BeautifulSoup(html_content, HTML_PARSER)

    candidates = soup.find_all(["h1", "h2", "p"])
    normalized_candidates = []
    for tag in candidates:
        text = clean_title_text(tag.get_text(" ", strip=True))
        if text:
            normalized_candidates.append((tag, normalize_for_match(text)))

    links: list[dict] = []
    used_ids: set[str] = set()
    cursor = 0

    for idx, row in enumerate(heading_rows, start=1):
        target_title = clean_title_text(row.get("text", ""))
        if not target_title:
            continue

        target_norm = normalize_for_match(target_title)
        target_label_norm = normalize_for_match(chapter_label(target_title))
        match_tag = None

        for j in range(cursor, len(normalized_candidates)):
            tag, cand_norm = normalized_candidates[j]
            if cand_norm == target_norm or cand_norm == target_label_norm:
                match_tag = tag
                cursor = j + 1
                break

        if match_tag is None:
            for j in range(0, cursor):
                tag, cand_norm = normalized_candidates[j]
                if cand_norm == target_norm or cand_norm == target_label_norm:
                    match_tag = tag
                    break

        if match_tag is None:
            continue

        anchor_id = sanitize_id(target_title)
        if not anchor_id or anchor_id in used_ids:
            anchor_id = f"heading-{idx:03d}"
        used_ids.add(anchor_id)

        match_tag["id"] = anchor_id
        links.append({"title": target_title, "id": anchor_id})

    body = soup.body if soup.body else soup
    return str(body), links


def mark_discography_and_bibliography_sections(html_content: str) -> str:
    """Tag discography/bibliography headings and tables for section-specific styling."""
    soup = BeautifulSoup(html_content, HTML_PARSER)
    body = soup.body if soup.body else soup

    active_section = ""

    for elem in body.find_all(["h1", "h2", "table"]):
        if elem.name in {"h1", "h2"}:
            title = normalize_for_match(elem.get_text(" ", strip=True))
            full_title = elem.get_text(" ", strip=True)
            anchor_id = sanitize_id(full_title)

            if "discography" in title or "discografia" in title:
                active_section = "discography"
                # Wrap heading in section opener div with ID for TOC linking
                wrapper = soup.new_tag("div", attrs={
                    "class": "section-opener section-opener-discography",
                    "id": anchor_id
                })
                label = soup.new_tag("p", attrs={"class": "section-label"})
                label.string = "Discography"
                heading = soup.new_tag("p", attrs={"class": "section-title"})
                heading.string = full_title
                wrapper.append(label)
                wrapper.append(heading)
                elem.replace_with(wrapper)
            elif "books by richard niles" in title or "bibliography" in title or "bibliografia" in title:
                active_section = "bibliography"
                # Wrap heading in section opener div with ID for TOC linking
                wrapper = soup.new_tag("div", attrs={
                    "class": "section-opener section-opener-bibliography",
                    "id": anchor_id
                })
                label = soup.new_tag("p", attrs={"class": "section-label"})
                label.string = "Bibliography"
                heading = soup.new_tag("p", attrs={"class": "section-title"})
                heading.string = full_title
                wrapper.append(label)
                wrapper.append(heading)
                elem.replace_with(wrapper)
            else:
                active_section = ""
            continue

        if elem.name == "table" and active_section:
            existing = elem.get("class", [])
            if active_section == "discography":
                elem["class"] = existing + ["discography-table"]
            elif active_section == "bibliography":
                elem["class"] = existing + ["bibliography-table"]

    return str(body)


def parse_body_html_to_chapters(html_path: Path, heading_rows: list[dict]) -> list[dict]:
    log("\n[PARSE] Building chapter blocks from body HTML")

    html = read_html_safely(html_path)
    soup = BeautifulSoup(html, HTML_PARSER)
    body = soup.body if soup.body else soup

    # Word exported HTML can be deeply nested; use recursive search.
    elements = body.find_all(["h1", "h2", "p", "table", "ul", "ol", "blockquote", "pre"])

    chapters: list[dict] = []
    current: dict | None = None
    heading_idx = 0

    for elem in elements:
        tag = elem.name.lower()

        if tag == "h1":
            if current:
                chapters.append(current)

            matched_title, heading_idx = find_matching_heading(
                heading_rows,
                elem.get_text(strip=True),
                heading_idx,
            )
            title = matched_title or elem.get_text(strip=True) or f"Chapter {len(chapters) + 1}"
            current = {"title": title, "elements": [elem], "subtitles": []}
            log(f"   [CHAPTER] {title}")
            continue

        if current is None:
            current = {"title": "Introduction", "elements": [], "subtitles": []}

        if tag == "h2":
            subtitle = elem.get_text(strip=True) or "Section"
            current["subtitles"].append({"title": subtitle, "id": ""})

        current["elements"].append(elem)

    if current:
        chapters.append(current)

    if not chapters:
        raise RuntimeError("No chapter content found in body HTML")

    for chapter in chapters:
        wrapper = soup.new_tag("div")
        for item in chapter["elements"]:
            wrapper.append(item)
        chapter["html"] = str(wrapper)

    log(f"   [OK] Total chapter blocks: {len(chapters)}")
    return chapters


def sanitize_id(value: str) -> str:
    normalized = re.sub(r"[^a-zA-Z0-9]+", "-", value.strip().lower()).strip("-")
    return normalized or "section"


def guess_media_type(path: Path) -> str:
    ext = path.suffix.lower()
    if ext == ".png":
        return "image/png"
    if ext in {".jpg", ".jpeg"}:
        return "image/jpeg"
    if ext == ".gif":
        return "image/gif"
    if ext == ".webp":
        return "image/webp"
    if ext == ".svg":
        return "image/svg+xml"
    return "application/octet-stream"


def rewrite_html_images_and_collect_assets(
    html_content: str,
    html_base_dir: Path,
    image_registry: dict[Path, str],
    next_image_index: int,
    file_location: str = "text/body.xhtml",
) -> tuple[str, int]:
    """Rewrite <img src> to EPUB-local paths and register image assets.
    
    Args:
        file_location: Where the HTML file will be stored (e.g., "text/body.xhtml")
                      Used to calculate correct relative paths to images/
    """
    soup = BeautifulSoup(html_content, HTML_PARSER)

    for img_tag in soup.find_all("img"):
        src = str(img_tag.get("src", "")).strip()
        if not src:
            continue

        lowered = src.lower()
        if lowered.startswith(("http://", "https://", "data:", "file://")):
            continue

        normalized_src = src.split("?", 1)[0].split("#", 1)[0].replace("\\", "/")
        source_path = (html_base_dir / normalized_src).resolve()
        if not source_path.exists():
            continue

        if source_path not in image_registry:
            epub_file_name = f"images/embedded/{next_image_index:04d}_{source_path.name}"
            image_registry[source_path] = epub_file_name
            next_image_index += 1

        # Calculate relative path from file_location to image
        # If file is in text/, images are in ../images/
        if "text/" in file_location:
            relative_img_path = f"../{image_registry[source_path]}"
        else:
            relative_img_path = image_registry[source_path]
        
        img_tag["src"] = relative_img_path

    return str(soup), next_image_index


def add_embedded_images_to_book(book: epub.EpubBook, image_registry: dict[Path, str]) -> None:
    if not image_registry:
        return

    log(f"   [IMAGES] Embedding {len(image_registry)} image assets into EPUB")

    for source_path, epub_file_name in image_registry.items():
        image_item = epub.EpubImage()
        image_item.file_name = epub_file_name
        image_item.media_type = guess_media_type(source_path)
        image_item.content = source_path.read_bytes()
        book.add_item(image_item)


def add_cover_page(book: epub.EpubBook, language: str, cover_path: Path, spine: list, toc: list) -> None:
    """Add front cover as first page (not just metadata cover)."""
    image_item = epub.EpubImage()
    image_item.file_name = f"images/cover{cover_path.suffix.lower()}"
    image_item.media_type = guess_media_type(cover_path)
    image_item.content = cover_path.read_bytes()
    book.add_item(image_item)

    page = epub.EpubHtml(title="Cover", file_name="text/cover.xhtml", lang=language)
    page.add_link(href="../styles/style.css", rel="stylesheet", type="text/css")
    page.content = f'<div style="margin: 0; padding: 0; display: flex; align-items: center; justify-content: center; width: 100%; height: 100vh;"><img src="../{image_item.file_name}" alt="Cover" style="display: block; width: 100%; height: 100%; object-fit: contain; margin: 0;"/></div>'
    book.add_item(page)
    spine.append(page)
    toc.append(page)
    log(f"   [COVER] Added cover page from {cover_path}")


def add_title_page(
    book: epub.EpubBook,
    language: str,
    cfg: dict,
    spine: list,
    toc: list,
    title_page_html: str = "",
) -> None:
    """Add title page, preferring the first real front-matter page when available."""
    page = epub.EpubHtml(title="Title Page", file_name="text/title.xhtml", lang=language)
    page.add_link(href="../styles/style.css", rel="stylesheet", type="text/css")

    if title_page_html.strip():
        page.content = title_page_html
        log("   [TITLE] Added title page from first front-matter page")
    else:
        title = escape(str(cfg.get("title", "Untitled")))
        author = escape(str(cfg.get("author", "Unknown Author")))
        publisher = escape(str(cfg.get("publisher", "")))
        year = escape(str(cfg.get("publication_year", "")))
        publisher_line = f"<p style='font-size: 0.9em; color: #666; margin-top: 1.5em; margin-bottom: 0;'>{publisher}</p>" if publisher else ""
        year_line = f"<p style='font-size: 0.9em; color: #999; margin-top: 0.3em;'>{year}</p>" if year else ""
        page.content = f'<div style="text-align: center; padding: 1.5em; display: flex; flex-direction: column; justify-content: center; min-height: 100vh;"><div><h1 style="font-size: 2em; font-weight: bold; margin: 0.5em 0;">{title}</h1><p style="font-size: 1.3em; color: #333; margin: 0.8em 0 0.3em 0;">by</p><p style="font-size: 1.3em; font-weight: 600; margin: 0;">{author}</p>{publisher_line}{year_line}</div></div>'
        log(f"   [TITLE] Added fallback title page: {title} by {author}")

    book.add_item(page)
    spine.append(page)
    toc.append(page)


def add_copyright_page(
    book: epub.EpubBook,
    language: str,
    cfg: dict,
    spine: list,
    toc: list,
    copyright_body_html: str = "",
) -> None:
    """Add copyright and publication information page."""
    page = epub.EpubHtml(title="Copyright", file_name="text/copyright.xhtml", lang=language)
    page.add_link(href="../styles/style.css", rel="stylesheet", type="text/css")

    if copyright_body_html.strip():
        page.content = copyright_body_html
        log("   [COPYRIGHT] Added copyright page from source DOCX")
    else:
        title = escape(str(cfg.get("title", "")))
        author = escape(str(cfg.get("author", "")))
        publisher = escape(str(cfg.get("publisher", "")))
        year = escape(str(cfg.get("publication_year", "")))
        isbn = escape(str(cfg.get("isbn", "")))
        copyright_holder = escape(str(cfg.get("copyright_holder", author)))

        isbn_line = f"<p>ISBN: {isbn}</p>" if isbn else ""
        page.content = f'<div style="font-size: 0.9em; color: #666; padding: 2em; margin-top: 4em;"><h2 style="font-size: 1.2em; margin-bottom: 1.5em;">Publication Information</h2><p><strong>{title}</strong></p><p>© {year} {copyright_holder}</p>{isbn_line}<p style="margin-top: 2em; font-style: italic; border-top: 1px solid #ddd; padding-top: 1em;">All rights reserved. No part of this book may be reproduced in any form or by any electronic or mechanical means, including information storage and retrieval systems, without permission in writing from the author, except by a reviewer who may quote brief passages in a review.</p><p style="margin-top: 1em;">Published by {publisher}</p></div>'
        log("   [COPYRIGHT] Added fallback copyright page")

    book.add_item(page)
    spine.append(page)
    toc.append(page)


def extract_body_fragment(html_content: str) -> str:
    """Return inner body fragment if present; otherwise return parsed content."""
    soup = BeautifulSoup(html_content, HTML_PARSER)
    body = soup.body if soup.body else soup
    return body.decode_contents().strip()


def extract_first_front_matter_page_fragment(html_content: str) -> str:
    """Return the first front-matter page/body fragment from Word filtered HTML."""
    soup = BeautifulSoup(html_content, HTML_PARSER)
    body = soup.body if soup.body else soup

    first_section = body.find("div", class_=re.compile(r"\bWordSection\d+\b"))
    if first_section is not None:
        return str(first_section)

    fragment_parts: list[str] = []
    for child in body.children:
        child_html = str(child)
        if "page-break-before:always" in child_html:
            break
        fragment_parts.append(child_html)

    return "".join(fragment_parts).strip() or body.decode_contents().strip()


def compact_title_page_fragment(html_content: str) -> str:
    """Tighten title-page spacing so the imported logo stays on the same page."""
    soup = BeautifulSoup(html_content, HTML_PARSER)
    root = soup.find("div", class_=re.compile(r"\bWordSection\d+\b")) or soup

    def is_blank_paragraph(tag) -> bool:
        if getattr(tag, "name", None) != "p":
            return False
        if tag.find("img"):
            return False
        text_value = tag.get_text(" ", strip=True).replace("\xa0", "").strip()
        return not text_value

    for paragraph in list(root.find_all("p")):
        if is_blank_paragraph(paragraph):
            paragraph.decompose()

    for image in root.find_all("img"):
        image.attrs.pop("width", None)
        image.attrs.pop("height", None)
        existing_style = str(image.get("style", "")).strip().rstrip(";")
        compact_style = "max-width: 220px; width: 42%; height: auto; display: block; margin: 0.75em auto 0"
        image["style"] = f"{existing_style}; {compact_style}" if existing_style else compact_style

    return str(root)


def build_visual_toc_content(heading_links: list[dict], chapter_entries: list[dict] | None = None) -> str:
    chapter_href_by_id = {}
    if chapter_entries:
        chapter_href_by_id = {
            str(entry.get("id", "")): f'{posixpath.relpath(entry["file"], "text")}#{entry["id"]}'
            for entry in chapter_entries
            if entry.get("id") and entry.get("file")
        }

    toc_items = ""
    for idx, entry in enumerate(heading_links, 1):
        indent = "&nbsp;&nbsp;&nbsp;&nbsp;" if entry.get("level", 1) > 1 else ""
        entry_id = entry.get("id", f"entry-{idx}")
        title = escape(str(entry.get("title", f"Entry {idx}")))
        href = chapter_href_by_id.get(entry_id, f"#{entry_id}")
        toc_items += f'<p style="margin: 0.5em 0;">{indent}<a href="{href}">{title}</a></p>\n'

    return f'<h1>Table of Contents</h1><div style="margin-left: 1em;">{toc_items}</div>'


def add_toc_page(book: epub.EpubBook, language: str, heading_links: list[dict], spine: list, toc: list) -> epub.EpubHtml | None:
    """Add table of contents as a visual page in the book."""
    if not heading_links:
        log("   [TOC] No headings found, skipping TOC page")
        return None

    page = epub.EpubHtml(title="Table of Contents", file_name="text/toc.xhtml", lang=language)
    page.add_link(href="../styles/style.css", rel="stylesheet", type="text/css")
    page.content = build_visual_toc_content(heading_links)
    book.add_item(page)
    spine.append(page)
    toc.append(page)
    log(f"   [TOC] Added visual table of contents with {len(heading_links)} entries")
    return page


def add_chapter_opener_pages(book: epub.EpubBook, language: str, heading_links: list[dict], spine: list) -> None:
    """Add full-page chapter opener pages before the main body."""
    if not heading_links:
        return
    
    log(f"\n[CHAPTERS] Adding {len(heading_links)} chapter opener pages")
    
    for idx, entry in enumerate(heading_links, 1):
        full_title = clean_title_text(str(entry.get("title", f"Chapter {idx}")))
        display_label, display_subtitle = split_heading_display_parts(full_title)

        label_html = f'<p style="font-size: 1.15em; color: #444; font-weight: 600; margin: 0 0 0.18em 0; line-height: 1.12;">{escape(display_label)}</p>' if display_label else ""
        subtitle_source = display_subtitle or ("" if display_label else full_title)
        subtitle_html = f'<p style="font-size: 2.5em; color: #111; font-weight: 700; margin: 0; line-height: 1.08;">{escape(subtitle_source)}</p>' if subtitle_source else ""
        
        page = epub.EpubHtml(
            title=full_title,
            file_name=f"text/chapter_{idx:03d}_opener.xhtml",
            lang=language
        )
        page.add_link(href="../styles/style.css", rel="stylesheet", type="text/css")
        page.content = f'<div style="display: flex; flex-direction: column; align-items: center; justify-content: center; min-height: 70vh; margin: 0; padding: 0.7em 1em 0.5em 1em; text-align: center; page-break-after: always;">{label_html}{subtitle_html}</div>'
        book.add_item(page)
        spine.insert(spine.index(spine[-1]) if spine else 0, page)


def partition_html_by_chapters(html_content: str, heading_links: list[dict]) -> list[dict]:
    """Partition HTML into chapter chunks using heading ID positions."""
    body_html = extract_body_fragment(html_content)
    heading_ids = [entry["id"] for entry in heading_links if entry.get("id")]
    id_to_title = {
        entry["id"]: clean_title_text(str(entry.get("title", "")))
        for entry in heading_links
        if entry.get("id")
    }

    starts: list[tuple[str, int]] = []
    for hid in heading_ids:
        pos = body_html.find(f'id="{hid}"')
        if pos < 0:
            pos = body_html.find(f"id='{hid}'")
        if pos < 0:
            continue
        tag_start = body_html.rfind("<", 0, pos)
        if tag_start < 0:
            tag_start = pos
        starts.append((hid, tag_start))

    starts.sort(key=lambda item: item[1])

    chapters: list[dict] = []
    for idx, (hid, start_pos) in enumerate(starts):
        end_pos = starts[idx + 1][1] if idx + 1 < len(starts) else len(body_html)
        chunk_html = body_html[start_pos:end_pos].strip()
        if not chunk_html:
            continue
        chapters.append(
            {
                "title": id_to_title.get(hid, hid),
                "heading_id": hid,
                "content": [chunk_html],
            }
        )

    log(f"   [PARTITION] Split into {len(chapters)} chapters")
    return chapters


def create_chapter_files(
    book: epub.EpubBook,
    chapters: list[dict],
    heading_links: list[dict],
    language: str,
    image_registry: dict[Path, str],
    next_image_index: int,
    body_html_path: Path,
    chapter_output_dir: Path,
    spine: list,
    toc: list
) -> tuple[int, list[dict]]:
    """Create individual chapter files and add to book."""
    chapter_entries = []
    heading_id_order = [entry["id"] for entry in heading_links if entry.get("id")]
    id_to_title = {
        entry["id"]: clean_title_text(str(entry.get("title", entry["id"])))
        for entry in heading_links
        if entry.get("id")
    }
    chapter_counter = 0

    def split_chunk_by_heading_ids(chunk_html: str, ids_in_order: list[str]) -> list[tuple[str, str]]:
        """Split one chunk into sub-chunks when multiple heading IDs exist in it."""
        present: list[tuple[str, int]] = []
        for hid in ids_in_order:
            pos = chunk_html.find(f'id="{hid}"')
            if pos < 0:
                pos = chunk_html.find(f"id='{hid}'")
            if pos < 0:
                continue
            tag_start = chunk_html.rfind("<", 0, pos)
            if tag_start < 0:
                tag_start = pos
            present.append((hid, tag_start))

        if not present:
            return []

        if len(present) == 1:
            return [(present[0][0], chunk_html)]

        present.sort(key=lambda item: item[1])
        fragments: list[tuple[str, str]] = []
        for idx, (hid, start_pos) in enumerate(present):
            end_pos = present[idx + 1][1] if idx + 1 < len(present) else len(chunk_html)
            fragment = chunk_html[start_pos:end_pos].strip()
            if fragment:
                fragments.append((hid, fragment))
        return fragments

    def clean_chapter_heading_html(html: str, chapter_num: int, chapter_title: str) -> str:
        """Replace complex heading markup with clean semantic headings."""
        soup = BeautifulSoup(html, HTML_PARSER)
        first_heading = soup.find(["h1", "h2", "h3", "div"], id=True)
        if first_heading:
            display_label, display_subtitle = split_heading_display_parts(chapter_title)
            if display_subtitle:
                replacement_html = (
                    f'<h1 class="chapter-title">{escape(display_label)}</h1>'
                    f'<h2 class="chapter-subtitle">{escape(display_subtitle)}</h2>'
                )
            else:
                # Keep non-chapter sections (e.g., Introduction) unnumbered.
                replacement_html = f'<h1 class="chapter-title">{escape(display_label or chapter_title)}</h1>'
            first_heading.replace_with(
                BeautifulSoup(replacement_html, HTML_PARSER)
            )
        return str(soup)

    def sanitize_chapter_fragment_html(fragment_html: str, language_code: str) -> str:
        """Clean Word-specific markup while preserving textual content and image refs."""
        soup = BeautifulSoup(fragment_html, HTML_PARSER)

        # Remove duplicate title fragments that remain after replacing the original
        # chapter heading block with a generated h1/h2 pair.
        first_h2 = soup.find("h2", class_="chapter-subtitle")
        if first_h2:
            subtitle_text = first_h2.get_text(" ", strip=True)
            sibling = first_h2.next_sibling
            while sibling is not None:
                next_sibling = sibling.next_sibling

                if isinstance(sibling, str):
                    if sibling.strip():
                        break
                    sibling = next_sibling
                    continue

                sibling_text = sibling.get_text(" ", strip=True)
                is_heading_like = sibling.name in {"h1", "h2", "h3", "p", "div"}
                if (
                    is_heading_like
                    and sibling_text
                    and (sibling_text in subtitle_text or subtitle_text.endswith(sibling_text))
                ):
                    sibling.decompose()
                    sibling = next_sibling
                    continue

                break

        for tag in soup.find_all(True):
            css_classes = tag.get("class") or []
            if css_classes:
                cleaned = [c for c in css_classes if not c.lower().startswith("mso")]
                if cleaned:
                    tag["class"] = cleaned
                else:
                    tag.attrs.pop("class", None)

            # Drop deprecated align attributes; prefer stylesheet classes.
            if "align" in tag.attrs:
                align_value = str(tag.attrs.pop("align", "")).strip().lower()
                if align_value in {"center", "left", "right"}:
                    tag_classes = tag.get("class") or []
                    tag_classes.append(f"align-{align_value}")
                    tag["class"] = sorted(set(tag_classes))

            # Remove deprecated image spacing attributes.
            for legacy_attr in ("hspace", "vspace"):
                tag.attrs.pop(legacy_attr, None)

            # Remove deprecated table presentational attributes.
            for legacy_attr in ("valign", "width", "height", "cellpadding", "cellspacing", "border", "clear"):
                tag.attrs.pop(legacy_attr, None)

            # Strip all inline styles; rely on stylesheet classes for presentation.
            tag.attrs.pop("style", None)

            # Keep explicit language only when it differs from document language.
            if "lang" in tag.attrs and str(tag["lang"]).lower() in {"en-us", language_code.lower()}:
                tag.attrs.pop("lang", None)

            # Ensure image alt exists and remove generated disclaimer text.
            if tag.name == "img":
                alt_text = str(tag.get("alt", "")).replace("AI-generated content may be incorrect.", "").strip()
                tag["alt"] = alt_text
                for dim_attr in ("width", "height"):
                    if dim_attr in tag.attrs:
                        try:
                            int(str(tag[dim_attr]))
                        except Exception:
                            tag.attrs.pop(dim_attr, None)

        # Remove empty spans without attributes, preserving text.
        for span in soup.find_all("span"):
            if not span.attrs:
                span.unwrap()

        # Remove empty paragraphs used as layout artifacts.
        for para in soup.find_all("p"):
            text_value = para.get_text(" ", strip=True).replace("\xa0", "").strip()
            if not text_value and not para.find("img"):
                para.decompose()

        return str(soup)

    def build_full_chapter_xhtml(fragment_html: str, chapter_num: int, chapter_title: str, language_code: str) -> str:
        """Wrap cleaned chapter fragment as complete EPUB XHTML document."""
        safe_title = clean_title_text(chapter_title) or f"Chapter {chapter_num}"
        return (
            '<?xml version="1.0" encoding="UTF-8"?>\n'
            "<!DOCTYPE html>\n"
            f'<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="{escape(language_code)}" lang="{escape(language_code)}">\n'
            "<head>\n"
            f"  <title>{escape(safe_title)}</title>\n"
            '  <meta charset="UTF-8"/>\n'
            '  <link rel="stylesheet" type="text/css" href="../styles/style.css"/>\n'
            "</head>\n"
            "<body>\n"
            f"{fragment_html}\n"
            "</body>\n"
            "</html>\n"
        )
    
    for chapter in chapters:
        # Create HTML for this chapter
        chapter_html_parts = []
        for element in chapter["content"]:
            chapter_html_parts.append(str(element))
        
        chapter_html = "\n".join(chapter_html_parts)

        matched_ids = [hid for hid in heading_id_order if f'id="{hid}"' in chapter_html or f"id='{hid}'" in chapter_html]
        split_fragments = split_chunk_by_heading_ids(chapter_html, matched_ids)
        if not split_fragments:
            split_fragments = [(chapter["heading_id"], chapter_html)]

        for fragment_heading_id, fragment_html in split_fragments:
            chapter_counter += 1
            chapter_title_text = id_to_title.get(fragment_heading_id, chapter.get("title", f"Chapter {chapter_counter}"))
         
            # Clean chapter heading markup
            fragment_html = clean_chapter_heading_html(fragment_html, chapter_counter, chapter_title_text)

            # Remove Word-export artifacts while preserving content and image references.
            fragment_html = sanitize_chapter_fragment_html(fragment_html, language)
            
            # Rewrite images for this chapter
            fragment_html, next_image_index = rewrite_html_images_and_collect_assets(
                fragment_html,
                body_html_path.parent,
                image_registry,
                next_image_index,
                file_location=f"text/chapter_{chapter_counter:03d}.xhtml",
            )
            
            # Persist chapter XHTML file on disk first, then load it for EPUB packaging.
            chapter_file_name = f"text/chapter_{chapter_counter:03d}.xhtml"
            chapter_disk_path = chapter_output_dir / f"chapter_{chapter_counter:03d}.xhtml"
            chapter_disk_path.parent.mkdir(parents=True, exist_ok=True)
            chapter_xhtml = build_full_chapter_xhtml(fragment_html, chapter_counter, chapter_title_text, language)
            chapter_disk_path.write_text(chapter_xhtml, encoding="utf-8")

            chapter_item = epub.EpubHtml(
                title=chapter_title_text,
                file_name=chapter_file_name,
                lang=language
            )
            chapter_item.add_link(href="../styles/style.css", rel="stylesheet", type="text/css")
            chapter_xhtml_from_disk = chapter_disk_path.read_text(encoding="utf-8")
            chapter_body_html = extract_body_fragment(chapter_xhtml_from_disk)
            chapter_item.content = chapter_body_html or fragment_html
            book.add_item(chapter_item)
            spine.append(chapter_item)
            
            # Add to TOC
            chapter_entries.append({
                "file": chapter_file_name,
                "title": chapter_title_text,
                "id": fragment_heading_id,
                "index": chapter_counter,
                "disk_path": str(chapter_disk_path)
            })
            
            log(f"   [CHAPTER] {chapter_counter:2d}. {chapter_title_text} ({len(fragment_html)} chars)")
    
    return next_image_index, chapter_entries


def create_epub(
    cfg: dict,
    front_html_path: Path | None,
    body_html_path: Path,
    heading_rows: list[dict],
    copyright_html_path: Path | None = None,
) -> Path:
    log("\n[EPUB] Building final EPUB package")

    book = epub.EpubBook()
    book.FOLDER_NAME = EPUB_PACKAGE_ROOT
    book.set_identifier(cfg["identifier"])
    book.set_title(cfg["title"])
    book.set_language(cfg["language"])
    book.add_author(cfg["author"])

    if cfg.get("cover_image_file"):
        cover = cfg["cover_image_file"]
        with cover.open("rb") as handle:
            book.set_cover(f"cover{cover.suffix.lower()}", handle.read())

    spine: list = ["nav"]
    toc: list = []
    image_registry: dict[Path, str] = {}
    next_image_index = 1
    title_page_html = ""
    copyright_body_html = ""

    if front_html_path and front_html_path.exists():
        front_html_full = read_html_safely(front_html_path)
        first_front_page_html = extract_first_front_matter_page_fragment(front_html_full)
        if first_front_page_html.strip():
            first_front_page_html, next_image_index = rewrite_html_images_and_collect_assets(
                first_front_page_html,
                front_html_path.parent,
                image_registry,
                next_image_index,
                file_location="text/title.xhtml",
            )
            title_page_html = extract_body_fragment(compact_title_page_fragment(first_front_page_html))

    if copyright_html_path and copyright_html_path.exists():
        copyright_html_full = read_html_safely(copyright_html_path)
        copyright_html_full, next_image_index = rewrite_html_images_and_collect_assets(
            copyright_html_full,
            copyright_html_path.parent,
            image_registry,
            next_image_index,
            file_location="text/copyright.xhtml",
        )
        copyright_body_html = extract_body_fragment(copyright_html_full)

    # Add professional front matter pages
    if cfg.get("cover_image_file"):
        add_cover_page(book, cfg["language"], cfg["cover_image_file"], spine, toc)
    
    add_title_page(book, cfg["language"], cfg, spine, toc, title_page_html)
    add_copyright_page(book, cfg["language"], cfg, spine, toc, copyright_body_html)

    # Process book body and extract headings
    body_html_full = read_html_safely(body_html_path)
    body_html_full, heading_links = anchor_headings_in_full_body(body_html_full, heading_rows)
    body_html_full = mark_discography_and_bibliography_sections(body_html_full)
    body_html_full, next_image_index = rewrite_html_images_and_collect_assets(
        body_html_full,
        body_html_path.parent,
        image_registry,
        next_image_index,
        file_location="text/body.xhtml",
    )

    # Add visual TOC page before main content
    toc_page = None
    if heading_links:
        toc_page = add_toc_page(book, cfg["language"], heading_links, spine, toc)
    
    # Add chapter opener pages
    if heading_links:
        add_chapter_opener_pages(book, cfg["language"], heading_links, spine)

    # Partition HTML by chapters and create chapter files
    chapters = partition_html_by_chapters(body_html_full, heading_links)
    chapter_output_dir = cfg["output_epub"].parent / "text"
    next_image_index, chapter_entries = create_chapter_files(
        book,
        chapters,
        heading_links,
        cfg["language"],
        image_registry,
        next_image_index,
        body_html_path,
        chapter_output_dir,
        spine,
        toc
    )
    
    # Build TOC with chapter links
    if chapter_entries:
        toc_links = tuple(
            epub.Link(entry["file"], entry["title"], entry["id"])
            for entry in chapter_entries
        )
        toc.append((epub.Section("Book Content"), toc_links))
        log(f"   [TOC] Added {len(chapter_entries)} chapter links")
        if toc_page is not None:
            toc_page.content = build_visual_toc_content(heading_links, chapter_entries)

    if cfg.get("back_cover_image_file"):
        back_cover = cfg["back_cover_image_file"]
        log(f"   [BACK] Using back cover image: {back_cover}")
        # Back cover can be handled as a simple image page if needed
        # add_back_cover_page(book, cfg["language"], back_cover, spine, toc)

    add_embedded_images_to_book(book, image_registry)

    # Add modern professional CSS stylesheet
    nav_css = epub.EpubItem(
        uid="style_nav",
        file_name="styles/style.css",
        media_type="text/css",
        content=get_modern_epub_css(),
    )
    book.add_item(nav_css)

    book.toc = tuple(toc)
    book.add_item(epub.EpubNcx())
    book.add_item(epub.EpubNav())
    book.spine = spine

    for item in book.get_items_of_type(ITEM_DOCUMENT):
        if isinstance(item, epub.EpubNav):
            continue
        body = item.get_body_content()
        log(f"   [DOC] {item.file_name}: {len(body or '')} chars")

    output_epub = cfg["output_epub"]
    output_epub.parent.mkdir(parents=True, exist_ok=True)
    epub.write_epub(str(output_epub), book, {})

    log(f"   [OK] EPUB written: {output_epub}")
    return output_epub


def maybe_cleanup_html_exports(cfg: dict, html_paths: list[Path]) -> None:
    if cfg.get("keep_html_exports", True):
        return

    log("\n[CLEANUP] Removing temp HTML exports")
    for path in html_paths:
        if path.exists():
            path.unlink(missing_ok=True)
            log(f"   [OK] Removed: {path}")


def get_modern_epub_css() -> str:
    """Generate professional modern CSS for EPUB with typography and layout."""
    return """\
/* ============================================================================
   MODERN EBOOK STYLING
   Professional typography, spacing, and visual hierarchy for EPUB/MOBI/PDF
   ========================================================================== */

/* === ROOT & BASE === */
html {
    font-size: 100%;
}

body {
    font-family: Georgia, "Times New Roman", serif;
    font-size: 1em;
    line-height: 1.6;
    margin: 0;
    padding: 1em 1.5em;
    color: #222;
    background: white;
    text-align: left;
}

/* === TYPOGRAPHY - HEADINGS === */
h1 {
    font-family: "Trebuchet MS", "Helvetica Neue", sans-serif;
    font-size: 2em;
    font-weight: bold;
    text-align: center;
    margin: 2em 0 1.5em 0;
    padding: 1em 0;
    color: #1a1a1a;
    page-break-before: always;
    page-break-after: avoid;
}

h2 {
    font-family: "Trebuchet MS", "Helvetica Neue", sans-serif;
    font-size: 1.5em;
    font-weight: bold;
    margin: 1.5em 0 1em 0;
    padding: 0.5em 0 0.5em 0;
    color: #333;
    border-bottom: 2px solid #007ACC;
    page-break-after: avoid;
}

h3 {
    font-family: "Trebuchet MS", "Helvetica Neue", sans-serif;
    font-size: 1.2em;
    font-weight: bold;
    margin: 1.2em 0 0.8em 0;
    color: #444;
    page-break-after: avoid;
}

h4, h5, h6 {
    font-family: "Trebuchet MS", "Helvetica Neue", sans-serif;
    font-weight: bold;
    margin: 1em 0 0.5em 0;
    page-break-after: avoid;
}

/* === TYPOGRAPHY - PARAGRAPHS & TEXT === */
p {
    margin: 1em 0;
    text-indent: 0;
    text-align: justify;
}

p:first-of-type {
    margin-top: 0;
}

em, i {
    font-style: italic;
}

strong, b {
    font-weight: bold;
}

/* === LISTS === */
ul, ol {
    margin: 1em 0;
    padding-left: 2em;
}

li {
    margin: 0.5em 0;
    line-height: 1.6;
}

/* === BLOCKQUOTES & CALLOUTS === */
blockquote {
    margin: 1.5em 0;
    padding: 1em 1em 1em 1.5em;
    border-left: 4px solid #007ACC;
    background: #f5f5f5;
    font-style: italic;
    color: #333;
    page-break-inside: avoid;
}

blockquote p {
    margin: 0.5em 0;
    text-indent: 0;
    text-align: left;
}

/* === IMAGES & MEDIA === */
img {
    display: block;
    margin: 1.5em auto;
    max-width: 100%;
    height: auto;
    text-align: center;
    page-break-inside: avoid;
}

figure {
    margin: 1.5em 0;
    page-break-inside: avoid;
}

figcaption {
    font-size: 0.9em;
    font-style: italic;
    text-align: center;
    margin-top: 0.5em;
    color: #666;
}

/* === TABLES === */
table {
    border-collapse: collapse;
    width: 100%;
    max-width: 100%;
    table-layout: fixed;
    margin: 1.5em 0;
    font-size: 0.74em;
    line-height: 1.25;
    page-break-inside: avoid;
}

table caption {
    font-weight: bold;
    margin-bottom: 0.5em;
    text-align: left;
}

thead {
    background: #f0f0f0;
    font-weight: bold;
}

th, td {
    border: 1px solid #ddd;
    padding: 0.22em 0.32em;
    text-align: left;
    vertical-align: top;
    white-space: normal;
    word-break: break-word;
    overflow-wrap: anywhere;
}

th {
    font-weight: bold;
    color: #1a1a1a;
    background: #efefef;
}

table p,
table span,
table div,
table a,
table font {
    font-size: 0.74em !important;
    line-height: 1.3 !important;
    max-width: 100% !important;
    white-space: normal !important;
}

table td p,
table td span,
table td div,
table td a,
table td font,
table th p,
table th span,
table th div,
table th a,
table th font {
    margin: 0 !important;
}

/* === DISCOGRAPHY & BIBLIOGRAPHY === */
.section-opener {
    display: flex;
    flex-direction: column;
    justify-content: center;
    align-items: center;
    min-height: 70vh;
    page-break-before: always;
    margin: 2em 0 1.5em 0;
    padding: 1em 0;
}

.section-label {
    font-size: 1em;
    font-weight: normal;
    color: #999;
    margin: 0 0 0.5em 0;
    text-align: center;
    letter-spacing: 0.1em;
}

.section-title {
    font-size: 2.2em;
    font-weight: bold;
    text-align: center;
    color: #1a1a1a;
    margin: 0;
    line-height: 1.3;
}

table.discography-table,
table.bibliography-table {
    font-size: 0.9em;
    line-height: 1.32;
    margin: 1em 0 1.4em 0;
}

table.discography-table th,
table.discography-table td,
table.bibliography-table th,
table.bibliography-table td {
    padding: 0.14em 0.2em;
}

table.discography-table td:first-child,
table.bibliography-table td:first-child {
    width: 12%;
    font-weight: bold;
}

table.discography-table p,
table.discography-table span,
table.discography-table div,
table.discography-table a,
table.discography-table font,
table.bibliography-table p,
table.bibliography-table span,
table.bibliography-table div,
table.bibliography-table a,
table.bibliography-table font {
    font-size: 0.9em !important;
    line-height: 1.32 !important;
}

/* === HORIZONTAL RULE === */
hr {
    margin: 2em 0;
    border: none;
    border-top: 1px solid #ccc;
    height: 0;
}

/* === LINKS === */
a {
    color: #007ACC;
    text-decoration: none;
}

a:visited {
    color: #7030A0;
}

a:hover, a:active {
    text-decoration: underline;
}

/* === SPECIAL SECTIONS === */
.foreword, .introduction, .preface {
    margin: 2em 0;
    padding: 1.5em;
    border: 1px solid #ddd;
    background: #fafafa;
    page-break-inside: avoid;
}

.author-note {
    font-style: italic;
    color: #666;
    padding: 1em;
    border-left: 3px solid #999;
    margin: 1em 0;
    page-break-inside: avoid;
}

.quote {
    font-style: italic;
    margin: 1em 2em;
    color: #555;
    text-align: center;
}

/* === SPACING & BREAKS === */
.page-break {
    page-break-after: always;
}

.no-indent {
    text-indent: 0;
}

/* === PSEUDO-ELEMENTS === */
::first-letter {
    font-size: 1.1em;
}

/* === PRINT-SPECIFIC RULES === */
@media print {
    body {
        margin: 0;
        padding: 1em;
    }
    h1, h2, h3 {
        page-break-after: avoid;
        page-break-inside: avoid;
    }
    img, figure {
        page-break-inside: avoid;
    }
    a {
        color: inherit;
        text-decoration: none;
    }
}

/* === REFLOWABLE EBOOK DEFAULTS === */
@supports (display: flex) {
    body {
        display: block;
    }
}
"""


# ================================================================================
# VALIDATION FUNCTIONS FOR --verify FLAG
# ================================================================================


class ValidationReport:
    """Container for EPUB validation results."""
    
    def __init__(self, epub_path: Path):
        self.epub_path = epub_path
        self.errors: list[str] = []
        self.warnings: list[str] = []
        self.info: list[str] = []
        self.stats: dict = {}
        self.character_analysis: dict = {}
    
    def add_error(self, msg: str) -> None:
        self.errors.append(msg)
    
    def add_warning(self, msg: str) -> None:
        self.warnings.append(msg)
    
    def add_info(self, msg: str) -> None:
        self.info.append(msg)
    
    def is_valid(self) -> bool:
        return len(self.errors) == 0
    
    def severity_score(self) -> int:
        """Lower is better. 0 = perfect."""
        return len(self.errors) * 100 + len(self.warnings) * 10


def analyze_character_encoding(content: str) -> dict:
    """Scan content for character encoding issues and mojibake risk."""
    result = {
        "total_chars": len(content),
        "ascii_only": True,
        "high_unicode": [],
        "control_chars": [],
        "problematic_ranges": {},
        "encoding_risk": "low",
    }
    
    control_count = 0
    high_count = 0
    problematic_found = False
    
    for char in set(content):
        code = ord(char)
        
        if code < 32 and char not in ("\n", "\r", "\t"):
            control_count += 1
            result["control_chars"].append(f"U+{code:04X}")
        
        elif code > 127:
            result["ascii_only"] = False
            high_count += 1
            
            # Common problem ranges
            if 0x2018 <= code <= 0x201F:  # Smart quotes
                problematic_found = True
                result["problematic_ranges"]["smart_quotes"] = result["problematic_ranges"].get("smart_quotes", 0) + 1
            elif 0x2010 <= code <= 0x2015:  # Dashes
                problematic_found = True
                result["problematic_ranges"]["dashes"] = result["problematic_ranges"].get("dashes", 0) + 1
            elif 0x00A0 <= code <= 0x00AD:  # Special spaces
                problematic_found = True
                result["problematic_ranges"]["special_spaces"] = result["problematic_ranges"].get("special_spaces", 0) + 1
    
    if control_count > 5:
        result["encoding_risk"] = "high"
    elif problematic_found or high_count > 100:
        result["encoding_risk"] = "medium"
    
    result["high_unicode_count"] = high_count
    result["control_char_count"] = control_count
    return result


def validate_character_encoding(epub_path: Path, report: ValidationReport) -> None:
    """Validate character encoding in EPUB content."""
    log("\n[VERIFY] Checking character encoding...")
    
    try:
        with zipfile.ZipFile(epub_path, "r") as zf:
            for name in zf.namelist():
                if name.endswith((".xhtml", ".html", ".xml")):
                    try:
                        content = zf.read(name).decode("utf-8", errors="replace")
                        analysis = analyze_character_encoding(content)
                        report.character_analysis[name] = analysis
                        
                        if analysis["encoding_risk"] == "high":
                            report.add_warning(
                                f"Character encoding risk in {name}: "
                                f"{analysis['control_char_count']} control chars detected"
                            )
                        elif analysis["encoding_risk"] == "medium":
                            if analysis["problematic_ranges"]:
                                types = ", ".join(analysis["problematic_ranges"].keys())
                                report.add_info(
                                    f"Potentially problematic characters in {name}: {types} "
                                    f"(KDP may need character normalization)"
                                )
                    except Exception as e:
                        report.add_warning(f"Could not analyze {name}: {e}")
    except Exception as e:
        report.add_error(f"Failed to analyze character encoding: {e}")


def validate_html_structure(epub_path: Path, report: ValidationReport) -> None:
    """Validate HTML structure for KDP compliance."""
    log("[VERIFY] Validating HTML structure...")
    
    try:
        with zipfile.ZipFile(epub_path, "r") as zf:
            body_names = get_manuscript_xhtml_files(set(zf.namelist()))
            
            if not body_names:
                report.add_error("No manuscript XHTML files found in EPUB")
                return
            
            for body_name in body_names:
                content = zf.read(body_name).decode("utf-8", errors="replace")
                soup = BeautifulSoup(content, HTML_PARSER)
                
                # Check for Heading 1 structure
                h1_count = len(soup.find_all("h1"))
                if h1_count == 0:
                    report.add_warning("No <h1> (Heading 1) tags found - KDP expects chapter structure")
                else:
                    report.add_info(f"Found {h1_count} <h1> heading tags in {body_name}")
                
                # Check for problematic Word styling
                styled_spans = soup.find_all("span", {"style": re.compile(".*")})
                if len(styled_spans) > 100:
                    report.add_warning(
                        f"Found {len(styled_spans)} inline-styled spans - "
                        "KDP prefers semantic markup, consider cleaning styles"
                    )
                
                # Check for headers/footers (should not exist)
                headers = soup.find_all(["header", "footer"])
                if headers:
                    report.add_error(
                        f"Found {len(headers)} header/footer tags - "
                        "KDP rejects fixed headers/footers (eBooks are reflowable)"
                    )
                
                # Check for page breaks (should not exist)
                page_breaks = soup.find_all(["pb", "pagebreak"])
                if page_breaks:
                    report.add_warning(
                        f"Found {len(page_breaks)} page break tags - "
                        "KDP converts to reflowable format (page breaks will be ignored)"
                    )
                
                report.stats[body_name] = {
                    "h1_count": h1_count,
                    "span_count": len(styled_spans),
                    "images": len(soup.find_all("img")),
                }
    
    except Exception as e:
        report.add_error(f"Failed to validate HTML structure: {e}")


def detect_epub_package_root(all_files: set[str]) -> str:
    for candidate in (EPUB_PACKAGE_ROOT, "EPUB"):
        prefix = f"{candidate}/"
        if any(name.startswith(prefix) for name in all_files):
            return candidate
    return EPUB_PACKAGE_ROOT


def get_manuscript_xhtml_files(all_files: set[str]) -> list[str]:
    chapter_files = sorted(
        name
        for name in all_files
        if name.endswith(".xhtml") and "/text/chapter_" in name and "_opener.xhtml" not in name
    )
    if chapter_files:
        return chapter_files

    body_files = sorted(name for name in all_files if "body" in name.lower() and name.endswith(".xhtml"))
    if body_files:
        return body_files

    return sorted(
        name
        for name in all_files
        if name.endswith(".xhtml") and "/text/" in name and not name.endswith(("nav.xhtml", "toc.xhtml"))
    )


def validate_images(epub_path: Path, report: ValidationReport) -> None:
    """Validate that images are embedded and references are valid."""
    log("[VERIFY] Validating image references...")
    
    try:
        with zipfile.ZipFile(epub_path, "r") as zf:
            all_files = set(zf.namelist())
            package_root = detect_epub_package_root(all_files)
            image_prefix = f"{package_root}/images/"
            image_files = {
                f
                for f in all_files
                if f.startswith(image_prefix) and f.split(".")[-1].lower() in {"jpg", "jpeg", "png", "gif", "webp"}
            }
            
            report.stats["embedded_images"] = len(image_files)
            
            # Check all HTML/XHTML files for image references
            broken_refs = []
            found_refs = 0
            
            for name in all_files:
                if name.endswith((".xhtml", ".html")):
                    try:
                        content = zf.read(name).decode("utf-8", errors="replace")
                        img_srcs = re.findall(r'<img[^>]+src=["\']([^"\']+)["\']', content)
                        found_refs += len(img_srcs)
                        
                        for src in img_srcs:
                            # Remove fragment
                            src_clean = src.split("#", 1)[0].split("?", 1)[0]
                            
                            # Skip URLs
                            if src_clean.startswith(("http://", "https://", "data:")):
                                continue
                            
                            # Resolve relative path from the current XHTML file location.
                            current_dir = posixpath.dirname(name)
                            candidate_paths = {
                                src_clean,
                                posixpath.normpath(posixpath.join(current_dir, src_clean)),
                            }

                            # Also try with/without the package root prefix for compatibility.
                            expanded_candidates = set(candidate_paths)
                            for candidate in candidate_paths:
                                if candidate.startswith(f"{package_root}/"):
                                    expanded_candidates.add(candidate[len(package_root) + 1:])
                                else:
                                    expanded_candidates.add(f"{package_root}/{candidate}")

                            file_found = any(candidate in all_files for candidate in expanded_candidates)
                            
                            if not file_found:
                                broken_refs.append(f"{src_clean} (referenced in {name})")
                    except Exception as e:
                        report.add_warning(f"Could not check images in {name}: {e}")
            
            if image_files:
                report.add_info(f"✓ Found {len(image_files)} embedded images")
            else:
                report.add_warning("No embedded images found (may be intentional)")
            
            if broken_refs:
                for ref in broken_refs:
                    report.add_error(f"Broken image reference: {ref}")
            else:
                if found_refs > 0:
                    report.add_info(f"✓ All {found_refs} image references are valid")
    
    except Exception as e:
        report.add_error(f"Failed to validate images: {e}")


def validate_kdp_compliance(epub_path: Path, report: ValidationReport) -> None:
    """Check KDP-specific requirements."""
    log("[VERIFY] Checking KDP compliance...")
    
    try:
        with zipfile.ZipFile(epub_path, "r") as zf:
            # Check package.opf for required metadata
            opf_files = [f for f in zf.namelist() if f.endswith(".opf")]
            if opf_files:
                opf_content = zf.read(opf_files[0]).decode("utf-8", errors="replace")
                
                # Basic KDP requirements
                if not re.search(r"<dc:title\b", opf_content) and "<title>" not in opf_content:
                    report.add_error("Missing title metadata in OPF")
                
                if not re.search(r"<dc:creator\b", opf_content) and "<author>" not in opf_content:
                    report.add_warning("No author metadata found")
            
            # Check for reflowable design (CSS should exist)
            css_files = [f for f in zf.namelist() if f.endswith(".css")]
            if not css_files:
                report.add_warning("No CSS files found - ensure responsive styling for reflowable design")
            
            # Check TOC
            nav_files = [f for f in zf.namelist() if "nav" in f.lower() and f.endswith(".xhtml")]
            if nav_files:
                nav_content = zf.read(nav_files[0]).decode("utf-8", errors="replace")
                nav_items = re.findall(r"<li>.*?</li>", nav_content, re.DOTALL)
                report.stats["toc_entries"] = len(nav_items)
                report.add_info(f"✓ TOC has {len(nav_items)} entries")
            else:
                report.add_warning("No navigation file found - KDP requires TOC")
            
            report.add_info("✓ KDP basic structure check complete")
    
    except Exception as e:
        report.add_error(f"Failed to validate KDP compliance: {e}")


def validate_content_completeness(epub_path: Path, report: ValidationReport) -> None:
    """Verify manuscript is complete (all chapters present)."""
    log("[VERIFY] Checking content completeness...")
    
    try:
        with zipfile.ZipFile(epub_path, "r") as zf:
            body_files = get_manuscript_xhtml_files(set(zf.namelist()))

            if not body_files:
                report.add_error("No manuscript content files found for completeness check")
                return
            
            total_size = 0
            chapter_count = 0
            
            for body_file in body_files:
                content = zf.read(body_file).decode("utf-8", errors="replace")
                
                # Count chapters
                chapter_count += len(re.findall(r"<h1[^>]*>", content))
                total_size += len(content)
                
                # Check for specific chapter markers
                for i in range(1, 50):
                    if re.search(f"chapter\\s+{i}", content, re.IGNORECASE):
                        continue  # Just counting
                    else:
                        break
                
                # Warning if suspiciously short
            if total_size < 100000:
                report.add_warning(f"Body content is small ({total_size} chars) - verify all chapters included")
            else:
                report.add_info(f"✓ Body content size: {total_size:,} chars")
            
            report.stats["chapters"] = chapter_count
            if chapter_count > 0:
                report.add_info(f"✓ Found {chapter_count} chapters")
            else:
                report.add_warning("No chapter structure detected")
    
    except Exception as e:
        report.add_error(f"Failed to validate content completeness: {e}")


def generate_verification_report(epub_path: Path, output_report_path: Path) -> ValidationReport:
    """Run all validations and generate comprehensive report."""
    
    if not epub_path.exists():
        raise FileNotFoundError(f"EPUB not found: {epub_path}")
    
    log(f"\n{'='*72}")
    log("EBOOK VERIFICATION START")
    log(f"{'='*72}")
    log(f"EPUB: {epub_path}")
    log(f"Size: {epub_path.stat().st_size:,} bytes")
    
    report = ValidationReport(epub_path)
    
    # Run all validations
    validate_character_encoding(epub_path, report)
    validate_html_structure(epub_path, report)
    validate_images(epub_path, report)
    validate_kdp_compliance(epub_path, report)
    validate_content_completeness(epub_path, report)
    
    # Generate text report
    report_lines = [
        "=" * 80,
        "EBOOK VERIFICATION REPORT",
        "=" * 80,
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"EPUB File: {epub_path.name}",
        f"File Size: {epub_path.stat().st_size:,} bytes",
        "",
    ]
    
    # Summary
    if report.is_valid():
        report_lines.append("STATUS: ✓ PASS - No critical issues found")
    else:
        report_lines.append(f"STATUS: ✗ FAIL - {len(report.errors)} critical error(s) found")
    
    report_lines.extend([
        "",
        f"SUMMARY:",
        f"  Errors:    {len(report.errors)}",
        f"  Warnings:  {len(report.warnings)}",
        f"  Info:      {len(report.info)}",
        f"  Severity Score: {report.severity_score()}/100",
        "",
    ])
    
    # Errors
    if report.errors:
        report_lines.extend([
            "CRITICAL ERRORS:",
            "─" * 80,
        ])
        for idx, err in enumerate(report.errors, 1):
            report_lines.append(f"  [{idx}] {err}")
        report_lines.append("")
    
    # Warnings
    if report.warnings:
        report_lines.extend([
            "WARNINGS:",
            "─" * 80,
        ])
        for idx, warn in enumerate(report.warnings, 1):
            report_lines.append(f"  [{idx}] {warn}")
        report_lines.append("")
    
    # Info
    if report.info:
        report_lines.extend([
            "INFORMATION:",
            "─" * 80,
        ])
        for idx, inf in enumerate(report.info, 1):
            report_lines.append(f"  [{idx}] {inf}")
        report_lines.append("")
    
    # Statistics
    if report.stats:
        report_lines.extend([
            "STATISTICS:",
            "─" * 80,
        ])
        for key, value in report.stats.items():
            if isinstance(value, dict):
                report_lines.append(f"  {key}:")
                for k, v in value.items():
                    report_lines.append(f"    {k}: {v}")
            else:
                report_lines.append(f"  {key}: {value}")
        report_lines.append("")
    
    # Character Analysis
    if report.character_analysis:
        report_lines.extend([
            "CHARACTER ENCODING ANALYSIS:",
            "─" * 80,
        ])
        for file_name, analysis in report.character_analysis.items():
            risk = analysis.get("encoding_risk", "unknown").upper()
            report_lines.append(f"  {file_name}:")
            report_lines.append(f"    Encoding Risk: {risk}")
            report_lines.append(f"    ASCII Only: {analysis.get('ascii_only', False)}")
            if analysis.get("problematic_ranges"):
                types = ", ".join(analysis["problematic_ranges"].keys())
                report_lines.append(f"    Problematic Characters: {types}")
        report_lines.append("")
    
    # Recommendations
    report_lines.extend([
        "RECOMMENDATIONS FOR KDP UPLOAD:",
        "─" * 80,
        "  1. Fix all CRITICAL ERRORS before uploading",
        "  2. Address WARNINGS related to formatting and character encoding",
        "  3. Consider running Kindle Previewer for final validation",
        "  4. Test in multiple Kindle apps/devices if possible",
        "  5. For character encoding issues, consider:",
        "     - Using 'conservative' mode to preserve typographic characters",
        "     - Or using 'kdp-safe' mode for ASCII-only compatibility",
        "",
    ])
    
    report_lines.extend([
        "=" * 80,
        f"KDP READY: {'YES' if report.is_valid() else 'NO (fix errors first)'}",
        "=" * 80,
    ])
    
    # Write report
    report_text = "\n".join(report_lines)
    output_report_path.parent.mkdir(parents=True, exist_ok=True)
    output_report_path.write_text(report_text, encoding="utf-8")
    
    # Print to console
    print(report_text)
    
    log(f"\nReport saved to: {output_report_path}")
    log("=" * 72)
    
    return report


def verify_epub(epub_path: Path | None = None) -> int:
    """Verify an existing EPUB file."""
    
    if epub_path is None:
        project_root = find_project_root()
        cfg = load_builder_config(project_root)
        epub_path = cfg["output_epub"]
    
    if not epub_path.exists():
        log(f"ERROR: EPUB file not found: {epub_path}")
        return 1
    
    report_path = epub_path.parent / f"{epub_path.stem}_verification_report.txt"
    report = generate_verification_report(epub_path, report_path)
    
    return 0 if report.is_valid() else 1


def find_kindlegen_executable(configured_path: str = "") -> Path:
    """Locate kindlegen from config, PATH, or Kindle Previewer install."""
    candidates: list[Path] = []

    if configured_path:
        candidates.append(Path(configured_path))

    env_path = os.environ.get("KINDLEGEN_PATH", "").strip()
    if env_path:
        candidates.append(Path(env_path))

    on_path = shutil.which("kindlegen")
    if on_path:
        candidates.append(Path(on_path))

    # Common Kindle Previewer bundled kindlegen locations.
    candidates.extend([
        Path(r"C:\Users\Hugo\AppData\Local\Amazon\Kindle Previewer 3\lib\fc\bin\kindlegen.exe"),
        Path(r"C:\Program Files\Amazon\Kindle Previewer 3\lib\fc\bin\kindlegen.exe"),
        Path(r"C:\Program Files (x86)\Amazon\Kindle Previewer 3\lib\fc\bin\kindlegen.exe"),
    ])

    for candidate in candidates:
        if candidate.exists():
            return candidate

    raise FileNotFoundError(
        "kindlegen executable not found. Install Kindle Previewer 3 or set options.kindlegen_executable in config."
    )


def convert_epub_to_mobi(epub_path: Path, mobi_path: Path, configured_kindlegen: str = "") -> Path:
    """Convert an EPUB file into a Kindle MOBI file."""
    if not epub_path.exists():
        raise FileNotFoundError(f"EPUB not found for MOBI conversion: {epub_path}")

    kindlegen_exe = find_kindlegen_executable(configured_kindlegen)

    log("\n[MOBI] Converting EPUB -> MOBI")
    log(f"   Input EPUB: {epub_path}")
    log(f"   Output MOBI: {mobi_path}")
    log(f"   Converter: {kindlegen_exe}")

    mobi_path.parent.mkdir(parents=True, exist_ok=True)

    # kindlegen writes output to working directory; pass output file name only.
    result = subprocess.run(
        [str(kindlegen_exe), str(epub_path), "-o", mobi_path.name],
        cwd=str(mobi_path.parent),
        capture_output=True,
        text=True,
        check=False,
    )

    if result.returncode != 0:
        combined = ((result.stdout or "") + "\n" + (result.stderr or "")).strip()
        raise RuntimeError(f"kindlegen failed with exit code {result.returncode}\n{combined}")

    if not mobi_path.exists():
        raise RuntimeError(f"kindlegen completed but MOBI file was not created: {mobi_path}")

    log(f"   [OK] MOBI written: {mobi_path}")
    return mobi_path


def build_mobi(epub_path: Path | None = None) -> Path:
    """Build MOBI using config defaults or a specified EPUB path."""
    project_root = find_project_root()
    cfg = load_builder_config(project_root)
    source_epub = epub_path or cfg["output_epub"]

    return convert_epub_to_mobi(
        source_epub,
        cfg["output_mobi"],
        cfg.get("kindlegen_executable", ""),
    )


def build_epub() -> Path:
    log("=" * 72)
    log("EBOOK BUILDER START")
    log("=" * 72)

    project_root = find_project_root()
    cfg = load_builder_config(project_root)

    log(f"   Config: {cfg['config_path']}")
    log(f"   Body DOCX: {cfg['book_body_file']}")
    log(f"   Output EPUB: {cfg['output_epub']}")

    # Preflight cleanup before Word COM work.
    log("\n[WORD] Preflight cleanup")
    kill_running_word_instances()

    # Word can lock or keep stale state between documents; clear between conversions.
    kill_running_word_instances()
    front_html: Path | None = None
    if cfg.get("front_matter_docx"):
        front_html = export_docx_to_filtered_html(cfg["front_matter_docx"], cfg["temp_dir"])

    kill_running_word_instances()
    body_html = export_docx_to_filtered_html(cfg["book_body_file"], cfg["temp_dir"])

    # Ensure heading extraction starts from a clean Word state.
    kill_running_word_instances()
    headings = extract_headings(cfg["book_body_file"])
    log(f"\n[HEADINGS] Count: {len(headings)}")

    copyright_html: Path | None = None
    if cfg.get("copyright_page_docx"):
        kill_running_word_instances()
        copyright_html = export_docx_to_filtered_html(cfg["copyright_page_docx"], cfg["temp_dir"])

    # KDP reflowable guidance: preserve complete manuscript flow and heading navigation.
    output_epub = create_epub(cfg, front_html, body_html, headings, copyright_html)
    cleanup_targets = [body_html]
    if front_html:
        cleanup_targets.append(front_html)
    if copyright_html:
        cleanup_targets.append(copyright_html)
    maybe_cleanup_html_exports(cfg, cleanup_targets)

    log("=" * 72)
    log("EBOOK BUILDER DONE")
    log(f"Generated: {output_epub}")
    log("=" * 72)
    return output_epub


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="EPUB eBook builder with KDP compatibility verification"
    )
    parser.add_argument(
        "--build",
        action="store_true",
        help="Build EPUB from DOCX inputs (default action)",
    )
    parser.add_argument(
        "--verify",
        metavar="EPUB_PATH",
        nargs="?",
        const="auto",
        help="Verify EPUB file for KDP compatibility. Use 'auto' or omit to verify output_epub from config",
    )
    parser.add_argument(
        "--mobi",
        action="store_true",
        help="After building EPUB, also generate MOBI using kindlegen",
    )
    parser.add_argument(
        "--mobi-from-epub",
        metavar="EPUB_PATH",
        nargs="?",
        const="auto",
        help="Convert an existing EPUB to MOBI. Use 'auto' or omit to use output_epub from config",
    )
    
    args = parser.parse_args()
    
    try:
        if args.verify is not None:
            # Verification mode
            if args.verify == "auto":
                exit_code = verify_epub(None)
            else:
                epub_to_verify = Path(args.verify)
                exit_code = verify_epub(epub_to_verify)
            sys.exit(exit_code)
        elif args.mobi_from_epub is not None:
            if args.mobi_from_epub == "auto":
                output_mobi = build_mobi(None)
            else:
                output_mobi = build_mobi(Path(args.mobi_from_epub))
            log(f"Generated: {output_mobi}")
        else:
            # Default: build mode
            output_epub = build_epub()
            if args.mobi:
                output_mobi = build_mobi(output_epub)
                log(f"Generated: {output_mobi}")
    except Exception as exc:
        log(f"\nERROR: {exc}")
        sys.exit(1)
