#!/usr/bin/env python3
"""
Direct DOCX → EPUB builder (no HTML intermediate)
Reads book body directly from DOCX using python-docx
"""

import json
import sys
import re
from pathlib import Path
from typing import Optional

from docx import Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
import ebooklib
from ebooklib import epub
from bs4 import BeautifulSoup
from html import escape

# Reuse utility functions from original builder
import sys
sys.path.insert(0, str(Path(__file__).parent))
from ebook_builder import (
    log, sanitize_id, guess_media_type, get_modern_epub_css,
    extract_body_fragment, clean_title_text, normalize_for_match,
    HTML_PARSER
)


def load_builder_config(config_path: Path) -> dict:
    """Load ebook builder configuration."""
    if not config_path.exists():
        log(f"[ERROR] Config not found: {config_path}")
        sys.exit(1)
    
    with open(config_path, "r", encoding="utf-8") as f:
        raw_cfg = json.load(f)
    
    # Flatten nested config
    cfg = {}
    cfg.update(raw_cfg.get("metadata", {}))
    cfg.update(raw_cfg.get("inputs", {}))
    cfg.update(raw_cfg.get("outputs", {}))
    cfg.update(raw_cfg.get("options", {}))
    
    # Resolve relative paths from repo root
    repo_root = config_path.parent.parent.parent  # src/ebook/../../..
    
    for key in ["cover_image_file", "back_cover_image_file", "book_body_file", "copyright_page_docx", "output_dir"]:
        if key in cfg and cfg[key]:
            path = Path(cfg[key])
            if not path.is_absolute():
                path = repo_root / path
            cfg[key] = str(path.resolve())
    
    return cfg


def extract_docx_content(docx_path: Path) -> tuple[list[dict], dict]:
    """
    Extract paragraphs, headings, and images from DOCX.
    Returns (content_blocks, image_registry)
    """
    log(f"\n[DOCX] Reading: {docx_path}")
    doc = Document(docx_path)
    
    content_blocks = []
    image_registry = {}
    
    for para_or_table in doc.element.body:
        if isinstance(para_or_table, CT_P):
            para = Paragraph(para_or_table, doc)
            if not para.text.strip():
                continue
            
            style_name = para.style.name if para.style else "Normal"
            
            block = {
                "type": "paragraph",
                "text": para.text,
                "style": style_name,
            }
            
            content_blocks.append(block)
        
        elif isinstance(para_or_table, CT_Tbl):
            table = Table(para_or_table, doc)
            block = {
                "type": "table",
                "rows": []
            }
            for row in table.rows:
                row_data = [cell.text for cell in row.cells]
                block["rows"].append(row_data)
            content_blocks.append(block)
    
    log(f"   [OK] Extracted {len(content_blocks)} content blocks")
    return content_blocks, image_registry


def extract_headings(content_blocks: list[dict]) -> list[dict]:
    """Extract heading information for TOC."""
    headings = []
    heading_num = 0
    
    for block in content_blocks:
        if block["type"] != "paragraph":
            continue
        
        style = block["style"]
        text = block["text"].strip()
        
        if not text:
            continue
        
        if style == "Heading 1":
            heading_num += 1
            headings.append({
                "level": 1,
                "text": text,
                "number": heading_num
            })
        elif style == "Heading 2":
            if headings:
                # Append to last Heading 1
                if "subtitle" not in headings[-1]:
                    headings[-1]["subtitle"] = text
                else:
                    headings[-1]["subtitle"] += f": {text}"
    
    log(f"\n[HEADINGS] Found {len(headings)} chapters")
    for h in headings[:5]:
        full = h["text"]
        if "subtitle" in h:
            full += f": {h['subtitle']}"
        log(f"   - {full}")
    
    return headings


def content_blocks_to_html(blocks: list[dict]) -> tuple[str, list[dict]]:
    """Convert content blocks to clean HTML, tracking headings for TOC."""
    html_parts = []
    toc_entries = []
    
    for block in blocks:
        if block["type"] == "paragraph":
            style = block["style"]
            text = escape(block["text"])
            
            if style.startswith("Heading"):
                level = 1 if style == "Heading 1" else 2
                # Generate ID for TOC linking
                heading_id = sanitize_id(text)
                html_parts.append(f"<h{level} id=\"{heading_id}\">{text}</h{level}>")
                
                # Track for TOC
                if level == 1:
                    toc_entries.append({
                        "id": heading_id,
                        "title": text,
                        "level": 1
                    })
            else:
                if text.strip():
                    html_parts.append(f"<p>{text}</p>")
        
        elif block["type"] == "table":
            html_parts.append("<table>")
            for row in block["rows"]:
                html_parts.append("<tr>")
                for cell in row:
                    html_parts.append(f"<td>{escape(cell)}</td>")
                html_parts.append("</tr>")
            html_parts.append("</table>")
    
    return "\n".join(html_parts), toc_entries


def add_embedded_images_to_book(book: epub.EpubBook, image_registry: dict[Path, str]) -> None:
    """Add extracted images to EPUB."""
    if not image_registry:
        return
    
    log(f"   [IMAGES] Embedding {len(image_registry)} image assets")
    for source_path, epub_file_name in image_registry.items():
        if not source_path.exists():
            continue
        image_item = epub.EpubImage()
        image_item.file_name = epub_file_name
        image_item.media_type = guess_media_type(source_path)
        image_item.content = source_path.read_bytes()
        book.add_item(image_item)


def build_epub_direct(cfg: dict, toc_entries: list[dict], body_html: str, image_registry: dict[Path, str]) -> Path:
    """Build EPUB directly from extracted content."""
    log("\n[EPUB] Building EPUB package")
    
    output_dir = Path(cfg.get("output_dir", "data/outputs/02_epub"))
    output_dir.mkdir(parents=True, exist_ok=True)
    
    filename = cfg.get("filename", "book.epub")
    output_epub = output_dir / filename
    
    book = epub.EpubBook()
    book.set_identifier(cfg.get("isbn", "unknown-isbn"))
    book.set_title(cfg.get("title", "Untitled"))
    book.set_language(cfg.get("language", "en"))
    book.add_author(cfg.get("author", "Unknown"))
    
    spine = []
    toc = []
    
    # Add body
    body_item = epub.EpubHtml(title="Book Body", file_name="text/body.xhtml", lang=cfg.get("language", "en"))
    body_item.add_link(href="../styles/style.css", rel="stylesheet", type="text/css")
    body_item.content = body_html
    book.add_item(body_item)
    spine.append(body_item)
    
    # Add CSS
    css_item = epub.EpubItem(
        uid="style",
        file_name="styles/style.css",
        media_type="text/css",
        content=get_modern_epub_css()
    )
    book.add_item(css_item)
    
    # Generate TOC with links
    if toc_entries:
        toc_links = tuple(
            epub.Link(f"{body_item.file_name}#{entry['id']}", entry["title"], entry["id"])
            for entry in toc_entries
        )
        toc.append((epub.Section("Book Content"), toc_links))
        log(f"   [TOC] Added {len(toc_entries)} entries")
    else:
        toc.append(body_item)
    
    # Add images
    add_embedded_images_to_book(book, image_registry)
    
    # Set metadata
    book.toc = tuple(toc)
    book.add_item(epub.EpubNcx())
    book.add_item(epub.EpubNav())
    book.spine = spine
    
    epub.write_epub(str(output_epub), book, {})
    log(f"   [OK] EPUB written: {output_epub}")
    
    return output_epub


def main():
    script_dir = Path(__file__).parent
    config_path = script_dir / "ebook_builder.config.json"
    
    log("=" * 72)
    log("EBOOK BUILDER DIRECT (python-docx, no HTML intermediate)")
    log("=" * 72)
    
    cfg = load_builder_config(config_path)
    
    body_docx = Path(cfg.get("book_body_file"))
    log(f"\n[DEBUG] Looking for: {body_docx}")
    log(f"[DEBUG] Exists: {body_docx.exists()}")
    
    if not body_docx.exists():
        log(f"[ERROR] Body DOCX not found: {body_docx}")
        sys.exit(1)
    
    # Extract content directly from DOCX
    content_blocks, image_registry = extract_docx_content(body_docx)
    
    # Extract headings for TOC
    headings = extract_headings(content_blocks)
    
    # Convert to HTML and extract TOC entries
    body_html, toc_entries = content_blocks_to_html(content_blocks)
    
    # Build EPUB
    output_epub = build_epub_direct(cfg, toc_entries, body_html, image_registry)
    
    log("\n" + "=" * 72)
    log(f"SUCCESS: {output_epub}")
    log("=" * 72)
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
