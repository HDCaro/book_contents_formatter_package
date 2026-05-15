#!/usr/bin/env python3
"""Validate EPUB internal package structure.

Checks a generated EPUB for a standard structure:
- mimetype at archive root (first entry, uncompressed)
- META-INF/container.xml exists
- container.xml points to OEBPS/content.opf
- OEBPS/content.opf exists
- OEBPS/toc.ncx exists
- OEBPS/text/, OEBPS/styles/, OEBPS/images/ have content hints
"""

from __future__ import annotations

import argparse
import sys
import zipfile
from pathlib import Path


def find_latest_epub(project_root: Path) -> Path:
    output_dir = project_root / "data" / "outputs" / "02_epub"
    candidates = sorted(output_dir.glob("*.epub"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not candidates:
        raise FileNotFoundError(f"No EPUB files found in {output_dir}")
    return candidates[0]


def validate_epub(epub_path: Path) -> tuple[list[str], list[str]]:
    errors: list[str] = []
    warnings: list[str] = []

    if not epub_path.exists():
        return [f"EPUB not found: {epub_path}"], warnings

    with zipfile.ZipFile(epub_path, "r") as archive:
        infos = archive.infolist()
        names = [info.filename for info in infos]
        name_set = set(names)

        # mimetype should be first and uncompressed for strict compatibility.
        if not infos:
            errors.append("Archive is empty")
            return errors, warnings

        first = infos[0]
        if first.filename != "mimetype":
            errors.append("First archive entry is not 'mimetype'")
        else:
            if first.compress_type != zipfile.ZIP_STORED:
                errors.append("'mimetype' must be stored uncompressed (ZIP_STORED)")
            try:
                mimetype_value = archive.read("mimetype").decode("utf-8", errors="ignore").strip()
                if mimetype_value != "application/epub+zip":
                    errors.append("'mimetype' content is not 'application/epub+zip'")
            except KeyError:
                errors.append("Missing required root file: mimetype")

        # Required files.
        if "META-INF/container.xml" not in name_set:
            errors.append("Missing required file: META-INF/container.xml")
        if "OEBPS/content.opf" not in name_set:
            errors.append("Missing required file: OEBPS/content.opf")
        if "OEBPS/toc.ncx" not in name_set:
            warnings.append("Missing OEBPS/toc.ncx (EPUB2 compatibility may be reduced)")

        # container.xml must point to OEBPS/content.opf.
        if "META-INF/container.xml" in name_set:
            container_text = archive.read("META-INF/container.xml").decode("utf-8", errors="ignore")
            if 'full-path="OEBPS/content.opf"' not in container_text:
                errors.append("META-INF/container.xml does not point to OEBPS/content.opf")

        # Useful content checks (warnings only).
        has_text = any(name.startswith("OEBPS/text/") and name.endswith(".xhtml") for name in names)
        has_styles = any(name.startswith("OEBPS/styles/") for name in names)
        has_images = any(name.startswith("OEBPS/images/") for name in names)

        if not has_text:
            warnings.append("No chapter XHTML files found under OEBPS/text/")
        if not has_styles:
            warnings.append("No stylesheet assets found under OEBPS/styles/")
        if not has_images:
            warnings.append("No images found under OEBPS/images/")

    return errors, warnings


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Validate EPUB structure and package layout")
    parser.add_argument("--epub", type=Path, help="Path to EPUB file (defaults to latest in output folder)")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    project_root = Path(__file__).resolve().parents[2]

    epub_path = args.epub.resolve() if args.epub else find_latest_epub(project_root)
    errors, warnings = validate_epub(epub_path)

    print("=" * 72)
    print("EPUB STRUCTURE VALIDATION")
    print("=" * 72)
    print(f"File: {epub_path}")

    if errors:
        print("\nErrors:")
        for item in errors:
            print(f"  - {item}")

    if warnings:
        print("\nWarnings:")
        for item in warnings:
            print(f"  - {item}")

    if not errors and not warnings:
        print("\nStatus: OK (no structural issues found)")
    elif not errors:
        print("\nStatus: OK with warnings")
    else:
        print("\nStatus: FAILED")

    print("=" * 72)
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
