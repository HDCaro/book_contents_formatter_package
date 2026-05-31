#!/usr/bin/env python3
"""Audit paperback image inventory and exported image sizes.

This script builds a page-indexed list of images from the paperback DOCX and
compares it against the exported image files used by the ebook pipeline.

Outputs:
- CSV: one row per detected paperback image with page and export match details
- JSON: summary counts, coverage, and sizing issues
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from pathlib import Path
from typing import Any

try:
    from PIL import Image
except Exception as exc:  # pragma: no cover
    raise RuntimeError("Pillow is required for exported image size checks") from exc

try:
    import win32com.client
except Exception as exc:  # pragma: no cover
    raise RuntimeError("pywin32 is required (win32com.client)") from exc

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.utils.book_project import get_active_book_root, resolve_book_path

WORD_PAGE_INFO_CODE = 3
PICTURE_SHAPE_TYPES = {11, 13}
EXPORT_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".tif", ".tiff"}


def log(message: str) -> None:
    print(message, flush=True)


def natural_sort_key(value: str) -> tuple:
    parts = re.split(r"(\d+)", value.lower())
    key: list[Any] = []
    for part in parts:
        if part.isdigit():
            key.append(int(part))
        else:
            key.append(part)
    return tuple(key)


def read_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def find_project_root() -> Path:
    current = Path(__file__).resolve()
    for parent in [current.parent] + list(current.parents):
        if (parent / "src").exists():
            return parent
    raise RuntimeError("Project root not found (expected folder: src)")


def resolve_from_config(config_path: Path) -> tuple[Path, Path, Path]:
    cfg = read_json(config_path)
    book_root = get_active_book_root()

    inputs = cfg.get("inputs", {})
    outputs = cfg.get("outputs", {})

    full_book_docx = inputs.get("full_book_docx")
    temp_dir = outputs.get("temp_dir")
    output_dir = outputs.get("output_dir")

    if not full_book_docx:
        raise RuntimeError("Config is missing inputs.full_book_docx")
    if not temp_dir:
        raise RuntimeError("Config is missing outputs.temp_dir")
    if not output_dir:
        raise RuntimeError("Config is missing outputs.output_dir")

    docx_path = resolve_book_path(book_root, full_book_docx)
    temp_path = resolve_book_path(book_root, temp_dir)
    out_path = resolve_book_path(book_root, output_dir)
    return docx_path, temp_path, out_path


def create_word_app():
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    word.ScreenUpdating = False
    return word


def safe_float(value: Any) -> float:
    try:
        return float(value)
    except Exception:
        return 0.0


def points_to_pixels(points: float, dpi: float=96.0) -> int:
    return int(round(points * dpi / 72.0))


def extract_paperback_images(docx_path: Path) -> list[dict]:
    if not docx_path.exists():
        raise FileNotFoundError(f"Paperback DOCX not found: {docx_path}")

    log(f"[DOCX] Scanning paperback images from: {docx_path}")

    word = create_word_app()
    doc = None
    records: list[dict] = []

    try:
        doc = word.Documents.Open(str(docx_path))
        doc.Repaginate()

        inline_total = int(doc.InlineShapes.Count)
        for i in range(1, inline_total + 1):
            try:
                ishape = doc.InlineShapes(i)
                page = int(ishape.Range.Information(WORD_PAGE_INFO_CODE) or 0)
                width_pt = safe_float(getattr(ishape, "Width", 0.0))
                height_pt = safe_float(getattr(ishape, "Height", 0.0))
                records.append(
                    {
                        "doc_image_index": len(records) + 1,
                        "shape_kind": "inline",
                        "shape_index": i,
                        "shape_type": int(getattr(ishape, "Type", 0) or 0),
                        "page": page,
                        "width_pt": round(width_pt, 2),
                        "height_pt": round(height_pt, 2),
                        "width_px_est": points_to_pixels(width_pt),
                        "height_px_est": points_to_pixels(height_pt),
                    }
                )
            except Exception as exc:
                log(f"   [WARN] InlineShape {i} failed: {exc}")

        floating_total = int(doc.Shapes.Count)
        for i in range(1, floating_total + 1):
            try:
                shape = doc.Shapes(i)
                shape_type = int(getattr(shape, "Type", 0) or 0)

                # Keep obvious picture shapes. Other floating shapes are usually text boxes.
                if shape_type not in PICTURE_SHAPE_TYPES:
                    continue

                anchor = getattr(shape, "Anchor", None)
                page = int(anchor.Information(WORD_PAGE_INFO_CODE) if anchor is not None else 0)
                width_pt = safe_float(getattr(shape, "Width", 0.0))
                height_pt = safe_float(getattr(shape, "Height", 0.0))
                records.append(
                    {
                        "doc_image_index": len(records) + 1,
                        "shape_kind": "floating",
                        "shape_index": i,
                        "shape_type": shape_type,
                        "page": page,
                        "width_pt": round(width_pt, 2),
                        "height_pt": round(height_pt, 2),
                        "width_px_est": points_to_pixels(width_pt),
                        "height_px_est": points_to_pixels(height_pt),
                    }
                )
            except Exception as exc:
                log(f"   [WARN] Shape {i} failed: {exc}")

    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        try:
            word.Quit()
        except Exception:
            pass

    records.sort(key=lambda row: (row["page"], row["doc_image_index"]))
    for idx, row in enumerate(records, start=1):
        row["doc_image_index"] = idx

    log(f"   [OK] Found {len(records)} paperback image shapes")
    return records


def scan_exported_images(images_dir: Path, min_width: int, min_height: int) -> list[dict]:
    if not images_dir.exists():
        raise FileNotFoundError(f"Exported images dir not found: {images_dir}")

    files = [
        p
        for p in images_dir.iterdir()
        if p.is_file() and p.suffix.lower() in EXPORT_EXTENSIONS
    ]
    files.sort(key=lambda p: natural_sort_key(p.name))

    records: list[dict] = []
    for i, path in enumerate(files, start=1):
        width = 0
        height = 0
        readable = True
        error = ""

        try:
            with Image.open(path) as img:
                width, height = img.size
        except Exception as exc:
            readable = False
            error = str(exc)

        too_small = readable and (width < min_width or height < min_height)
        records.append(
            {
                "export_index": i,
                "export_name": path.name,
                "export_path": str(path),
                "bytes": path.stat().st_size,
                "width_px": width,
                "height_px": height,
                "readable": readable,
                "too_small": too_small,
                "size_ok": readable and not too_small,
                "error": error,
            }
        )

    log(f"[EXPORT] Found {len(records)} exported image files in: {images_dir}")
    return records


def combine_records(doc_images: list[dict], exported_images: list[dict]) -> list[dict]:
    combined: list[dict] = []
    max_len = max(len(doc_images), len(exported_images))

    for idx in range(max_len):
        d = doc_images[idx] if idx < len(doc_images) else {}
        e = exported_images[idx] if idx < len(exported_images) else {}

        row = {
            "ordinal": idx + 1,
            "doc_image_index": d.get("doc_image_index", ""),
            "page": d.get("page", ""),
            "shape_kind": d.get("shape_kind", ""),
            "shape_type": d.get("shape_type", ""),
            "doc_width_pt": d.get("width_pt", ""),
            "doc_height_pt": d.get("height_pt", ""),
            "doc_width_px_est": d.get("width_px_est", ""),
            "doc_height_px_est": d.get("height_px_est", ""),
            "export_index": e.get("export_index", ""),
            "export_name": e.get("export_name", ""),
            "export_path": e.get("export_path", ""),
            "export_bytes": e.get("bytes", ""),
            "export_width_px": e.get("width_px", ""),
            "export_height_px": e.get("height_px", ""),
            "export_readable": e.get("readable", ""),
            "export_size_ok": e.get("size_ok", ""),
            "export_too_small": e.get("too_small", ""),
            "export_error": e.get("error", ""),
            "has_doc_image": bool(d),
            "has_export_image": bool(e),
            "paired": bool(d) and bool(e),
        }

        combined.append(row)

    return combined


def summarize(doc_images: list[dict], exported_images: list[dict], combined: list[dict]) -> dict:
    by_page: dict[int, int] = {}
    for row in doc_images:
        page = int(row.get("page") or 0)
        by_page[page] = by_page.get(page, 0) + 1

    unreadable = [r for r in exported_images if not r["readable"]]
    too_small = [r for r in exported_images if r["too_small"]]

    missing_export = [r for r in combined if r["has_doc_image"] and not r["has_export_image"]]
    orphan_export = [r for r in combined if r["has_export_image"] and not r["has_doc_image"]]

    return {
        "paperback_images": len(doc_images),
        "exported_images": len(exported_images),
        "paired_images": sum(1 for r in combined if r["paired"]),
        "missing_export_for_doc_image": len(missing_export),
        "orphan_export_images": len(orphan_export),
        "unreadable_export_images": len(unreadable),
        "too_small_export_images": len(too_small),
        "images_by_page": dict(sorted(by_page.items(), key=lambda kv: kv[0])),
        "missing_export_details": missing_export[:50],
        "orphan_export_details": orphan_export[:50],
        "unreadable_details": unreadable[:50],
        "too_small_details": too_small[:50],
    }


def write_csv(path: Path, rows: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [
        "ordinal",
        "doc_image_index",
        "page",
        "shape_kind",
        "shape_type",
        "doc_width_pt",
        "doc_height_pt",
        "doc_width_px_est",
        "doc_height_px_est",
        "export_index",
        "export_name",
        "export_path",
        "export_bytes",
        "export_width_px",
        "export_height_px",
        "export_readable",
        "export_size_ok",
        "export_too_small",
        "export_error",
        "has_doc_image",
        "has_export_image",
        "paired",
    ]
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def write_json(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Audit paperback images and exported image sizes")
    parser.add_argument(
        "--config",
        type=Path,
        default=Path(__file__).with_name("ebook_builder.config.json"),
        help="Path to ebook_builder.config.json",
    )
    parser.add_argument("--docx", type=Path, default=None, help="Override paperback DOCX path")
    parser.add_argument(
        "--exported-images-dir",
        type=Path,
        default=None,
        help="Override exported image directory",
    )
    parser.add_argument("--output-csv", type=Path, default=None, help="Output CSV path")
    parser.add_argument("--output-json", type=Path, default=None, help="Output JSON path")
    parser.add_argument("--min-width", type=int, default=120, help="Minimum acceptable exported width")
    parser.add_argument("--min-height", type=int, default=120, help="Minimum acceptable exported height")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    project_root = find_project_root()

    config_path = args.config
    if not config_path.is_absolute():
        config_path = project_root / config_path
    if not config_path.exists():
        raise FileNotFoundError(f"Config not found: {config_path}")

    cfg_docx, cfg_temp_dir, cfg_output_dir = resolve_from_config(config_path)

    docx_path = args.docx or cfg_docx
    exported_dir = args.exported_images_dir or (cfg_temp_dir / "ebook_source_files")

    output_csv = args.output_csv or (cfg_output_dir / "paperback_image_audit.csv")
    output_json = args.output_json or (cfg_output_dir / "paperback_image_audit.json")

    doc_images = extract_paperback_images(docx_path)
    exported_images = scan_exported_images(exported_dir, args.min_width, args.min_height)
    combined = combine_records(doc_images, exported_images)
    summary = summarize(doc_images, exported_images, combined)

    write_csv(output_csv, combined)
    write_json(output_json, summary)

    log("\n[REPORT] Done")
    log(f"   CSV:  {output_csv}")
    log(f"   JSON: {output_json}")
    log(f"   Paperback images: {summary['paperback_images']}")
    log(f"   Exported images:  {summary['exported_images']}")
    log(f"   Missing exports:  {summary['missing_export_for_doc_image']}")
    log(f"   Too small files:  {summary['too_small_export_images']}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
