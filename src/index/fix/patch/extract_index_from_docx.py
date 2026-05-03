"""
===============================================================================
FILE: extract_index_from_docx.py
LOCATION: src/index/fix/patch/
-------------------------------------------------------------------------------
Extract entries from approved index DOCX into JSON

✔ Keeps working parser
✔ Adds normalized field
✔ Converts "Last, First" → "First Last" for key
✔ NEW: exports CSV for manual editing
===============================================================================
"""

import re
import json
import csv
from pathlib import Path
from docx import Document

# ---------------- ROOT ---------------- #

def find_project_root():
    current = Path(__file__).resolve()
    for parent in [current] + list(current.parents):
        if (parent / "src").exists() and (parent / "data").exists():
            return parent
    raise RuntimeError("Project root not found")

BASE_DIR = find_project_root()

# ---------------- PATHS ---------------- #

INPUT_DOCX = BASE_DIR / "data/index/input/index_source/HITS AND HAPPINESS FINAL 2 Format MOM Discog-index-wrong-pages.docx"
OUTPUT_JSON = BASE_DIR / "data/index/intermediate/index_curated_extracted.json"
OUTPUT_CSV  = BASE_DIR / "data/index/output/index_curated_extracted_edit.csv"

# ---------------- HELPERS ---------------- #

def expand_pages(text):
    pages = set()

    parts = re.split(r",\s*", text)
    for part in parts:
        part = part.strip()
        if not part:
            continue

        # Accept both en dash and hyphen ranges (e.g. 10–12 or 10-12)
        range_match = re.fullmatch(r"(\d+)\s*[–-]\s*(\d+)", part)
        if range_match:
            try:
                a, b = int(range_match.group(1)), int(range_match.group(2))
                if a <= b:
                    pages.update(range(a, b + 1))
                else:
                    pages.update(range(b, a + 1))
            except ValueError:
                print(f"WARNING: Could not parse page range: {part!r}")
        else:
            try:
                pages.add(int(part))
            except ValueError:
                print(f"WARNING: Could not parse page number: {part!r}")

    return sorted(pages)


def is_page_line(line):
    return bool(re.fullmatch(r"[0-9,\s–-]+", line))


def is_section_header(line):
    return len(line) == 1 or line == "0–9"


def invert_name(name):
    if "," in name:
        parts = [p.strip() for p in name.split(",")]
        return " ".join(parts[::-1])
    return name


def upsert_entry(entries, key, normalized, aliases, pages):
    existing = entries.get(key)
    if existing:
        existing["aliases"] = sorted(
            set(existing.get("aliases", [])) | set(aliases or [])
        )
        existing["pages"] = sorted(
            set(existing.get("pages", [])) | set(pages or [])
        )
        return

    entries[key] = {
        "normalized": normalized,
        "aliases": aliases or [],
        "aliases_external": [],
        "pages": pages or [],
    }

# ---------------- MAIN ---------------- #

def main():
    print("\n=== EXTRACT INDEX FROM DOCX ===\n")

    print(f"INPUT DOCX:  {INPUT_DOCX}")
    print(f"OUTPUT JSON: {OUTPUT_JSON}")
    print(f"OUTPUT CSV:  {OUTPUT_CSV}\n")

    if not INPUT_DOCX.exists():
        print("ERROR: INPUT FILE NOT FOUND")
        return

    doc = Document(INPUT_DOCX)

    lines = []
    for p in doc.paragraphs:
        for l in p.text.split("\n"):
            l = l.strip()
            if l:
                lines.append(l)

    entries = {}

    current_name = None
    current_aliases = []

    for line in lines:

        if is_section_header(line):
            continue

        # AKA
        if re.match(r"^\(\s*AKA\s*:", line, flags=re.IGNORECASE):
            aliases = re.sub(r"^\(\s*AKA\s*:\s*", "", line, flags=re.IGNORECASE)
            aliases = aliases.rstrip(")")
            current_aliases = [a.strip() for a in aliases.split(";")]
            continue

        # Page line
        if is_page_line(line):
            if current_name:
                key = invert_name(current_name)
                upsert_entry(
                    entries=entries,
                    key=key,
                    normalized=current_name,
                    aliases=current_aliases,
                    pages=expand_pages(line),
                )

            current_name = None
            current_aliases = []
            continue

        # Single-line entry
        match = re.match(r"^(.*?),\s*([0-9,\s–-]+)$", line)
        if match:
            raw_name = match.group(1).strip()
            pages = expand_pages(match.group(2))

            key = invert_name(raw_name)

            upsert_entry(
                entries=entries,
                key=key,
                normalized=raw_name,
                aliases=[],
                pages=pages,
            )
            continue

        # Multi-line name
        current_name = line
        current_aliases = []

    # ---------------- SAVE JSON ---------------- #

    OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(entries, f, indent=2)

    # ---------------- SAVE CSV ---------------- #

    OUTPUT_CSV.parent.mkdir(parents=True, exist_ok=True)

    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as csvfile:
        writer = csv.writer(csvfile)

        writer.writerow([
            "original_key",
            "edited_key",
            "normalized",
            "pages"
        ])

        for key, entry in sorted(entries.items()):
            pages_str = ", ".join(map(str, entry.get("pages", [])))

            writer.writerow([
                key,
                key,  # editable column
                entry.get("normalized"),
                pages_str
            ])

    print(f"\nJSON saved -> {OUTPUT_JSON}")
    print(f"CSV saved  -> {OUTPUT_CSV}")
    print(f"Extracted {len(entries)} entries\n")


if __name__ == "__main__":
    main()