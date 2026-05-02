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

        if "–" in part:
            try:
                a, b = part.split("–")
                pages.update(range(int(a), int(b) + 1))
            except:
                pass
        else:
            try:
                pages.add(int(part))
            except:
                pass

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

# ---------------- MAIN ---------------- #

def main():
    print("\n=== EXTRACT INDEX FROM DOCX ===\n")

    print(f"📥 INPUT DOCX:  {INPUT_DOCX}")
    print(f"📤 OUTPUT JSON: {OUTPUT_JSON}")
    print(f"📤 OUTPUT CSV:  {OUTPUT_CSV}\n")

    if not INPUT_DOCX.exists():
        print("❌ INPUT FILE NOT FOUND")
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
        if line.startswith("(AKA:"):
            aliases = line.replace("(AKA:", "").replace(")", "")
            current_aliases = [a.strip() for a in aliases.split(";")]
            continue

        # Page line
        if is_page_line(line):
            if current_name:
                entries[current_name] = {
                    "normalized": current_name,
                    "aliases": current_aliases,
                    "aliases_external": [],
                    "pages": expand_pages(line)
                }

            current_name = None
            current_aliases = []
            continue

        # Single-line entry
        match = re.match(r"^(.*?),\s*([0-9,\s–-]+)$", line)
        if match:
            raw_name = match.group(1).strip()
            pages = expand_pages(match.group(2))

            key = invert_name(raw_name)

            entries[key] = {
                "normalized": raw_name,
                "aliases": [],
                "aliases_external": [],
                "pages": pages
            }
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

    print(f"\n💾 JSON saved → {OUTPUT_JSON}")
    print(f"💾 CSV saved  → {OUTPUT_CSV}")
    print(f"✅ Extracted {len(entries)} entries\n")


if __name__ == "__main__":
    main()