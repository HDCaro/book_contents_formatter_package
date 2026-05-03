"""
Apply fixed index pages from CSV into the wrong-pages index DOCX.

Input:
- data/index/input/index_source/HITS AND HAPPINESS FINAL 2 Format MOM Discog-index-wrong-pages.docx
- data/index/output/index_extracted_word_fixed_report.csv (uses `word` column)

Output:
- release/v1/HITS AND HAPPINESS FINAL 2 Format MOM Discog-index-fixed.docx
"""

import csv
import re
from pathlib import Path

from docx import Document


def find_project_root():
    current = Path(__file__).resolve()
    for parent in [current] + list(current.parents):
        if (parent / "src").exists() and (parent / "data").exists():
            return parent
    raise RuntimeError("Project root not found")


BASE_DIR = find_project_root()

INPUT_DOCX = BASE_DIR / "data/index/input/index_source/HITS AND HAPPINESS FINAL 2 Format MOM Discog-index-wrong-pages.docx"
INPUT_CSV = BASE_DIR / "data/index/output/index_extracted_word_fixed_report.csv"
OUTPUT_DOCX = BASE_DIR / "release/v1/HITS AND HAPPINESS FINAL 2 Format MOM Discog-index-fixed.docx"


def normalize(text):
    text = (text or "").lower().strip()
    text = re.sub(r"^the\s+", "", text)
    text = re.sub(r"[^\w\s]", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def invert_name(name):
    if "," in name:
        parts = [p.strip() for p in name.split(",") if p.strip()]
        if len(parts) >= 2:
            return " ".join(parts[1:] + [parts[0]])
    return name


def is_page_line(line):
    return bool(re.fullmatch(r"[0-9,\s–\-]+", (line or "").strip()))


def build_page_lookup(csv_path):
    lookup = {}
    with open(csv_path, "r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            key = (row.get("key") or "").strip()
            word_pages = (row.get("word") or "").strip()
            if not key or not word_pages:
                continue
            lookup[normalize(key)] = word_pages
    return lookup


def resolve_pages(name, lookup):
    candidates = [name, invert_name(name)]
    for candidate in candidates:
        found = lookup.get(normalize(candidate))
        if found:
            return found
    return None


def replace_line_in_paragraph(paragraph, old_line, new_line):
    """
    Replace a line inside a paragraph while preserving run formatting when possible.
    Falls back to paragraph text replacement if run-local replacement is not possible.
    """
    for run in paragraph.runs:
        if old_line in run.text:
            run.text = run.text.replace(old_line, new_line, 1)
            return True

    if old_line in paragraph.text:
        paragraph.text = paragraph.text.replace(old_line, new_line, 1)
        return True

    return False


def main():
    print("\n=== APPLY WORD FIXED PAGES TO INDEX DOCX ===\n")
    print(f"INPUT DOCX:  {INPUT_DOCX}")
    print(f"INPUT CSV:   {INPUT_CSV}")
    print(f"OUTPUT DOCX: {OUTPUT_DOCX}\n")

    if not INPUT_DOCX.exists():
        raise FileNotFoundError(f"Missing input DOCX: {INPUT_DOCX}")
    if not INPUT_CSV.exists():
        raise FileNotFoundError(f"Missing input CSV: {INPUT_CSV}")

    page_lookup = build_page_lookup(INPUT_CSV)
    doc = Document(INPUT_DOCX)

    current_name = None
    replaced = 0
    missing = 0
    unresolved_names = set()

    for paragraph in doc.paragraphs:
        if not paragraph.text.strip():
            continue

        lines = paragraph.text.split("\n")
        for line in lines:
            stripped = line.strip()
            if not stripped:
                continue

            # Single-line entry: "Name, 1, 2-3"
            single = re.match(r"^(.*?),\s*([0-9,\s–\-]+)$", stripped)
            if single:
                name = single.group(1).strip()
                old_pages = single.group(2).strip()
                new_pages = resolve_pages(name, page_lookup)
                if new_pages and old_pages != new_pages:
                    new_line = f"{name}, {new_pages}"
                    if replace_line_in_paragraph(paragraph, stripped, new_line):
                        replaced += 1
                elif not new_pages:
                    missing += 1
                    unresolved_names.add(name)
                current_name = None
                continue

            # Multi-line pages entry: page line following name line
            if is_page_line(stripped):
                if current_name:
                    new_pages = resolve_pages(current_name, page_lookup)
                    if new_pages and stripped != new_pages:
                        if replace_line_in_paragraph(paragraph, stripped, new_pages):
                            replaced += 1
                    elif not new_pages:
                        missing += 1
                        unresolved_names.add(current_name)
                current_name = None
                continue

            # Name line (for possible following pages line)
            current_name = stripped

    OUTPUT_DOCX.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUTPUT_DOCX)

    print(f"Replacements applied: {replaced}")
    print(f"Entries not resolved: {missing}")
    if unresolved_names:
        sample = sorted(unresolved_names)[:15]
        print("\nSample unresolved names:")
        for name in sample:
            print(f"- {name}")

    print(f"\nSaved fixed DOCX: {OUTPUT_DOCX}\n")


if __name__ == "__main__":
    main()
