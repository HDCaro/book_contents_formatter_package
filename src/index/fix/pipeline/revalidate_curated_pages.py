"""
===============================================================================
FILE: revalidate_curated_pages.py
-------------------------------------------------------------------------------
PURPOSE
-------
Recalculate page numbers for ALL entries in index_curated_old.json
using Word COM.

This step ensures:
✔ Correct pagination
✔ Alias coverage (aliases + aliases_external)
✔ Clean separation of editorial vs data layers

INPUT
-----
index_curated_old.json

OUTPUT
------
index_curated_old_pages.json

NOTES
-----
✔ Preserves ALL editorial fields
✔ ONLY updates "pages"
✔ Does NOT depend on raw or transaction files

===============================================================================
"""

import json
import re
from pathlib import Path
import win32com.client as win32

# ---------------- CONFIG ---------------- #

BASE_DIR = Path(__file__).resolve().parents[3]

DOCX_INPUT = BASE_DIR / "data/index/input/HITS AND HAPPINESS FINAL 2 Format MOM Discog.docx"

INPUT_JSON = BASE_DIR / "data/index/intermediate/index_curated_old.json"
OUTPUT_JSON = BASE_DIR / "data/index/intermediate/index_curated_old_pages.json"

PROGRESS_EVERY = 25
MIN_PAGES_WARNING = 1  # warn if fewer pages found

# ---------------- WORD ---------------- #

def open_word(path):
    print("🟡 Opening Word...")
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(str(path))
    print("🟢 Document loaded")
    return word, doc


def close_word(word, doc):
    doc.Close(False)
    word.Quit()

# ---------------- SEARCH ---------------- #

def find_pages(doc, text):
    """
    Find all pages where a term appears.
    """
    pages = set()

    pattern = re.compile(rf"\b{re.escape(text)}\b", re.IGNORECASE)
    full_text = doc.Content.Text

    for match in pattern.finditer(full_text):
        try:
            rng = doc.Range(Start=match.start(), End=match.end())
            page = int(rng.Information(1))
            pages.add(page)
        except:
            continue

    return pages


def collect_pages_for_entry(doc, name, entry):
    """
    Collect pages using:
    - main name
    - aliases
    - external aliases
    """
    search_terms = [name]

    search_terms += entry.get("aliases", [])
    search_terms += entry.get("aliases_external", [])

    pages = set()

    for term in search_terms:
        pages.update(find_pages(doc, term))

    return sorted(pages)

# ---------------- MAIN ---------------- #

def main():
    print("\n=== REVALIDATE CURATED PAGES ===\n")

    # --- Load curated ---
    with open(INPUT_JSON, "r", encoding="utf-8") as f:
        curated = json.load(f)

    total_entries = len(curated)
    print(f"📚 Entries to process: {total_entries}\n")

    # --- Open Word ---
    word, doc = open_word(DOCX_INPUT)

    updated = {}
    warnings = []
    processed = 0

    try:
        print("🟡 Repaginating...")
        doc.Repaginate()
        print("🟢 Pagination ready\n")

        for name, entry in curated.items():

            pages = collect_pages_for_entry(doc, name, entry)

            # --- warning for suspicious entries ---
            if len(pages) <= MIN_PAGES_WARNING:
                warnings.append(name)

            # --- preserve entry, replace only pages ---
            new_entry = dict(entry)
            new_entry["pages"] = pages

            updated[name] = new_entry

            processed += 1

            if processed % PROGRESS_EVERY == 0:
                print(f"[PROGRESS] {processed}/{total_entries}")

    finally:
        close_word(word, doc)

    # --- Save output ---
    OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(updated, f, indent=2)

    # --- Report ---
    print("\n=== SUMMARY ===\n")
    print(f"✔ Entries processed: {processed}")
    print(f"⚠ Low-page entries: {len(warnings)}")

    if warnings:
        print("\n⚠ Entries with very few pages (first 20):")
        for name in warnings[:20]:
            print(f"  - {name}")

    print(f"\n💾 Saved → {OUTPUT_JSON}")
    print("\n✅ Done\n")


if __name__ == "__main__":
    main()