"""
===============================================================================
FILE: revalidate_missing_entries.py
-------------------------------------------------------------------------------
PURPOSE
-------
Find entries present in curated but missing in transaction,
recalculate their pages using Word COM, and preserve editorial fields.

PRESERVES:
- action
- destination
- type
- normalized

UPDATES:
- pages

===============================================================================
"""

import json
import re
from pathlib import Path
import win32com.client as win32

# ---------------- CONFIG ---------------- #

BASE_DIR = Path(__file__).resolve().parents[3]

DOCX_INPUT = BASE_DIR / "data/index/input/HITS AND HAPPINESS FINAL 2 Format MOM Discog.docx"

TRANSACTION_JSON = BASE_DIR / "data/index/intermediate/index_transaction_edit.json"
CURATED_JSON = BASE_DIR / "data/index/intermediate/index_curated_old.json"

OUTPUT_JSON = BASE_DIR / "data/index/intermediate/index_missing_revalidated.json"

MIN_PAGES = 1

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

# ---------------- MAIN ---------------- #

def main():
    print("\n=== REVALIDATE MISSING ENTRIES ===\n")

    with open(TRANSACTION_JSON, "r", encoding="utf-8") as f:
        tx = json.load(f)

    with open(CURATED_JSON, "r", encoding="utf-8") as f:
        curated = json.load(f)

    # --- find missing entries ---
    missing = [name for name in curated if name not in tx]

    print(f"⚠ Missing entries to revalidate: {len(missing)}")

    word, doc = open_word(DOCX_INPUT)

    results = {}

    try:
        doc.Repaginate()

        for i, name in enumerate(missing, 1):
            pages = find_pages(doc, name)

            if len(pages) < MIN_PAGES:
                continue

            # --- START FROM CURATED ENTRY ---
            entry = dict(curated[name])  # preserve everything

            # --- UPDATE ONLY PAGES ---
            entry["pages"] = sorted(pages)

            results[name] = entry

            if i % 20 == 0:
                print(f"[PROGRESS] {i}/{len(missing)}")

    finally:
        close_word(word, doc)

    # --- save ---
    OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2)

    print(f"\n✔ Revalidated entries: {len(results)}")
    print(f"💾 Saved → {OUTPUT_JSON}\n")


if __name__ == "__main__":
    main()