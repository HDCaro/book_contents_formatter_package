"""
===============================================================================
FILE: fix_pages_from_extracted.py
LOCATION: src/index/fix/patch/
-------------------------------------------------------------------------------
Recalculate and UPDATE pages for extracted index

✔ Uses Word COM to detect ALL occurrences
✔ Merges ALL found pages (including extra pages)
✔ Replaces pages list with corrected version
✔ Keeps aliases intact

OUTPUT
------
index_curated_fixed_pages.json

DEBUG
-----
✔ Prints input/output paths
✔ Shows page counts per entry
✔ Shows page differences (optional visibility)
===============================================================================
"""

import json
import re
from pathlib import Path
import win32com.client as win32

# ---------------- ROOT ---------------- #

def find_project_root():
    current = Path(__file__).resolve()
    for parent in [current] + list(current.parents):
        if (parent / "src").exists() and (parent / "data").exists():
            return parent
    raise RuntimeError("Project root not found")

BASE_DIR = find_project_root()

# ---------------- PATHS ---------------- #

INPUT_JSON = BASE_DIR / "data/index/intermediate/index_curated_extracted.json"
DOCX_INPUT = BASE_DIR / "data/index/input/book/HITS AND HAPPINESS FINAL 2 Format MOM Discog.docx"
OUTPUT_JSON = BASE_DIR / "data/index/intermediate/index_curated_fixed_pages.json"

# ---------------- SEARCH ---------------- #

def find_pages(doc, term):
    """
    Find ALL pages where term appears
    """
    pages = set()

    pattern = re.compile(rf"\b{re.escape(term)}\b", re.IGNORECASE)
    text = doc.Content.Text

    for m in pattern.finditer(text):
        try:
            rng = doc.Range(Start=m.start(), End=m.end())
            page = int(rng.Information(1))
            pages.add(page)
        except:
            pass

    return pages

# ---------------- MAIN ---------------- #

def main():
    print("\n=== FIX + MERGE PAGES FROM EXTRACTED ===\n")

    print(f"📥 INPUT JSON:  {INPUT_JSON}")
    print(f"📥 BOOK DOCX:   {DOCX_INPUT}")
    print(f"📤 OUTPUT JSON: {OUTPUT_JSON}\n")

    if not INPUT_JSON.exists():
        print("❌ Missing input JSON")
        return

    if not DOCX_INPUT.exists():
        print("❌ Missing book DOCX")
        return

    data = json.load(open(INPUT_JSON, encoding="utf-8"))

    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(str(DOCX_INPUT))

    fixed = {}

    try:
        doc.Repaginate()

        for name, entry in data.items():

            original_pages = set(entry.get("pages", []))

            # 🔥 IMPORTANT: search using name + aliases
            terms = [name]
            terms += entry.get("aliases", [])
            terms += entry.get("aliases_external", [])

            found_pages = set()

            for term in terms:
                found_pages.update(find_pages(doc, term))

            # 🔥 FINAL MERGE: keep everything found
            final_pages = sorted(found_pages)

            fixed[name] = {
                "aliases": entry.get("aliases", []),
                "aliases_external": entry.get("aliases_external", []),
                "pages": final_pages
            }

            # ---- DEBUG ----
            added = found_pages - original_pages
            removed = original_pages - found_pages

            print(f"✔ {name}")
            print(f"   old: {len(original_pages)} pages")
            print(f"   new: {len(final_pages)} pages")

            if added:
                print(f"   ➕ added: {sorted(added)}")

            if removed:
                print(f"   ➖ removed: {sorted(removed)}")

    finally:
        doc.Close(False)
        word.Quit()

    OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(fixed, f, indent=2)

    print(f"\n💾 Saved → {OUTPUT_JSON}")
    print(f"✅ Updated {len(fixed)} entries\n")


if __name__ == "__main__":
    main()