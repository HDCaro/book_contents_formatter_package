"""
===============================================================================
FILE: fix_pages_from_extracted.py
LOCATION: src/index/fix/patch/
-------------------------------------------------------------------------------
Recalculate pages using Word COM (correct pagination)
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
    pages = set()
    pattern = re.compile(rf"\b{re.escape(term)}\b", re.IGNORECASE)

    text = doc.Content.Text

    for m in pattern.finditer(text):
        try:
            rng = doc.Range(Start=m.start(), End=m.end())
            pages.add(int(rng.Information(1)))
        except:
            pass

    return pages

# ---------------- MAIN ---------------- #

def main():
    print("\n=== FIX PAGES FROM EXTRACTED ===\n")

    data = json.load(open(INPUT_JSON, encoding="utf-8"))

    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(str(DOCX_INPUT))

    fixed = {}

    try:
        doc.Repaginate()

        for name, entry in data.items():

            terms = [name]
            terms += entry.get("aliases", [])
            terms += entry.get("aliases_external", [])

            pages = set()

            for t in terms:
                pages.update(find_pages(doc, t))

            fixed[name] = {
                "aliases": entry.get("aliases", []),
                "aliases_external": entry.get("aliases_external", []),
                "pages": sorted(pages)
            }

            print(f"✔ {name}: {len(pages)} pages")

    finally:
        doc.Close(False)
        word.Quit()

    OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(fixed, f, indent=2)

    print(f"\n💾 Saved → {OUTPUT_JSON}\n")


if __name__ == "__main__":
    main()