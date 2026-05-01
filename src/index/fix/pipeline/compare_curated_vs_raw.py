"""
===============================================================================
FILE: compare_curated_vs_raw.py
-------------------------------------------------------------------------------
PURPOSE
-------
Compare curated entries (editorial) against consolidated raw data.

✔ Detect missing pages in curated
✔ Detect entries present in raw but not curated
✔ Provide a sanity check before final merge

INPUT
-----
data/index/intermediate/index_raw_fixed.json
data/index/intermediate/index_curated_old_pages.json

OUTPUT
------
Console report only

===============================================================================
"""

import json
from pathlib import Path

# ---------------- PATHS ---------------- #

def find_project_root():
    current = Path(__file__).resolve()
    for parent in [current] + list(current.parents):
        if (parent / "src").exists() and (parent / "data").exists():
            return parent
    raise RuntimeError("Project root not found")

BASE_DIR = find_project_root()

RAW = BASE_DIR / "data/index/intermediate/index_raw_fixed.json"
CURATED = BASE_DIR / "data/index/intermediate/index_curated_old_pages.json"

# ---------------- MAIN ---------------- #

def main():
    print("\n=== COMPARE CURATED vs RAW ===\n")

    if not RAW.exists():
        print(f"❌ Missing RAW file: {RAW}")
        return

    if not CURATED.exists():
        print(f"❌ Missing CURATED file: {CURATED}")
        return

    raw = json.load(open(RAW, encoding="utf-8"))
    curated = json.load(open(CURATED, encoding="utf-8"))

    missing_pages_count = 0
    missing_entries = []

    # --- compare pages ---
    for name, c in curated.items():
        raw_entry = raw.get(name)

        if not raw_entry:
            continue

        raw_pages = set(raw_entry.get("pages", []))
        cur_pages = set(c.get("pages", []))

        extra = raw_pages - cur_pages

        if extra:
            print(f"\n📌 {name}")
            print(f"  Missing pages in curated: {sorted(extra)}")
            missing_pages_count += 1

    # --- entries missing in curated ---
    for name in raw:
        if name not in curated:
            missing_entries.append(name)

    # --- summary ---
    print("\n=== SUMMARY ===\n")
    print(f"Entries with missing pages: {missing_pages_count}")
    print(f"Entries not in curated: {len(missing_entries)}")

    if missing_entries:
        print("\nSample entries not in curated:")
        for name in missing_entries[:20]:
            print(f"  - {name}")

    print("\n✅ Comparison complete\n")


if __name__ == "__main__":
    main()