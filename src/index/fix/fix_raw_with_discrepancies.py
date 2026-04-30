"""
===============================================================================
FILE: fix_raw_with_discrepancies.py
-------------------------------------------------------------------------------
PURPOSE
-------
Creates index_raw_fixed.json by merging extra pages from verification
into index_raw.json.

This is a DELIVERY PATCH step:
RAW + discrepancies.extra → FIXED RAW

===============================================================================
"""

import json
from pathlib import Path

# ---------------- BASE PATH ---------------- #

BASE_DIR = Path(__file__).resolve().parents[3]

# ---------------- INPUT (INTERMEDIATE) ---------------- #

RAW_JSON = BASE_DIR / "data/index/intermediate/index_raw.json"
DISCREPANCY_JSON = BASE_DIR / "data/index/intermediate/index_discrepancies.json"

# ---------------- OUTPUT (INTERMEDIATE) ---------------- #

OUTPUT_JSON = BASE_DIR / "data/index/intermediate/index_raw_fixed.json"

def main():
    print("\n=== FIX RAW WITH DISCREPANCIES ===\n")

    with open(RAW_JSON, "r", encoding="utf-8") as f:
        raw = json.load(f)

    with open(DISCREPANCY_JSON, "r", encoding="utf-8") as f:
        discrepancies = json.load(f)

    updated = 0

    for name, data in discrepancies.items():
        extra_pages = data.get("extra", [])

        if not extra_pages:
            continue

        if name not in raw:
            continue

        current_pages = set(raw[name]["pages"])
        new_pages = current_pages.union(extra_pages)

        if new_pages != current_pages:
            raw[name]["pages"] = sorted(new_pages)
            updated += 1

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(raw, f, indent=2)

    print(f"✔ Updated entries: {updated}")
    print(f"💾 Saved → {OUTPUT_JSON}\n")


if __name__ == "__main__":
    main()