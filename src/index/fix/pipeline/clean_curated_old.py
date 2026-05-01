"""
===============================================================================
FILE: clean_curated_old.py
-------------------------------------------------------------------------------
PURPOSE
-------
Clean old curated file by:

1) Removing entries with less than 2 pages
2) Removing semantic noise (My Mom, My Dad, etc.)

Preserves:
- action
- type
- destination
- normalized

===============================================================================
"""

import json
from pathlib import Path

# ---------------- CONFIG ---------------- #

BASE_DIR = Path(__file__).resolve().parents[3]

CURATED_JSON = BASE_DIR / "data/index/intermediate/index_curated_old.json"
OUTPUT_JSON = BASE_DIR / "data/index/intermediate/index_curated_old_filtered.json"

MIN_PAGES = 2

# ---------------- NOISE FILTER ---------------- #

def is_noise(name):
    words = name.lower().split()

    if not words:
        return False

    if words[0] in {"my", "your", "his", "her", "our", "their"}:
        return True

    return False

# ---------------- MAIN ---------------- #

def main():
    print("\n=== CLEAN CURATED OLD ===\n")

    with open(CURATED_JSON, "r", encoding="utf-8") as f:
        curated = json.load(f)

    cleaned = {}

    removed_single_page = []
    removed_noise = []

    for name, entry in curated.items():

        pages = entry.get("pages", [])

        # --- FILTER: single page ---
        if len(pages) < MIN_PAGES:
            removed_single_page.append(name)
            continue

        # --- FILTER: noise ---
        if is_noise(name):
            removed_noise.append(name)
            continue

        cleaned[name] = entry

    # ---------------- SAVE ---------------- #

    OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(cleaned, f, indent=2)

    # ---------------- REPORT ---------------- #

    print(f"✔ Kept entries: {len(cleaned)}")
    print(f"✖ Removed (single page): {len(removed_single_page)}")
    print(f"✖ Removed (noise): {len(removed_noise)}")

    if removed_noise:
        print("\nRemoved noise entries (first 20):")
        for name in removed_noise[:20]:
            print(f"  - {name}")

    print(f"\n💾 Saved → {OUTPUT_JSON}\n")


if __name__ == "__main__":
    main()