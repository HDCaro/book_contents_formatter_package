"""
===============================================================================
FILE: fix_raw_with_discrepancies.py
-------------------------------------------------------------------------------
PURPOSE
-------
Creates index_transaction_edit.json by merging additional pages from verification
into index_raw.json.

This prepares a human-editable working file for editorial refinement.

FLOW
----
RAW + discrepancies.additional_pages → index_transaction_edit.json

===============================================================================
"""

import json
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[4]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.utils.book_project import get_active_book_root

# ---------------- BASE PATH ---------------- #

BASE_DIR = PROJECT_ROOT
BOOK_ROOT = get_active_book_root(BASE_DIR)

# ---------------- INPUT (INTERMEDIATE) ---------------- #

RAW_JSON = BOOK_ROOT / "work" / "index" / "intermediate" / "index_raw.json"
DISCREPANCY_JSON = BOOK_ROOT / "work" / "index" / "intermediate" / "index_discrepancies.json"

# ---------------- OUTPUT (INTERMEDIATE) ---------------- #

OUTPUT_JSON = BOOK_ROOT / "work" / "index" / "intermediate" / "index_transaction_edit.json"


def main():
    print("\n=== BUILD TRANSACTION EDIT FILE ===\n")

    with open(RAW_JSON, "r", encoding="utf-8") as f:
        raw = json.load(f)

    with open(DISCREPANCY_JSON, "r", encoding="utf-8") as f:
        discrepancies = json.load(f)

    updated = 0

    for name, data in discrepancies.items():
        additional_pages = data.get("additional_pages", [])

        if not additional_pages:
            continue

        if name not in raw:
            continue

        current_pages = set(raw[name]["pages"])
        new_pages = current_pages.union(additional_pages)

        if new_pages != current_pages:
            raw[name]["pages"] = sorted(new_pages)
            updated += 1

    # Ensure folder exists
    OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(raw, f, indent=2)

    print(f"✔ Updated entries: {updated}")
    print(f"💾 Saved → {OUTPUT_JSON}\n")


if __name__ == "__main__":
    main()
