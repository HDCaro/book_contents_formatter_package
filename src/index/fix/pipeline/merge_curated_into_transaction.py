"""
===============================================================================
FILE: merge_curated_into_transaction.py
-------------------------------------------------------------------------------
PURPOSE
-------
Merges old curated file into new transaction_edit file.

This preserves all editorial work (action, type, merges)
while keeping the new data structure.

OUTPUT
------
index_transaction_edit_merged.json

===============================================================================
"""

import json
from pathlib import Path

# ---------------- PATHS ---------------- #

BASE_DIR = Path(__file__).resolve().parents[3]

TRANSACTION_JSON = BASE_DIR / "data/index/intermediate/index_transaction_edit.json"
CURATED_JSON = BASE_DIR / "data/index/intermediate/index_curated_old.json"

OUTPUT_JSON = BASE_DIR / "data/index/intermediate/index_transaction_edit_merged.json"

# ---------------- LOAD ---------------- #

def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

# ---------------- MERGE ---------------- #

def main():
    print("\n=== MERGE CURATED INTO TRANSACTION ===\n")

    tx = load_json(TRANSACTION_JSON)
    curated = load_json(CURATED_JSON)

    updated = 0
    merged = {}

    for name, tx_entry in tx.items():

        new_entry = dict(tx_entry)  # copy

        curated_entry = curated.get(name)

        if curated_entry:
            # --- COPY EDITORIAL FIELDS ---
            if "action" in curated_entry:
                new_entry["action"] = curated_entry["action"]

            if "type" in curated_entry:
                new_entry["type"] = curated_entry["type"]

            if "destination" in curated_entry:
                new_entry["destination"] = curated_entry["destination"]

            if "exclude_pages" in curated_entry:
                new_entry["exclude_pages"] = curated_entry["exclude_pages"]

            updated += 1

        merged[name] = new_entry

    # --- FIND CURATED ENTRIES NOT IN TX --- #

    missing = []

    for name in curated:
        if name not in tx:
            missing.append(name)

    # ---------------- SAVE ---------------- #

    OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(merged, f, indent=2)

    # ---------------- REPORT ---------------- #

    print(f"✔ Entries updated from curated: {updated}")
    print(f"⚠ Entries in curated not in transaction: {len(missing)}")

    if missing:
        print("\nMissing entries (first 20):")
        for name in missing[:20]:
            print(f"  - {name}")

    print(f"\n💾 Saved → {OUTPUT_JSON}\n")


if __name__ == "__main__":
    main()