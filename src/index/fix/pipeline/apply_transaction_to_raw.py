"""
===============================================================================
FILE: apply_transaction_to_raw.py
-------------------------------------------------------------------------------
PURPOSE
-------
Preview editorial changes ONLY (pages are already fixed in RAW).

Focus:
- remove
- merge
- type changes
- missing entries

===============================================================================
"""

import json
from pathlib import Path

# ---------------- CONFIG ---------------- #

BASE_DIR = Path(__file__).resolve().parents[3]

RAW_FIXED_JSON = BASE_DIR / "data/index/intermediate/index_raw_fixed.json"
TRANSACTION_JSON = BASE_DIR / "data/index/intermediate/index_transaction_edit.json"

# ---------------- LOAD ---------------- #

def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

# ---------------- MAIN ---------------- #

def main():
    print("\n=== TRANSACTION PREVIEW (STRUCTURAL ONLY) ===\n")

    raw = load_json(RAW_FIXED_JSON)
    tx = load_json(TRANSACTION_JSON)

    removed = []
    merged = []
    type_changed = []
    missing_in_tx = []

    # --- PROCESS TRANSACTION ENTRIES --- #

    for name, rule in tx.items():

        action = rule.get("action", "keep")
        raw_entry = raw.get(name)

        # ---- REMOVE ----
        if action == "remove":
            print(f"\n🗑 REMOVE: {name}")
            removed.append(name)
            continue

        # ---- MERGE ----
        if action == "merge":
            dest = rule.get("destination")
            print(f"\n🔗 MERGE: {name} → {dest}")
            merged.append((name, dest))
            continue

        if not raw_entry:
            print(f"\n⚠ Missing in RAW: {name}")
            continue

        raw_type = raw_entry.get("type", "unknown")
        new_type = rule.get("type", raw_type)

        # ---- TYPE CHANGE ----
        if raw_type != new_type:
            print(f"\n🔁 TYPE CHANGE: {name}")
            print(f"  {raw_type} → {new_type}")
            type_changed.append(name)

    # --- FIND RAW ENTRIES NOT IN TX --- #

    for name in raw:
        if name not in tx:
            missing_in_tx.append(name)

    # ---------------- SUMMARY ---------------- #

    print("\n=== SUMMARY ===\n")

    print(f"🗑 Removed:        {len(removed)}")
    print(f"🔗 Merged:         {len(merged)}")
    print(f"🔁 Type changes:   {len(type_changed)}")
    print(f"⚠ Missing in TX:  {len(missing_in_tx)}")

    if missing_in_tx:
        print("\n⚠ Entries not in transaction file:")
        for name in missing_in_tx[:20]:
            print(f"  - {name}")

    print("\n🟡 DRY RUN — no file written\n")


if __name__ == "__main__":
    main()