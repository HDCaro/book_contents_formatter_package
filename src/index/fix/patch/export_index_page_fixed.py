"""
===============================================================================
FILE: export_index_page_fixed.py
LOCATION: src/index/fix/patch/
-------------------------------------------------------------------------------
Generate final CSV from fully corrected index JSON

INPUT:
- index_curated_extracted_fixed.json

OUTPUT:
- index_page_fixed.csv

DESCRIPTION:
This script exports the final, fully corrected index entries into a CSV file.

The input JSON must already include:
✔ corrected page numbers (raw, normalized, manual overrides applied)
✔ finalized keys and normalized values

This CSV is intended for:
✔ final validation
✔ manual review (optional)
✔ feeding into DOCX generator pipeline

NOTE:
This script does NOT perform any corrections.
It assumes all fixes were applied in previous steps.
===============================================================================
"""

import json
import csv
from pathlib import Path

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

BASE_PATH = Path(r"C:\Projects\Python\book_contents_formatter_package")

INPUT_JSON = BASE_PATH / "data/index/intermediate/index_curated_extracted_fixed.json"
OUTPUT_CSV = BASE_PATH / "data/index/output/index_page_fixed.csv"

# ---------------------------------------------------------------------------
# LOAD INPUT
# ---------------------------------------------------------------------------

print("\n=== EXPORT FINAL INDEX CSV ===\n")

if not INPUT_JSON.exists():
    print(f"❌ Missing input JSON: {INPUT_JSON}")
    exit()

data = json.load(open(INPUT_JSON, encoding="utf-8"))

# ---------------------------------------------------------------------------
# BUILD ROWS
# ---------------------------------------------------------------------------

rows = []

for key, entry in data.items():
    rows.append({
        "key": key,
        "normalized": entry.get("normalized", ""),
        "pages": ", ".join(map(str, entry.get("pages", [])))
    })

# ---------------------------------------------------------------------------
# SAVE CSV
# ---------------------------------------------------------------------------

OUTPUT_CSV.parent.mkdir(parents=True, exist_ok=True)

with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:
    writer = csv.DictWriter(f, fieldnames=["key", "normalized", "pages"])
    writer.writeheader()
    writer.writerows(rows)

print(f"📄 Output CSV → {OUTPUT_CSV}")

# ---------------------------------------------------------------------------
# SUMMARY
# ---------------------------------------------------------------------------

print("\n=== SUMMARY ===\n")
print(f"Total entries exported: {len(rows)}\n")