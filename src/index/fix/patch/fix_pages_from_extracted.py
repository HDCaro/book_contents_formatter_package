"""
===============================================================================
FILE: fix_pages_from_extracted.py
-------------------------------------------------------------------------------
Validate and optionally fix page numbers in extracted index

MODES:
--validate → generate reports only
--update   → update DIFF entries

OUTPUT:
- index_page_ok.csv
- index_page_diff.csv
- index_page_missing.csv
- index_curated_extracted_fixed.json (update mode)
===============================================================================
"""

import json
import csv
import re
import sys
from pathlib import Path

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

BASE_PATH = Path(__file__).resolve().parents[4]

EXTRACTED_JSON = BASE_PATH / "data/index/intermediate/index_curated_extracted.json"
RAW_JSON = BASE_PATH / "data/index/intermediate/index_raw_fixed.json"

OUTPUT_JSON = BASE_PATH / "data/index/intermediate/index_curated_extracted_fixed.json"

OUTPUT_DIR = BASE_PATH / "data/index/output"
CSV_OK = OUTPUT_DIR / "index_page_ok.csv"
CSV_DIFF = OUTPUT_DIR / "index_page_diff.csv"
CSV_MISSING = OUTPUT_DIR / "index_page_missing.csv"

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

def normalize(s):
    s = s.lower().strip()
    s = re.sub(r"^the\s+", "", s)
    s = re.sub(r"[^\w\s]", "", s)
    return s

# ---------------------------------------------------------------------------
# LOAD
# ---------------------------------------------------------------------------

print("\n=== FIX / VALIDATE PAGES ===\n")

if not EXTRACTED_JSON.exists():
    print("❌ Missing extracted JSON")
    exit()

if not RAW_JSON.exists():
    print("❌ Missing raw_fixed JSON")
    exit()

extracted = json.load(open(EXTRACTED_JSON, encoding="utf-8"))
raw = json.load(open(RAW_JSON, encoding="utf-8"))

# ---------------------------------------------------------------------------
# BUILD LOOKUP
# ---------------------------------------------------------------------------

raw_lookup = {normalize(k): v for k, v in raw.items()}

# ---------------------------------------------------------------------------
# MODE
# ---------------------------------------------------------------------------

mode = "validate"
if "--update" in sys.argv:
    mode = "update"

print(f"Mode: {mode.upper()}\n")

# ---------------------------------------------------------------------------
# PROCESS
# ---------------------------------------------------------------------------

ok_rows = []
diff_rows = []
missing_rows = []

updated_count = 0

for key, entry in extracted.items():

    norm_key = normalize(key)
    raw_entry = raw_lookup.get(norm_key)

    pages_extracted = entry.get("pages", [])
    pages_raw = raw_entry.get("pages", []) if raw_entry else []

    if not raw_entry:
        status = "MISSING"
    elif pages_extracted == pages_raw:
        status = "OK"
    else:
        status = "DIFF"

    row = {
        "key": key,
        "pages_extracted": ", ".join(map(str, pages_extracted)),
        "pages_raw": ", ".join(map(str, pages_raw))
    }

    if status == "OK":
        ok_rows.append(row)

    elif status == "DIFF":
        diff_rows.append(row)

        if mode == "update":
            entry["pages"] = pages_raw
            updated_count += 1

    else:
        missing_rows.append(row)

# ---------------------------------------------------------------------------
# SAVE CSVs
# ---------------------------------------------------------------------------

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

def write_csv(path, rows):
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["key", "pages_extracted", "pages_raw"])
        writer.writeheader()
        writer.writerows(rows)

write_csv(CSV_OK, ok_rows)
write_csv(CSV_DIFF, diff_rows)
write_csv(CSV_MISSING, missing_rows)

print(f"📄 OK CSV      → {CSV_OK}")
print(f"📄 DIFF CSV    → {CSV_DIFF}")
print(f"📄 MISSING CSV → {CSV_MISSING}")

# ---------------------------------------------------------------------------
# SAVE UPDATED JSON
# ---------------------------------------------------------------------------

if mode == "update":
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(extracted, f, indent=2)

    print(f"\n💾 Updated JSON → {OUTPUT_JSON}")
    print(f"✔ Entries updated: {updated_count}")

# ---------------------------------------------------------------------------
# SUMMARY
# ---------------------------------------------------------------------------

print("\n=== SUMMARY ===\n")
print(f"OK:       {len(ok_rows)}")
print(f"DIFF:     {len(diff_rows)}")
print(f"MISSING:  {len(missing_rows)}\n")