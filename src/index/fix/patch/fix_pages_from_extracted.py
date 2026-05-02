"""
===============================================================================
FILE: fix_pages_from_extracted.py
-------------------------------------------------------------------------------
Compare extracted vs raw_fixed and classify entries

MODES:
--validate → generate CSV reports only
--update   → update DIFF + NORMALIZED_MATCH entries

OUTPUT:
- index_page_ok.csv
- index_page_diff.csv
- index_page_missing.csv
- index_page_normalized.csv
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
CURATED_FINAL_JSON = BASE_PATH / "data/index/intermediate/index_curated_final.json"

OUTPUT_JSON = BASE_PATH / "data/index/intermediate/index_curated_extracted_fixed.json"

OUTPUT_DIR = BASE_PATH / "data/index/output"
CSV_OK = OUTPUT_DIR / "index_page_ok.csv"
CSV_DIFF = OUTPUT_DIR / "index_page_diff.csv"
CSV_MISSING = OUTPUT_DIR / "index_page_missing.csv"
CSV_NORMALIZED = OUTPUT_DIR / "index_page_normalized.csv"

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

def normalize(s):
    s = (s or "").lower().strip()
    s = re.sub(r"^the\s+", "", s)
    s = re.sub(r"[^\w\s]", "", s)
    return s

# ---------------------------------------------------------------------------
# LOAD
# ---------------------------------------------------------------------------

print("\n=== VALIDATE / FIX PAGES ===\n")

if not EXTRACTED_JSON.exists():
    print("❌ Missing extracted JSON")
    exit()

if not RAW_JSON.exists():
    print("❌ Missing raw_fixed JSON")
    exit()

extracted = json.load(open(EXTRACTED_JSON, encoding="utf-8"))
raw = json.load(open(RAW_JSON, encoding="utf-8"))

curated = {}
if CURATED_FINAL_JSON.exists():
    curated = json.load(open(CURATED_FINAL_JSON, encoding="utf-8"))
    print("✔ Using curated_final fallback\n")
else:
    print("⚠ curated_final not found\n")

# ---------------------------------------------------------------------------
# LOOKUPS
# ---------------------------------------------------------------------------

raw_lookup = {normalize(k): v for k, v in raw.items()}

curated_lookup = {
    normalize(k): v.get("normalized")
    for k, v in curated.items()
}

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
normalized_rows = []

updated = 0

for key, entry in extracted.items():

    norm_key = normalize(key)
    pages_extracted = entry.get("pages", [])

    raw_entry = None
    normalized_used = ""
    source = ""

    # --- 1) DIRECT ---
    raw_entry = raw_lookup.get(norm_key)
    if raw_entry:
        source = "DIRECT"

    # --- 2) CURATED NORMALIZED ---
    if not raw_entry:
        norm_name = curated_lookup.get(norm_key)
        if norm_name:
            raw_entry = raw_lookup.get(normalize(norm_name))
            if raw_entry:
                normalized_used = norm_name
                source = "CURATED_NORMALIZED"

    # --- 3) EXTRACTED NORMALIZED ---
    if not raw_entry:
        norm_name = entry.get("normalized")
        if norm_name:
            raw_entry = raw_lookup.get(normalize(norm_name))
            if raw_entry:
                normalized_used = norm_name
                source = "EXTRACTED_NORMALIZED"

    pages_raw = raw_entry.get("pages", []) if raw_entry else []

    # --- CLASSIFY ---
    if not raw_entry:
        status = "MISSING"
    elif pages_extracted == pages_raw:
        status = "OK"
    elif source in ("CURATED_NORMALIZED", "EXTRACTED_NORMALIZED"):
        status = "NORMALIZED_MATCH"
    else:
        status = "DIFF"

    # --- UPDATE ---
    if mode == "update" and status in ("DIFF", "NORMALIZED_MATCH"):
        entry["pages"] = pages_raw
        updated += 1

    # --- COMMON ROW ---
    base_row = {
        "key": key,
        "pages_extracted": ", ".join(map(str, pages_extracted)),
        "pages_raw": ", ".join(map(str, pages_raw)),
    }

    # --- BUCKETS ---
    if status == "OK":
        ok_rows.append(base_row)

    elif status == "DIFF":
        diff_rows.append(base_row)

    elif status == "NORMALIZED_MATCH":
        row = dict(base_row)
        row["normalized_used"] = normalized_used
        row["source"] = source
        normalized_rows.append(row)

    else:
        missing_rows.append({
            "key": key,
            "normalized": entry.get("normalized", ""),
            "pages_extracted": ", ".join(map(str, pages_extracted)),
            "pages_raw": ""
        })

# ---------------------------------------------------------------------------
# SAVE CSVs
# ---------------------------------------------------------------------------

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

def write_csv(path, rows, fields):
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)

write_csv(CSV_OK, ok_rows, ["key", "pages_extracted", "pages_raw"])
write_csv(CSV_DIFF, diff_rows, ["key", "pages_extracted", "pages_raw"])
write_csv(
    CSV_NORMALIZED,
    normalized_rows,
    ["key", "pages_extracted", "pages_raw", "normalized_used", "source"]
)
write_csv(
    CSV_MISSING,
    missing_rows,
    ["key", "normalized", "pages_extracted", "pages_raw"]
)

print(f"📄 OK CSV         → {CSV_OK}")
print(f"📄 DIFF CSV       → {CSV_DIFF}")
print(f"📄 NORMALIZED CSV → {CSV_NORMALIZED}")
print(f"📄 MISSING CSV    → {CSV_MISSING}")

# ---------------------------------------------------------------------------
# SAVE UPDATED JSON
# ---------------------------------------------------------------------------

if mode == "update":
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(extracted, f, indent=2)

    print(f"\n💾 Updated JSON → {OUTPUT_JSON}")
    print(f"✔ Entries updated: {updated}")

# ---------------------------------------------------------------------------
# SUMMARY
# ---------------------------------------------------------------------------

print("\n=== SUMMARY ===\n")
print(f"OK:                 {len(ok_rows)}")
print(f"DIFF:               {len(diff_rows)}")
print(f"NORMALIZED_MATCH:   {len(normalized_rows)}")
print(f"MISSING:            {len(missing_rows)}\n")