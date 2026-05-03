"""
===============================================================================
FILE: fix_pages_from_extracted.py
-------------------------------------------------------------------------------
Compare extracted vs raw_fixed and classify entries + apply manual overrides

MODES:
--validate → generate CSV reports only
--update   → apply fixes (raw, normalized, manual overrides)

EDITED VALUES:
0 → untouched → normal pipeline
1 → found in raw → force raw lookup
2 → search in book using normalized
3 → manual override → use pages_raw (DO NOT SEARCH)

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

def parse_pages(pages_str):
    if not pages_str:
        return []
    return [int(p.strip()) for p in pages_str.split(",") if p.strip().isdigit()]

# Placeholder: replace with your real Word COM search
def search_in_book(term):
    print(f"🔍 Searching in book: {term}")
    return []  # return list of page numbers

# ---------------------------------------------------------------------------
# LOAD
# ---------------------------------------------------------------------------

print("\n=== VALIDATE / FIX PAGES (WITH EDITED LOGIC) ===\n")

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

# ---------------------------------------------------------------------------
# LOOKUPS
# ---------------------------------------------------------------------------

raw_lookup = {normalize(k): v for k, v in raw.items()}

curated_lookup = {
    normalize(k): v.get("normalized")
    for k, v in curated.items()
}

# ---------------------------------------------------------------------------
# LOAD CSV (edited + manual pages)
# ---------------------------------------------------------------------------

edited_map = {}
manual_pages = {}

if CSV_MISSING.exists():
    with open(CSV_MISSING, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            key = normalize(row["key"])

            edited_map[key] = row.get("edited", "0")

            if row.get("pages_raw"):
                manual_pages[key] = parse_pages(row["pages_raw"])

print(f"✔ Loaded edited entries: {len(edited_map)}")
print(f"✔ Loaded manual pages: {len(manual_pages)}\n")

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
    edited = edited_map.get(norm_key, "0")

    pages_extracted = entry.get("pages", [])
    raw_entry = None
    normalized_used = ""
    source = ""

    # ---------------------------------------------------
    # EDITED = 3 → MANUAL OVERRIDE
    # ---------------------------------------------------
    if edited == "3" and norm_key in manual_pages:
        pages_raw = manual_pages[norm_key]

        if mode == "update":
            entry["pages"] = pages_raw
            updated += 1

        continue

    # ---------------------------------------------------
    # EDITED = 2 → SEARCH BOOK
    # ---------------------------------------------------
    if edited == "2":
        term = entry.get("normalized") or key
        pages_raw = search_in_book(term)

        if pages_raw and mode == "update":
            entry["pages"] = pages_raw
            updated += 1

        continue

    # ---------------------------------------------------
    # EDITED = 1 → FORCE RAW
    # ---------------------------------------------------
    if edited == "1":
        raw_entry = raw_lookup.get(norm_key)
        source = "FORCED_RAW"

    # ---------------------------------------------------
    # NORMAL PIPELINE
    # ---------------------------------------------------
    if not raw_entry:
        raw_entry = raw_lookup.get(norm_key)
        if raw_entry:
            source = "DIRECT"

    if not raw_entry:
        norm_name = curated_lookup.get(norm_key)
        if norm_name:
            raw_entry = raw_lookup.get(normalize(norm_name))
            if raw_entry:
                normalized_used = norm_name
                source = "CURATED_NORMALIZED"

    if not raw_entry:
        norm_name = entry.get("normalized")
        if norm_name:
            raw_entry = raw_lookup.get(normalize(norm_name))
            if raw_entry:
                normalized_used = norm_name
                source = "EXTRACTED_NORMALIZED"

    pages_raw = raw_entry.get("pages", []) if raw_entry else []

    # ---------------------------------------------------
    # CLASSIFY
    # ---------------------------------------------------
    if not raw_entry:
        status = "MISSING"
    elif pages_extracted == pages_raw:
        status = "OK"
    elif source in ("CURATED_NORMALIZED", "EXTRACTED_NORMALIZED"):
        status = "NORMALIZED_MATCH"
    else:
        status = "DIFF"

    # ---------------------------------------------------
    # UPDATE
    # ---------------------------------------------------
    if mode == "update" and status in ("DIFF", "NORMALIZED_MATCH"):
        entry["pages"] = pages_raw
        updated += 1

    # ---------------------------------------------------
    # CSV OUTPUT
    # ---------------------------------------------------
    base_row = {
        "key": key,
        "pages_extracted": ", ".join(map(str, pages_extracted)),
        "pages_raw": ", ".join(map(str, pages_raw)),
    }

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
            "edited": "0",
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
write_csv(CSV_NORMALIZED, normalized_rows,
          ["key", "pages_extracted", "pages_raw", "normalized_used", "source"])
write_csv(CSV_MISSING, missing_rows,
          ["key", "normalized", "edited", "pages_extracted", "pages_raw"])

print(f"📄 OK CSV         → {CSV_OK}")
print(f"📄 DIFF CSV       → {CSV_DIFF}")
print(f"📄 NORMALIZED CSV → {CSV_NORMALIZED}")
print(f"📄 MISSING CSV    → {CSV_MISSING}")

# ---------------------------------------------------------------------------
# SAVE JSON
# ---------------------------------------------------------------------------

if mode == "update":
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(extracted, f, indent=2)

    print(f"\n💾 Updated JSON → {OUTPUT_JSON}")
    print(f"✔ Entries updated: {updated}")

print("\n=== DONE ===\n")