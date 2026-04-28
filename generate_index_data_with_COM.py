"""
===============================================================================
FILE: generate_index_data_with_COM.py
-------------------------------------------------------------------------------

PURPOSE
-------
Scans a Word document using Microsoft Word COM automation and extracts
index candidates (names, titles, entities) with their page numbers.

This script:
1. Reads the full DOCX content via Word COM
2. Detects candidate entries using regex
3. Assigns page numbers using Word's internal pagination
4. Applies filtering (frequency + heuristics)
5. Classifies entries (person, band, work, etc.)
6. Produces a RAW index file
7. Compares RAW vs CURATED
8. Updates ONLY existing curated entries
9. Saves NEW entries separately for manual review

IMPORTANT
---------
This script DOES NOT build the final index document.
That is handled by:
    build_index_docx.py

This script is responsible ONLY for:
    detection + candidate generation + safe merge

INPUT FILES
-----------
- Source DOCX (book manuscript)
- index_curated.json (manually maintained)

OUTPUT FILES
------------
- index_raw.json
    → Automatically generated detected entries

- index_candidates_to_add.json
    → NEW entries NOT in curated (for manual review)

- index_curated.json
    → Updated ONLY for existing entries (no auto additions)

PROCESS SUMMARY
---------------
DOCX → detect candidates → classify → raw index
     → compare with curated
     → update existing entries (respecting action)
     → export NEW entries separately

FEATURES
--------
- Word COM page detection (accurate pagination)
- Hybrid filtering (frequency + heuristics)
- Discogs lookup (optional classification)
- Safe curated workflow (no automatic pollution)
- Candidate review pipeline

===============================================================================
"""

import re
import json
import time
from pathlib import Path
import requests
import win32com.client as win32

# ---------------- CONFIG ---------------- #

DOCX_INPUT = "HITS AND HAPPINESS FINAL 2 Format MOM Discog.docx"
OUTPUT_JSON = "index_raw.json"
CURATED_JSON = "index_curated.json"
CANDIDATES_JSON = "index_candidates_to_add.json"
CACHE_FILE = "discogs_cache.json"

MIN_OCCURRENCES = 2
DISCOGS_TOKEN = ""

# ---------------- REGEX ---------------- #

PATTERN = re.compile(
    r"\b([A-Z][a-zA-Z]+(?:[-'][A-Za-z]+)*(?:\s+[A-Z][a-zA-Z]+(?:[-'][A-Za-z]+)*)+)\b"
)

# ---------------- NORMALIZATION ---------------- #

def normalize_caps(text):
    def fix_word(word):
        parts = word.split('-')
        return "-".join(p.capitalize() if p.isupper() else p for p in parts)
    return " ".join(fix_word(w) for w in text.split())

# ---------------- HYBRID FILTER ---------------- #

def is_strong_candidate(name, pages):
    if len(pages) >= MIN_OCCURRENCES:
        return True

    words = name.split()

    if len(words) == 2:
        return True

    if any(c.isdigit() for c in name):
        return True

    if "-" in name:
        return True

    return False

# ---------------- DISCOGS CACHE ---------------- #

def load_cache():
    try:
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return {}

def save_cache(cache):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, indent=2)

discogs_cache = load_cache()

def discogs_search(query):
    key = query.lower()

    if key in discogs_cache:
        return discogs_cache[key]

    headers = {"User-Agent": "IndexBuilder/1.0"}
    if DISCOGS_TOKEN:
        headers["Authorization"] = f"Discogs token={DISCOGS_TOKEN}"

    try:
        r = requests.get(
            "https://api.discogs.com/database/search",
            params={"q": query},
            headers=headers,
            timeout=5
        )
        results = r.json().get("results", [])

        found = False
        dtype = None

        if results:
            found = True
            dtype = results[0].get("type")

        discogs_cache[key] = (found, dtype)
        return found, dtype

    except:
        discogs_cache[key] = (False, None)
        return False, None

# ---------------- CLASSIFICATION ---------------- #

def classify_entity(name):
    parts = name.split()

    if len(parts) == 2 and parts[0] != "The":
        return f"{parts[1]}, {parts[0]}", "person"

    if parts[0] == "The":
        return f"{' '.join(parts[1:])}, The", "band"

    found, dtype = discogs_search(name)

    if found:
        if dtype == "artist":
            return name, "band"
        if dtype in ("release", "master"):
            return name, "album"

    return name, "work"

# ---------------- WORD COM ---------------- #

def open_word(path):
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc = word.Documents.Open(str(Path(path).resolve()))
    return word, doc

def close_word(word, doc):
    doc.Close(False)
    word.Quit()

# ---------------- INDEX BUILD ---------------- #

def build_index_fast(doc):
    index = {}

    text = normalize_caps(doc.Content.Text)

    print("[INFO] Scanning document once...\n")
    start = time.time()

    for match in PATTERN.finditer(text):
        name = match.group(1)

        try:
            rng = doc.Range(Start=match.start(), End=match.end())
            page = int(rng.Information(3))
        except:
            continue

        index.setdefault(name, set()).add(page)

    print(f"\n[INFO] Scan complete in {time.time() - start:.2f}s")
    print(f"[INFO] Unique entries: {len(index)}\n")

    result = {}

    for name, pages in index.items():
        pages = sorted(pages)

        words = name.split()

        keep_basic = (
            len(words) <= 4 and
            all(w[0].isupper() for w in words)
        )

        if not keep_basic:
            continue

        if not is_strong_candidate(name, pages):
            continue

        normalized, typ = classify_entity(name)

        result[name] = {
            "normalized": normalized,
            "type": typ,
            "pages": pages,
            "action": "keep"
        }

    print(f"[INFO] Final entries: {len(result)}\n")

    save_cache(discogs_cache)

    return result

# ---------------- MERGE ---------------- #

def merge_curated(raw_data):
    path = Path(CURATED_JSON)

    if not path.exists():
        print("❌ curated file missing")
        return

    with open(path, "r", encoding="utf-8") as f:
        curated = json.load(f)

    raw_keys = set(raw_data.keys())
    cur_keys = set(curated.keys())

    added = raw_keys - cur_keys
    common = raw_keys & cur_keys

    print("\n=== MERGE REPORT ===\n")

    print("🔄 PROCESS EXISTING:\n")

    for k in sorted(common):
        action = curated[k].get("action", "keep")

        raw_pages = set(raw_data[k]["pages"])
        cur_pages = set(curated[k]["pages"])

        if action == "remove":
            print(f"[REMOVE] {k}")
            continue

        if action == "update":
            print(f"[UPDATE] {k}")
            curated[k]["pages"] = sorted(raw_pages)
            continue

        if action == "keep":
            union = sorted(cur_pages | raw_pages)
            if union != sorted(cur_pages):
                print(f"[KEEP→UPDATE] {k}")
                curated[k]["pages"] = union
            else:
                print(f"[KEEP] {k}")

    print("\n➕ NEW CANDIDATES (separate file):\n")

    candidates = {}

    for k in sorted(added):
        print(f"+ {k}")
        candidates[k] = raw_data[k]

    with open(CANDIDATES_JSON, "w", encoding="utf-8") as f:
        json.dump(candidates, f, indent=2)

    print(f"\n💾 Saved: {CANDIDATES_JSON}")

    with open(path, "w", encoding="utf-8") as f:
        json.dump(curated, f, indent=2)

    print("\n✅ Curated updated (clean)")

# ---------------- MAIN ---------------- #

def main():
    print("\n=== GENERATE INDEX DATA (COM) ===\n")

    word, doc = open_word(DOCX_INPUT)

    try:
        raw = build_index_fast(doc)

        with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
            json.dump(raw, f, indent=2)

        print("✅ Raw index generated")

        merge_curated(raw)

    finally:
        close_word(word, doc)

if __name__ == "__main__":
    main()