"""
===============================================================================
FILE: generate_index_batch.py
-------------------------------------------------------------------------------
BATCH INDEX GENERATOR

PHASES:
1) Regex discovery (broad, keeps variants)
2) Page assignment in one pass (fast)
3) Filtering:
    - remove semantic noise (My Dad, My Mom, etc.)
    - exclude single-page entries
    - filter overly frequent entries

DESIGN PRINCIPLE:
- Generator = detection (broad, slightly noisy OK)
- Curated = editorial truth (merge, normalize)

===============================================================================
"""

import re
import json
import time
from pathlib import Path
import win32com.client as win32

# ---------------- CONFIG ---------------- #

DOCX_INPUT = "HITS AND HAPPINESS FINAL 2 Format MOM Discog.docx"

OUTPUT_JSON = "index_raw.json"
EXCLUDED_JSON = "index_raw_excluded.json"
FILTERED_JSON = "index_raw_filtered_out.json"

MIN_OCCURRENCES = 2
MAX_PAGES_PER_ENTRY = 50
PROGRESS_EVERY = 5000

# ---------------- REGEX ---------------- #

PATTERN = re.compile(
    r"\b([A-Z][a-zA-Z]+(?:[-'][A-Za-z]+)*(?:\s+[A-Z][a-zA-Z]+(?:[-'][A-Za-z]+)*)+)\b"
)

# ---------------- NOISE FILTER ---------------- #

def is_noise(name):
    words = name.lower().split()

    if not words:
        return False

    if words[0] in {"my", "your", "his", "her", "our", "their"}:
        return True

    return False

# ---------------- WORD ---------------- #

def open_word(path):
    print("🟡 Opening Word...")
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(str(Path(path).resolve()))
    print("🟢 Document loaded\n")
    return word, doc

def close_word(word, doc):
    doc.Close(False)
    word.Quit()

# ---------------- STEP 1: DISCOVERY ---------------- #

def discover_candidates(text):
    print("🔍 Discovering candidates (regex)...")

    candidates = {}

    for match in PATTERN.finditer(text):
        name = match.group(1)

        if is_noise(name):
            continue

        if name not in candidates:
            candidates[name] = {
                "normalized": name,
                "type": "unknown",
                "pages": set(),
                "action": "keep"
            }

    print(f"🟢 Candidates discovered: {len(candidates)}\n")
    return candidates

# ---------------- STEP 2: PAGE ASSIGNMENT ---------------- #

def assign_pages(doc, text, candidates):
    print("🟡 Assigning pages (batch mode)...\n")

    start = time.time()

    for i, match in enumerate(PATTERN.finditer(text), 1):
        name = match.group(1)

        if name not in candidates:
            continue

        try:
            rng = doc.Range(Start=match.start(), End=match.end())
            page = int(rng.Information(1))
        except:
            continue

        candidates[name]["pages"].add(page)

        if i % PROGRESS_EVERY == 0:
            print(f"[SCAN] {i} matches processed...")

    print(f"\n🟢 Page assignment complete in {time.time() - start:.2f}s\n")

# ---------------- STEP 3: FINALIZE ---------------- #

def finalize(candidates):
    result = {}
    excluded = {}
    filtered_out = {}

    for name, entry in candidates.items():
        pages = sorted(entry["pages"])
        entry["pages"] = pages

        if not pages:
            continue

        if len(pages) > MAX_PAGES_PER_ENTRY:
            filtered_out[name] = entry
        elif len(pages) < MIN_OCCURRENCES:
            excluded[name] = entry
        else:
            result[name] = entry

    print("🟢 Final counts:\n")
    print(f"✔ Kept: {len(result)}")
    print(f"✖ Excluded (single-page): {len(excluded)}")
    print(f"⚠ Filtered-out (too many pages): {len(filtered_out)}\n")

    return result, excluded, filtered_out

# ---------------- MAIN ---------------- #

def main():
    print("\n=== BATCH INDEX GENERATOR ===\n")

    word, doc = open_word(DOCX_INPUT)

    try:
        print("🟡 Repaginating...")
        doc.Repaginate()
        print("🟢 Pagination ready\n")

        text = doc.Content.Text

        # STEP 1
        candidates = discover_candidates(text)

        # STEP 2
        assign_pages(doc, text, candidates)

        # STEP 3
        result, excluded, filtered_out = finalize(candidates)

        with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
            json.dump(result, f, indent=2)

        with open(EXCLUDED_JSON, "w", encoding="utf-8") as f:
            json.dump(excluded, f, indent=2)

        with open(FILTERED_JSON, "w", encoding="utf-8") as f:
            json.dump(filtered_out, f, indent=2)

        print(f"💾 Saved → {OUTPUT_JSON}")
        print(f"💾 Saved → {EXCLUDED_JSON}")
        print(f"💾 Saved → {FILTERED_JSON}")

    finally:
        close_word(word, doc)

    print("\n✅ Done\n")

if __name__ == "__main__":
    main()