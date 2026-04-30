"""
===============================================================================
FILE: generate_index_data_with_COM.py
-------------------------------------------------------------------------------
Refactored generator using Word Find for accurate pagination

FEATURES
--------
- Accurate page numbers via Word Find
- Repagination fix
- Filters:
    - single-page → index_raw_excluded.json
    - too many pages → index_raw_filtered_out.json
- Clean progress reporting

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
CANDIDATES_JSON = "index_candidates_to_add.json"

MIN_OCCURRENCES = 2
MAX_PAGES_PER_ENTRY = 50

PROGRESS_EVERY = 50

# ---------------- REGEX ---------------- #

PATTERN = re.compile(
    r"\b([A-Z][a-zA-Z]+(?:[-'][A-Za-z]+)*(?:\s+[A-Z][a-zA-Z]+(?:[-'][A-Za-z]+)*)+)\b"
)

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

# ---------------- DISCOVERY ---------------- #

def discover_candidates(text):
    print("🔍 Discovering candidates (regex)...")

    candidates = set()

    for match in PATTERN.finditer(text):
        candidates.add(match.group(1))

    print(f"🟢 Candidates found: {len(candidates)}\n")
    return sorted(candidates)

# ---------------- WORD FIND ---------------- #

def find_pages_word(doc, term):
    pages = set()

    rng = doc.Content.Duplicate
    find = rng.Find

    find.ClearFormatting()
    find.Text = term
    find.Forward = True
    find.Wrap = 1  # wdFindContinue

    find.MatchCase = False
    find.MatchWholeWord = True
    find.MatchWildcards = False
    find.MatchPhrase = True

    while find.Execute():
        try:
            page = int(rng.Information(1))
            pages.add(page)
        except:
            pass

        next_start = rng.End
        if next_start >= doc.Content.End:
            break

        rng.Start = next_start
        rng.End = doc.Content.End

    return pages

# ---------------- BUILD ---------------- #

def build_index(doc, candidates):
    print("🟡 Assigning pages using Word Find...\n")

    result = {}
    excluded = {}
    filtered_out = {}

    start = time.time()

    total = len(candidates)

    for i, name in enumerate(candidates, 1):

        pages = sorted(find_pages_word(doc, name))

        if not pages:
            continue

        entry = {
            "normalized": name,
            "type": "unknown",
            "pages": pages,
            "action": "keep"
        }

        # 🔥 too many pages → noisy
        if len(pages) > MAX_PAGES_PER_ENTRY:
            filtered_out[name] = entry
        elif len(pages) < MIN_OCCURRENCES:
            excluded[name] = entry
        else:
            result[name] = entry

        # 🔥 VERBOSE PROGRESS
        if i % PROGRESS_EVERY == 0 or i == total:
            elapsed = time.time() - start
            short_name = (name[:40] + "...") if len(name) > 40 else name

            print(
                f"[PROGRESS] {i}/{total} | "
                f"'{short_name}' | "
                f"{len(pages)} pages | "
                f"{elapsed:.1f}s"
            )

    print(f"\n🟢 Page assignment complete in {time.time() - start:.2f}s\n")

    print(f"✔ Kept (≥{MIN_OCCURRENCES} pages): {len(result)}")
    print(f"✖ Excluded (single-page): {len(excluded)}")
    print(f"⚠ Filtered-out (too many pages): {len(filtered_out)}\n")

    return result, excluded, filtered_out

# ---------------- MAIN ---------------- #

def main():
    print("\n=== GENERATE INDEX (WORD FIND MODE) ===\n")

    word, doc = open_word(DOCX_INPUT)

    try:
        print("🟡 Repaginating...")
        doc.Repaginate()
        print("🟢 Pagination ready\n")

        text = doc.Content.Text

        candidates = discover_candidates(text)

        result, excluded, filtered_out = build_index(doc, candidates)

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