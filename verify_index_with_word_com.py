"""
===============================================================================
FILE: verify_index_with_word_com.py
-------------------------------------------------------------------------------
Full verification with:
- normalized + reverse names
- aliases + reverse aliases
- union of all matches
- statistics + debug
===============================================================================
"""

import json
import re
from pathlib import Path
import win32com.client as win32

# ---------------- CONFIG ---------------- #

DOCX_INPUT = "HITS AND HAPPINESS FINAL 2 Format MOM Discog.docx"
FINAL_JSON = "index_curated_final.json"

MIN_PAGES = 2
DEBUG = False

# ---------------- WORD ---------------- #

def open_word(path):
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(str(Path(path).resolve()))
    return word, doc


def close_word(word, doc):
    doc.Close(False)
    word.Quit()


# ---------------- NAME UTIL ---------------- #

def reverse_name(name):
    if "," in name:
        parts = [p.strip() for p in name.split(",")]
        if len(parts) == 2:
            return f"{parts[1]} {parts[0]}"
    return None


# ---------------- SEARCH ---------------- #

def find_pages(doc, text):
    pages = set()

    pattern = re.compile(rf"\b{re.escape(text)}\b", re.IGNORECASE)
    full_text = doc.Content.Text

    for match in pattern.finditer(full_text):
        try:
            rng = doc.Range(Start=match.start(), End=match.end())
            page = rng.Information(3)
            pages.add(page)
        except:
            continue

    return pages


# ---------------- VERIFY ---------------- #

def verify(doc, index):
    stats = {
        "total": len(index),
        "checked": 0,
        "skipped": 0,
        "correct": 0,
        "mismatch": 0,
        "missing": 0,
        "extra": 0
    }

    mismatches = {}

    print(f"\n🔍 Verifying {stats['total']} entries...\n")

    for name, data in index.items():

        expected = set(data["pages"])

        if len(expected) < MIN_PAGES:
            stats["skipped"] += 1
            continue

        stats["checked"] += 1

        # 🔥 build full identity search set
        search_terms = set()

        # canonical
        search_terms.add(name)

        # reverse canonical
        rev = reverse_name(name)
        if rev:
            search_terms.add(rev)

        # aliases
        for alias in data.get("aliases", []):
            search_terms.add(alias)

            rev_alias = reverse_name(alias)
            if rev_alias:
                search_terms.add(rev_alias)

        # 🔥 union of all matches
        found = set()

        for term in search_terms:
            found.update(find_pages(doc, term))

        if DEBUG:
            print(f"\n{name}")
            print(f" terms: {search_terms}")
            print(f" expected: {sorted(expected)}")
            print(f" found:    {sorted(found)}")

        if expected == found:
            stats["correct"] += 1
        else:
            stats["mismatch"] += 1

            missing_pages = expected - found
            extra_pages = found - expected

            if missing_pages:
                stats["missing"] += 1

            if extra_pages:
                stats["extra"] += 1

            mismatches[name] = {
                "expected": sorted(expected),
                "found": sorted(found),
                "missing": sorted(missing_pages),
                "extra": sorted(extra_pages)
            }

        if stats["checked"] % 50 == 0:
            print(f"[PROGRESS] {stats['checked']} checked")

    return stats, mismatches


# ---------------- REPORT ---------------- #

def print_report(stats, mismatches):
    print("\n=== FINAL REPORT ===\n")

    print(f"Total entries:        {stats['total']}")
    print(f"Checked entries:      {stats['checked']}")
    print(f"Skipped (<{MIN_PAGES} pages): {stats['skipped']}")
    print()

    print(f"Correct entries:      {stats['correct']}")
    print(f"Mismatches:           {stats['mismatch']}")
    print(f"  Missing pages:      {stats['missing']}")
    print(f"  Extra pages:        {stats['extra']}")
    print()

    if stats["checked"] > 0:
        accuracy = (stats["correct"] / stats["checked"]) * 100
        print(f"Accuracy:             {accuracy:.2f}%")

    if mismatches:
        print("\n--- SAMPLE MISMATCHES ---\n")
        for name, data in list(mismatches.items())[:10]:
            print(name)
            print(f"  expected: {data['expected']}")
            print(f"  found:    {data['found']}")
            if data["missing"]:
                print(f"  missing:  {data['missing']}")
            if data["extra"]:
                print(f"  extra:    {data['extra']}")
            print()


# ---------------- MAIN ---------------- #

def main():
    print("\n=== VERIFY INDEX WITH WORD COM ===\n")

    with open(FINAL_JSON, "r", encoding="utf-8") as f:
        index = json.load(f)

    word, doc = open_word(DOCX_INPUT)

    try:
        stats, mismatches = verify(doc, index)
    finally:
        close_word(word, doc)

    print_report(stats, mismatches)

    print("\n✅ Verification complete\n")


if __name__ == "__main__":
    main()