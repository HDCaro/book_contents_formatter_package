"""
===============================================================================
FILE: verify_raw_index_with_word_com.py
-------------------------------------------------------------------------------
PURPOSE
-------
Verifies RAW index against Word document.

UPDATED TERMINOLOGY:
- wrong_paging     → expected pages not matching actual (shift errors)
- additional_pages → extra valid occurrences found in document

OUTPUT
------
1) index_discrepancies.json
2) index_transaction_suggestions.json
===============================================================================
"""

import json
import re
from pathlib import Path
import win32com.client as win32

# ---------------- BASE PATH ---------------- #

BASE_DIR = Path(__file__).resolve().parents[3]

# ---------------- INPUT ---------------- #

DOCX_INPUT = BASE_DIR / "data/index/input/HITS AND HAPPINESS FINAL 2 Format MOM Discog.docx"
RAW_JSON = BASE_DIR / "data/index/intermediate/index_raw.json"

# ---------------- OUTPUT ---------------- #

DISCREPANCY_JSON = BASE_DIR / "data/index/intermediate/index_discrepancies.json"
TRANSACTION_JSON = BASE_DIR / "data/index/intermediate/index_transaction_suggestions.json"

MIN_PAGES = 2

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
            page = int(rng.Information(1))
            pages.add(page)
        except:
            continue

    return pages

# ---------------- VERIFY ---------------- #

def verify(doc, index):
    stats = {
        "total": len(index),
        "checked": 0,
        "correct": 0,
        "mismatch": 0,
        "wrong_paging": 0,
        "additional": 0
    }

    mismatches = {}
    transactions = {}

    print(f"\n🔍 Verifying RAW index ({stats['total']} entries)...\n")

    for i, (name, data) in enumerate(index.items(), 1):

        expected = set(data["pages"])

        if len(expected) < MIN_PAGES:
            continue

        stats["checked"] += 1

        search_terms = {name}
        rev = reverse_name(name)
        if rev:
            search_terms.add(rev)

        found = set()
        for term in search_terms:
            found.update(find_pages(doc, term))

        wrong_paging = sorted(expected - found)
        additional_pages = sorted(found - expected)

        # --- TRANSACTION SCAFFOLD ---
        transactions[name] = {
            "action": "keep"
        }

        if expected == found:
            stats["correct"] += 1
        else:
            stats["mismatch"] += 1

            if wrong_paging:
                stats["wrong_paging"] += 1

            if additional_pages:
                stats["additional"] += 1
                transactions[name]["additional_pages"] = additional_pages

            mismatches[name] = {
                "expected": sorted(expected),
                "found": sorted(found),
                "wrong_paging": wrong_paging,
                "additional_pages": additional_pages
            }

        if stats["checked"] % 50 == 0:
            print(f"[PROGRESS] {stats['checked']} checked")

    return stats, mismatches, transactions

# ---------------- REPORT ---------------- #

def print_report(stats):
    print("\n=== FINAL REPORT ===\n")

    print(f"Total entries:        {stats['total']}")
    print(f"Checked entries:      {stats['checked']}")
    print()

    print(f"Correct entries:      {stats['correct']}")
    print(f"Mismatches:           {stats['mismatch']}")
    print(f"  Wrong paging:       {stats['wrong_paging']}")
    print(f"  Additional pages:   {stats['additional']}")
    print()

    if stats["checked"] > 0:
        accuracy = (stats["correct"] / stats["checked"]) * 100
        print(f"Accuracy:             {accuracy:.2f}%")

# ---------------- SAVE ---------------- #

def save_json(data, path, label):
    path.parent.mkdir(parents=True, exist_ok=True)

    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)

    print(f"💾 Saved → {label}")

# ---------------- MAIN ---------------- #

def main():
    print("\n=== VERIFY RAW INDEX WITH WORD COM ===\n")

    with open(RAW_JSON, "r", encoding="utf-8") as f:
        index = json.load(f)

    word, doc = open_word(DOCX_INPUT)

    try:
        doc.Repaginate()
        stats, mismatches, transactions = verify(doc, index)
    finally:
        close_word(word, doc)

    print_report(stats)

    save_json(mismatches, DISCREPANCY_JSON, DISCREPANCY_JSON)
    save_json(transactions, TRANSACTION_JSON, TRANSACTION_JSON)

    print("\n✅ Verification complete\n")

if __name__ == "__main__":
    main()