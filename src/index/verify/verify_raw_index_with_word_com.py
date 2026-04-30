"""
===============================================================================
FILE: verify_raw_index_with_word_com.py
-------------------------------------------------------------------------------
PURPOSE
-------
Verifies RAW index and generates:

1) index_discrepancies.json  → diagnostic
2) index_transaction_suggestions.json → editable transaction scaffold

===============================================================================
"""

import json
import re
from pathlib import Path
import win32com.client as win32

# ---------------- CONFIG ---------------- #
# ---------------- BASE PATH ---------------- #
BASE_DIR = Path(__file__).resolve().parents[3]

# ---------------- INPUT ---------------- #
DOCX_INPUT = BASE_DIR / "data/index/input/HITS AND HAPPINESS FINAL 2 Format MOM Discog.docx"
RAW_JSON = BASE_DIR / "data/index/intermediate/index_raw.json"

# ---------------- OUTPUT (INTERMEDIATE) ---------------- #
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
        "missing": 0,
        "extra": 0
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

        missing_pages = sorted(expected - found)
        extra_pages = sorted(found - expected)

        # --- TRANSACTION SCAFFOLD ---
        transactions[name] = {
            "action": "keep"
        }

        if expected == found:
            stats["correct"] += 1
        else:
            stats["mismatch"] += 1

            if missing_pages:
                stats["missing"] += 1

            if extra_pages:
                stats["extra"] += 1
                transactions[name]["extra_pages"] = extra_pages

            mismatches[name] = {
                "expected": sorted(expected),
                "found": sorted(found),
                "missing": missing_pages,
                "extra": extra_pages
            }

        if stats["checked"] % 50 == 0:
            print(f"[PROGRESS] {stats['checked']} checked")

    return stats, mismatches, transactions

# ---------------- REPORT ---------------- #

def print_report(stats):
    print("\n=== FINAL REPORT ===\n")
    print(f"Total entries:   {stats['total']}")
    print(f"Checked:         {stats['checked']}")
    print(f"Correct:         {stats['correct']}")
    print(f"Mismatches:      {stats['mismatch']}")
    print(f"Missing pages:   {stats['missing']}")
    print(f"Extra pages:     {stats['extra']}")
    accuracy = (stats["correct"] / stats["checked"]) * 100
    print(f"Accuracy:        {accuracy:.2f}%")

# ---------------- SAVE ---------------- #

def save_json(data, path, label):
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