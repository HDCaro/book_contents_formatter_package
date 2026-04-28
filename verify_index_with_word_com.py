import json
import win32com.client as win32
from pathlib import Path

DOCX_INPUT = "HITS AND HAPPINESS FINAL 2 Format MOM Discog.docx"
FINAL_JSON = "index_curated_final.json"

DEBUG = True


def open_word(path):
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(str(Path(path).resolve()))
    return word, doc


def close_word(word, doc):
    doc.Close(False)
    word.Quit()


def find_pages(doc, text):
    pages = set()

    rng = doc.Content
    find = rng.Find

    find.ClearFormatting()
    find.Text = text
    find.Forward = True
    find.Wrap = 0  # wdFindStop

    while find.Execute():
        page = rng.Information(3)  # wdActiveEndPageNumber
        pages.add(page)
        rng.Collapse(0)

    return sorted(pages)


def verify(doc, index):
    mismatches = {}
    missing = {}
    extra = {}

    total = len(index)
    print(f"\n🔍 Verifying {total} entries...\n")

    for i, (name, data) in enumerate(index.items(), 1):
        expected = set(data["pages"])

        found = set(find_pages(doc, name))

        if expected != found:
            mismatches[name] = {
                "expected": sorted(expected),
                "found": sorted(found)
            }

            missing_pages = expected - found
            extra_pages = found - expected

            if missing_pages:
                missing[name] = sorted(missing_pages)

            if extra_pages:
                extra[name] = sorted(extra_pages)

        if DEBUG and i % 50 == 0:
            print(f"[PROGRESS] {i}/{total}")

    return mismatches, missing, extra


def main():
    print("\n=== VERIFY INDEX WITH WORD COM ===\n")

    with open(FINAL_JSON, "r", encoding="utf-8") as f:
        index = json.load(f)

    word, doc = open_word(DOCX_INPUT)

    try:
        mismatches, missing, extra = verify(doc, index)

    finally:
        close_word(word, doc)

    print("\n=== RESULTS ===\n")

    print(f"Total mismatches: {len(mismatches)}")

    if mismatches:
        print("\n--- MISMATCHES ---")
        for k, v in list(mismatches.items())[:20]:
            print(f"{k}")
            print(f"  expected: {v['expected']}")
            print(f"  found:    {v['found']}")

    if missing:
        print("\n--- MISSING PAGES ---")
        for k, v in list(missing.items())[:20]:
            print(f"{k}: {v}")

    if extra:
        print("\n--- EXTRA PAGES ---")
        for k, v in list(extra.items())[:20]:
            print(f"{k}: {v}")

    print("\n✅ Verification complete\n")


if __name__ == "__main__":
    main()