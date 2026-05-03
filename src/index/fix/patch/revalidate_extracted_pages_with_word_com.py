"""
===============================================================================
FILE: revalidate_extracted_pages_with_word_com.py
LOCATION: src/index/fix/patch/
-------------------------------------------------------------------------------
Revalidate pages for canonical extracted index entries using Word COM.

Primary goal:
- Canonical entry list comes from index_curated_extracted.json
- Recalculate pages from book DOCX (Word COM truth)
- Emit explicit not-found list for entries with zero Word hits

Secondary goal:
- Provide raw/curated lookup diagnostics and optional fallback pages

OUTPUTS:
- data/index/intermediate/index_curated_extracted_word_fixed.json
- data/index/output/index_extracted_word_revalidation_report.csv
- data/index/output/index_extracted_not_found_word.csv
- data/index/output/index_extracted_not_found_word.json
===============================================================================
"""

import argparse
import csv
import json
import re
from pathlib import Path

import win32com.client as win32


def find_project_root():
    current = Path(__file__).resolve()
    for parent in [current] + list(current.parents):
        if (parent / "src").exists() and (parent / "data").exists():
            return parent
    raise RuntimeError("Project root not found")


BASE_DIR = find_project_root()

BOOK_DOCX = BASE_DIR / "data/index/input/book/HITS AND HAPPINESS FINAL 2 Format MOM Discog.docx"
EXTRACTED_JSON = BASE_DIR / "data/index/intermediate/index_curated_extracted.json"
RAW_JSON = BASE_DIR / "data/index/intermediate/index_raw_fixed.json"
CURATED_JSON = BASE_DIR / "data/index/intermediate/index_curated_final.json"

OUTPUT_JSON = BASE_DIR / "data/index/intermediate/index_curated_extracted_word_fixed.json"
OUTPUT_REPORT_CSV = BASE_DIR / "data/index/output/index_extracted_word_revalidation_report.csv"
OUTPUT_REPORT_JSON = BASE_DIR / "data/index/output/index_extracted_word_revalidation_report.json"
OUTPUT_NOT_FOUND_CSV = BASE_DIR / "data/index/output/index_extracted_not_found_word.csv"
OUTPUT_NOT_FOUND_JSON = BASE_DIR / "data/index/output/index_extracted_not_found_word.json"
OUTPUT_FIXED_CSV = BASE_DIR / "data/index/output/index_curated_extracted_fixed.csv"
OUTPUT_NOT_FIXED_CSV = BASE_DIR / "data/index/output/index_curated_extracted_not_fixed.csv"
OUTPUT_WORD_FIXED_REPORT_CSV = BASE_DIR / "data/index/output/index_extracted_word_fixed_report.csv"
OUTPUT_WORD_FIXED_REPORT_JSON = BASE_DIR / "data/index/output/index_extracted_word_fixed_report.json"


def normalize(text):
    text = (text or "").lower().strip()
    text = re.sub(r"^the\s+", "", text)
    text = re.sub(r"[^\w\s]", "", text)
    return re.sub(r"\s+", " ", text).strip()


def reverse_comma_name(name):
    if "," not in name:
        return None
    parts = [p.strip() for p in name.split(",") if p.strip()]
    if len(parts) < 2:
        return None
    return " ".join(parts[1:] + [parts[0]])


def comma_name(name):
    parts = [p.strip() for p in name.split() if p.strip()]
    if len(parts) < 2:
        return None
    return f"{parts[-1]}, {' '.join(parts[:-1])}"


def build_lookup(data):
    lookup = {}
    for key, value in data.items():
        key_norm = normalize(key)
        if key_norm and key_norm not in lookup:
            lookup[key_norm] = value

        normalized = normalize(value.get("normalized", ""))
        if normalized and normalized not in lookup:
            lookup[normalized] = value
    return lookup


def open_word(path):
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(str(path))
    return word, doc


def close_word(word, doc):
    doc.Close(False)
    word.Quit()


def find_pages(doc, text):
    if not text or not text.strip():
        return set()

    pages = set()
    pattern = re.compile(rf"\b{re.escape(text)}\b", re.IGNORECASE)
    full_text = doc.Content.Text

    for match in pattern.finditer(full_text):
        try:
            rng = doc.Range(Start=match.start(), End=match.end())
            page = int(rng.Information(1))
            pages.add(page)
        except Exception:
            continue

    return pages


def unique_terms(terms):
    out = []
    seen = set()
    for term in terms:
        term = (term or "").strip()
        if not term:
            continue
        key = term.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(term)
    return out


def build_search_terms(key, entry):
    normalized = entry.get("normalized", "")
    aliases = entry.get("aliases", [])
    aliases_external = entry.get("aliases_external", [])

    terms = [key, normalized]
    terms.extend(aliases)
    terms.extend(aliases_external)

    rev_key = reverse_comma_name(key)
    if rev_key:
        terms.append(rev_key)

    rev_norm = reverse_comma_name(normalized)
    if rev_norm:
        terms.append(rev_norm)

    comma_key = comma_name(key)
    if comma_key:
        terms.append(comma_key)

    comma_norm = comma_name(normalized)
    if comma_norm:
        terms.append(comma_norm)

    return unique_terms(terms)


def parse_args():
    parser = argparse.ArgumentParser(
        description="Revalidate extracted index pages using Word COM."
    )
    parser.add_argument(
        "--no-fallback",
        action="store_true",
        help="Do not fallback to raw/curated pages when Word finds no pages.",
    )
    return parser.parse_args()


def write_csv(path, rows, fields):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)


def compress_pages(pages):
    if not pages:
        return ""

    ordered = sorted(set(int(p) for p in pages))
    ranges = []
    start = ordered[0]
    prev = ordered[0]

    for page in ordered[1:]:
        if page == prev + 1:
            prev = page
            continue

        if start == prev:
            ranges.append(str(start))
        else:
            ranges.append(f"{start}-{prev}")

        start = page
        prev = page

    if start == prev:
        ranges.append(str(start))
    else:
        ranges.append(f"{start}-{prev}")

    return ", ".join(ranges)


def main():
    args = parse_args()

    print("\n=== REVALIDATE EXTRACTED PAGES WITH WORD COM ===\n")
    print(f"BOOK DOCX:  {BOOK_DOCX}")
    print(f"EXTRACTED:  {EXTRACTED_JSON}")
    print(f"RAW:        {RAW_JSON}")
    print(f"CURATED:    {CURATED_JSON}\n")

    if not BOOK_DOCX.exists():
        raise FileNotFoundError(f"Missing book DOCX: {BOOK_DOCX}")
    if not EXTRACTED_JSON.exists():
        raise FileNotFoundError(f"Missing extracted JSON: {EXTRACTED_JSON}")

    with open(EXTRACTED_JSON, "r", encoding="utf-8") as f:
        extracted = json.load(f)

    raw = {}
    if RAW_JSON.exists():
        with open(RAW_JSON, "r", encoding="utf-8") as f:
            raw = json.load(f)

    curated = {}
    if CURATED_JSON.exists():
        with open(CURATED_JSON, "r", encoding="utf-8") as f:
            curated = json.load(f)

    raw_lookup = build_lookup(raw)
    curated_lookup = build_lookup(curated)

    word, doc = open_word(BOOK_DOCX)
    try:
        print("Repaginating Word document...")
        doc.Repaginate()
        print("Pagination ready.\n")

        fixed = {}
        report_rows = []
        not_found_rows = []
        fixed_rows = []
        not_fixed_rows = []
        word_fixed_report_rows = []

        total = len(extracted)
        for idx, (key, entry) in enumerate(extracted.items(), 1):
            pages_extracted = entry.get("pages", [])
            search_terms = build_search_terms(key, entry)

            pages_word = set()
            for term in search_terms:
                pages_word.update(find_pages(doc, term))
            pages_word = sorted(pages_word)

            key_norm = normalize(key)
            raw_entry = raw_lookup.get(key_norm)
            curated_entry = curated_lookup.get(key_norm)

            pages_raw = raw_entry.get("pages", []) if raw_entry else []
            pages_curated = curated_entry.get("pages", []) if curated_entry else []

            if pages_word:
                pages_final = pages_word
                source = "WORD"
            elif not args.no_fallback and pages_raw:
                pages_final = pages_raw
                source = "RAW_FALLBACK"
            elif not args.no_fallback and pages_curated:
                pages_final = pages_curated
                source = "CURATED_FALLBACK"
            else:
                pages_final = pages_extracted
                source = "EXTRACTED_ORIGINAL"

            new_entry = dict(entry)
            new_entry["pages"] = pages_final
            new_entry["page_source"] = source
            fixed[key] = new_entry

            not_found_word = not pages_word

            fixed_rows.append(
                {
                    "key": key,
                    "old_pages": ", ".join(map(str, pages_extracted)),
                    "new_pages": ", ".join(map(str, pages_final)),
                }
            )
            word_fixed_report_rows.append(
                {
                    "key": key,
                    "word": compress_pages(pages_final),
                    "old_pages": ", ".join(map(str, pages_extracted)),
                    "new_pages": ", ".join(map(str, pages_final)),
                    "source": source,
                }
            )

            report_rows.append(
                {
                    "key": key,
                    "normalized": entry.get("normalized", ""),
                    "source": source,
                    "not_found_word": "1" if not_found_word else "0",
                    "pages_extracted": ", ".join(map(str, pages_extracted)),
                    "pages_word": ", ".join(map(str, pages_word)),
                    "pages_raw": ", ".join(map(str, pages_raw)),
                    "pages_curated": ", ".join(map(str, pages_curated)),
                    "pages_final": ", ".join(map(str, pages_final)),
                    "search_terms": " | ".join(search_terms),
                }
            )

            if not_found_word:
                not_found_rows.append(
                    {
                        "key": key,
                        "normalized": entry.get("normalized", ""),
                        "pages_extracted": ", ".join(map(str, pages_extracted)),
                        "pages_raw": ", ".join(map(str, pages_raw)),
                        "pages_curated": ", ".join(map(str, pages_curated)),
                        "fallback_source": source,
                        "search_terms": " | ".join(search_terms),
                    }
                )
                not_fixed_rows.append(
                    {
                        "key": key,
                        "old_pages": ", ".join(map(str, pages_extracted)),
                        "new_pages": "",
                    }
                )

            if idx % 25 == 0:
                print(f"[PROGRESS] {idx}/{total}")

    finally:
        close_word(word, doc)

    OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(fixed, f, indent=2)

    write_csv(
        OUTPUT_REPORT_CSV,
        report_rows,
        [
            "key",
            "normalized",
            "source",
            "not_found_word",
            "pages_extracted",
            "pages_word",
            "pages_raw",
            "pages_curated",
            "pages_final",
            "search_terms",
        ],
    )
    with open(OUTPUT_REPORT_JSON, "w", encoding="utf-8") as f:
        json.dump(report_rows, f, indent=2)

    write_csv(
        OUTPUT_NOT_FOUND_CSV,
        not_found_rows,
        [
            "key",
            "normalized",
            "pages_extracted",
            "pages_raw",
            "pages_curated",
            "fallback_source",
            "search_terms",
        ],
    )

    with open(OUTPUT_NOT_FOUND_JSON, "w", encoding="utf-8") as f:
        json.dump(not_found_rows, f, indent=2)

    write_csv(
        OUTPUT_FIXED_CSV,
        fixed_rows,
        ["key", "old_pages", "new_pages"],
    )

    write_csv(
        OUTPUT_NOT_FIXED_CSV,
        not_fixed_rows,
        ["key", "old_pages", "new_pages"],
    )
    write_csv(
        OUTPUT_WORD_FIXED_REPORT_CSV,
        word_fixed_report_rows,
        ["key", "word", "old_pages", "new_pages", "source"],
    )
    with open(OUTPUT_WORD_FIXED_REPORT_JSON, "w", encoding="utf-8") as f:
        json.dump(word_fixed_report_rows, f, indent=2)

    total = len(report_rows)
    not_found_count = len(not_found_rows)
    word_count = sum(1 for row in report_rows if row["source"] == "WORD")
    fallback_count = total - word_count

    print("\n=== SUMMARY ===\n")
    print(f"Entries processed:   {total}")
    print(f"Word matches:        {word_count}")
    print(f"Fallback used:       {fallback_count}")
    print(f"Not found in Word:   {not_found_count}")
    print(f"\nSaved JSON:          {OUTPUT_JSON}")
    print(f"Saved report CSV:    {OUTPUT_REPORT_CSV}")
    print(f"Saved report JSON:   {OUTPUT_REPORT_JSON}")
    print(f"Saved not-found CSV: {OUTPUT_NOT_FOUND_CSV}")
    print(f"Saved not-found JSON:{OUTPUT_NOT_FOUND_JSON}")
    print(f"Saved fixed CSV:     {OUTPUT_FIXED_CSV}")
    print(f"Saved not-fixed CSV: {OUTPUT_NOT_FIXED_CSV}")
    print(f"Saved word-fixed CSV:{OUTPUT_WORD_FIXED_REPORT_CSV}")
    print(f"Saved word-fixed JSON:{OUTPUT_WORD_FIXED_REPORT_JSON}")
    print("\nDone.\n")


if __name__ == "__main__":
    main()
