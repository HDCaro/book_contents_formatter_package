"""
===============================================================================
FILE: build_index_docx.py
-------------------------------------------------------------------------------

PURPOSE
-------
Builds the final book index (DOCX) from processed JSON data.

- Applies curation rules (remove, update, merge, merged/destination)
- Aggregates pages across merged entries
- Prevents alias duplication (same normalized identity)
- Generates:
    • index_curated_final.json
    • Final formatted DOCX index (2 columns, soft breaks)

IMPORTANT
---------
This file DOES NOT scan the document.
That is handled by:
    generate_index_data_with_COM.py

===============================================================================
"""

import json
import re
import sys
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.shared import Inches, Pt
from docx.opc.exceptions import PackageNotFoundError
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ---------------- CONFIG ---------------- #

DOCX_INPUT = "HITS AND HAPPINESS FINAL 2 Format MOM Discog.docx"
RAW_JSON = "index_raw.json"
CURATED_JSON = "index_curated.json"
FINAL_JSON = "index_curated_final.json"

DOCX_OUTPUT = Path(DOCX_INPUT).with_name(
    Path(DOCX_INPUT).stem + "-index.docx"
)

WORDS_PER_PAGE = 600

# ---------------- VALIDATION ---------------- #

def validate_inputs(raw, curated):
    print("\n🔍 Validating input files...\n")

    sample_raw = next(iter(raw.values()), {})
    if "normalized" not in sample_raw or "pages" not in sample_raw:
        print("❌ RAW JSON invalid")
        sys.exit(1)

    sample_cur = next(iter(curated.values()), {})
    if "action" not in sample_cur:
        print("❌ CURATED JSON invalid")
        sys.exit(1)

    print("✅ Input validation passed\n")

def validate_final_json(data):
    print("\n🔍 Validating final JSON...\n")

    sample = next(iter(data.values()), {})
    if "pages" not in sample:
        print("❌ FINAL JSON invalid")
        sys.exit(1)

    print("✅ Final JSON structure OK\n")

def validate_page_override(key, raw_pages, curated_pages):
    raw_set = set(raw_pages)
    curated_set = set(curated_pages)

    removed = sorted(raw_set - curated_set)
    added = sorted(curated_set - raw_set)

    if removed:
        print(f"\n⚠️ Page reduction detected: {key} → {removed}")

    if added:
        print(f"\n➕ Page addition detected: {key} → {added}")

# ---------------- TEXT ---------------- #

def extract_text(docx):
    try:
        doc = Document(docx)
        return "\n".join(p.text for p in doc.paragraphs)
    except PackageNotFoundError:
        print(f"❌ File not found: {docx}")
        sys.exit(1)

def split_pages(text):
    words = text.split()
    return [
        " ".join(words[i:i + WORDS_PER_PAGE])
        for i in range(0, len(words), WORDS_PER_PAGE)
    ]

# ---------------- CURATION ---------------- #

def apply_curation(raw, curated):
    final = deepcopy(raw)

    for v in final.values():
        v["pages"] = set(v["pages"])

    # REMOVE
    for k, r in curated.items():
        if r.get("action") == "remove" and k in final:
            del final[k]

    # UPDATE
    for k, r in curated.items():
        if r.get("action") == "update" and k in final:
            if "normalized" in r:
                final[k]["normalized"] = r["normalized"]

            if "type" in r:
                final[k]["type"] = r["type"]

            if "pages" in r and r["pages"]:
                validate_page_override(k, final[k]["pages"], r["pages"])
                final[k]["pages"] = set(r["pages"])

    # MERGE TARGETS
    for k, r in curated.items():
        if r.get("action") == "merge" and k not in final:
            final[k] = {
                "normalized": r.get("normalized", k),
                "type": r.get("type", "unknown"),
                "pages": set()
            }

    # MERGED → DESTINATION
    for k, r in curated.items():
        if r.get("action") == "merged":
            dest = r.get("destination")

            if not dest:
                print(f"⚠️ Missing destination: {k}")
                continue

            if dest not in final:
                final[dest] = {
                    "normalized": dest,
                    "type": "unknown",
                    "pages": set()
                }

            if k in final:
                final[dest]["pages"].update(final[k]["pages"])
                del final[k]

    return final

# ---------------- ADD ENRICH ---------------- #

def enrich_add_entries(pages, curated, final):
    for key, rule in curated.items():
        if rule.get("action") == "add":
            pattern = re.compile(rf"\b{re.escape(key)}\b", re.IGNORECASE)
            found = set()

            for i, page in enumerate(pages, 1):
                if pattern.search(page):
                    found.add(i)

            final[key] = {
                "normalized": rule["normalized"],
                "type": rule.get("type", "unknown"),
                "pages": found
            }

# ---------------- NORMALIZED INDEX ---------------- #

def build_normalized_index(final, curated):
    normalized_index = {}

    for key, v in final.items():
        norm = v["normalized"]

        normalized_index.setdefault(norm, {
            "pages": set(),
            "aliases": set()
        })

        normalized_index[norm]["pages"].update(v["pages"])

    # ---------- ALIAS LOGIC (FIXED) ---------- #

    for ck, rule in curated.items():
        if rule.get("action") == "merged":
            dest = rule.get("destination")

            if not dest or dest not in final or ck not in final:
                continue

            dest_norm = final[dest]["normalized"]
            source_norm = final.get(ck, {}).get("normalized")

            # 🚫 Prevent self-alias (same normalized identity)
            if source_norm == dest_norm:
                continue

            normalized_index[dest_norm]["aliases"].add(ck)

    # manual aliases
    for key, rule in curated.items():
        if key in final:
            norm = final[key]["normalized"]
            for alias in rule.get("aliases", []):
                normalized_index[norm]["aliases"].add(alias)

    return normalized_index

def sort_normalized_index(index):
    return dict(sorted(index.items(), key=lambda x: x[0].lower()))

# ---------------- SAVE ---------------- #

def save_final_json(index):
    output = {}

    for norm, v in index.items():
        if not v["pages"]:
            continue

        entry = {
            "pages": sorted(v["pages"])
        }

        if v["aliases"]:
            entry["aliases"] = sorted(v["aliases"])

        output[norm] = entry

    with open(FINAL_JSON, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2)

    print(f"💾 Saved: {FINAL_JSON}")

# ---------------- LOAD ---------------- #

def load_final_json():
    with open(FINAL_JSON, "r", encoding="utf-8") as f:
        data = json.load(f)

    for v in data.values():
        v["pages"] = set(v["pages"])

    validate_final_json(data)

    return data

# ---------------- INDEX BUILD ---------------- #

def compress(pages):
    if not pages:
        return ""

    pages = sorted(pages)
    ranges = []
    start = prev = pages[0]

    for p in pages[1:]:
        if p == prev + 1:
            prev = p
        else:
            ranges.append((start, prev))
            start = prev = p

    ranges.append((start, prev))

    return ", ".join(
        f"{a}–{b}" if a != b else str(a)
        for a, b in ranges
    )

def build_alpha(index):
    grouped = {}

    for name, v in index.items():
        if not v["pages"]:
            continue

        pages = compress(list(v["pages"]))
        letter = name[0].upper()

        if v.get("aliases"):
            grouped.setdefault(letter, []).append({
                "type": "alias",
                "name": name,
                "aliases": ", ".join(sorted(v["aliases"])),
                "pages": pages
            })
        else:
            grouped.setdefault(letter, []).append({
                "type": "normal",
                "line": f"{name}, {pages}"
            })

    return dict(
        sorted(
            (k, sorted(v, key=lambda x: x.get("name", x.get("line"))))
            for k, v in grouped.items()
        )
    )

# ---------------- DOCX ---------------- #

def set_two_columns(doc):
    section = doc.sections[0]
    sectPr = section._sectPr

    cols = sectPr.xpath('./w:cols')
    if cols:
        cols[0].set(qn('w:num'), "2")
    else:
        cols = OxmlElement('w:cols')
        cols.set(qn('w:num'), "2")
        sectPr.append(cols)

def create_doc(alpha):
    doc = Document()

    style = doc.styles['Normal']
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(0)

    title = doc.add_paragraph()
    run = title.add_run("INDEX")
    run.bold = True
    run.font.size = Pt(20)
    title.alignment = 1

    set_two_columns(doc)

    for letter, entries in alpha.items():

        p = doc.add_paragraph()

        fmt = p.paragraph_format
        fmt.space_before = Pt(6)
        fmt.space_after = Pt(2)
        fmt.left_indent = Inches(0.3)
        fmt.first_line_indent = Inches(-0.3)

        run = p.add_run(letter)
        run.bold = True
        run.font.size = Pt(14)
        run.add_break()

        for entry in entries:

            if entry["type"] == "alias":
                p.add_run(entry["name"])
                p.add_run().add_break()

                p.add_run(f"(AKA: {entry['aliases']})")
                p.add_run().add_break()

                p.add_run(entry["pages"])
                p.add_run().add_break()
            else:
                p.add_run(entry["line"])
                p.add_run().add_break()

    doc.save(DOCX_OUTPUT)

# ---------------- MAIN ---------------- #

def main():
    print("\n=== BUILD INDEX PIPELINE ===\n")

    raw = json.load(open(RAW_JSON))
    curated = json.load(open(CURATED_JSON))

    validate_inputs(raw, curated)

    final = apply_curation(raw, curated)

    text = extract_text(DOCX_INPUT)
    pages = split_pages(text)

    enrich_add_entries(pages, curated, final)

    normalized_index = build_normalized_index(final, curated)
    normalized_index = sort_normalized_index(normalized_index)

    save_final_json(normalized_index)

    final = load_final_json()

    alpha = build_alpha(final)
    create_doc(alpha)

    print(f"\n✅ Index created: {DOCX_OUTPUT}")

if __name__ == "__main__":
    main()