"""
===============================================================================
FILE: build_index_docx.py
-------------------------------------------------------------------------------
Builds final DOCX index from raw + curated JSON.

Includes:
- merge logic
- alias + aliases_external
- normalization override
- filtering (MIN_PAGES)
- filtered entries audit file
- DEBUG mode

===============================================================================
"""

import json
import sys
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ---------------- CONFIG ---------------- #

DOCX_INPUT = "HITS AND HAPPINESS FINAL 2 Format MOM Discog.docx"
RAW_JSON = "index_raw.json"
CURATED_JSON = "index_curated.json"
FINAL_JSON = "index_curated_final.json"
FILTERED_JSON = "index_filtered_out.json"

DEBUG_INDEX = True
MIN_PAGES = 2

DOCX_OUTPUT = Path(DOCX_INPUT).with_name(
    Path(DOCX_INPUT).stem + "-index.docx"
)

# ---------------- VALIDATION ---------------- #

def validate_inputs(raw, curated):
    if not raw or not curated:
        print("❌ Invalid input files")
        sys.exit(1)

# ---------------- CURATION ---------------- #

def apply_curation(raw, curated):
    final = deepcopy(raw)

    for v in final.values():
        v["pages"] = set(v["pages"])

    # REMOVE
    for k, r in curated.items():
        if r.get("action") == "remove" and k in final:
            del final[k]

    # MERGED
    for k, r in curated.items():
        if r.get("action") == "merged":
            dest = r.get("destination")

            if dest not in final:
                final[dest] = {
                    "normalized": dest,
                    "type": "unknown",
                    "pages": set()
                }

            if k in final:
                final[dest]["pages"].update(final[k]["pages"])
                del final[k]

    # FORCE normalization
    for k, r in curated.items():
        if k in final and "normalized" in r:
            final[k]["normalized"] = r["normalized"]

    return final

# ---------------- NORMALIZED INDEX ---------------- #

def build_normalized_index(final, curated):
    normalized_index = {}

    for key, v in final.items():
        norm = v["normalized"]

        normalized_index.setdefault(norm, {
            "pages": set(),
            "aliases": set(),
            "aliases_external": set()
        })

        normalized_index[norm]["pages"].update(v["pages"])

    # merged aliases
    for ck, rule in curated.items():
        if rule.get("action") == "merged":
            dest = rule.get("destination")

            if not dest or dest not in final:
                continue

            dest_norm = final[dest]["normalized"]
            source_norm = rule.get("normalized")

            if source_norm == dest_norm:
                continue

            normalized_index[dest_norm]["aliases"].add(ck)

    # manual aliases
    for key, rule in curated.items():
        if key in final:
            norm = final[key]["normalized"]

            for a in rule.get("aliases", []):
                normalized_index[norm]["aliases"].add(a)

            for a in rule.get("aliases_external", []):
                normalized_index[norm]["aliases_external"].add(a)

    if DEBUG_INDEX:
        print("\n=== DEBUG: NORMALIZED INDEX ===\n")
        for k, v in normalized_index.items():
            print(f"{k}")
            print(f"  pages: {sorted(v['pages'])}")
            print(f"  aliases: {sorted(v['aliases'])}")
            print(f"  external: {sorted(v['aliases_external'])}")
            print()

    return normalized_index

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

        if v["aliases_external"]:
            entry["aliases_external"] = sorted(v["aliases_external"])

        output[norm] = entry

    with open(FINAL_JSON, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2)

# ---------------- FILTER ---------------- #

def compress(pages):
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
    filtered_out = {}

    for name, v in index.items():
        if not v["pages"]:
            continue

        # 🔥 FILTER
        if len(v["pages"]) < MIN_PAGES:
            filtered_out[name] = {
                "pages": sorted(v["pages"]),
                "aliases": sorted(v["aliases"]),
                "aliases_external": sorted(v["aliases_external"])
            }

            if DEBUG_INDEX:
                print(f"[FILTERED] {name} → {v['pages']}")

            continue

        pages = compress(v["pages"])
        letter = name[0].upper()

        if v.get("aliases") or v.get("aliases_external"):
            grouped.setdefault(letter, []).append({
                "type": "alias",
                "name": name,
                "aliases": ", ".join(sorted(v["aliases"])) if v["aliases"] else "",
                "aliases_external": ", ".join(sorted(v["aliases_external"])) if v["aliases_external"] else "",
                "pages": pages
            })
        else:
            grouped.setdefault(letter, []).append({
                "type": "normal",
                "line": f"{name}, {pages}"
            })

    # save filtered entries
    with open(FILTERED_JSON, "w", encoding="utf-8") as f:
        json.dump(filtered_out, f, indent=2)

    print(f"\n💾 Filtered entries saved → {FILTERED_JSON}")

    return dict(sorted(grouped.items()))

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

    title = doc.add_paragraph()
    run = title.add_run("INDEX")
    run.bold = True
    run.font.size = Pt(20)
    title.alignment = 1

    set_two_columns(doc)

    for letter, entries in alpha.items():
        p = doc.add_paragraph()

        run = p.add_run(letter)
        run.bold = True
        run.font.size = Pt(14)
        run.add_break()

        for entry in entries:

            if entry["type"] == "alias":
                p.add_run(entry["name"])
                p.add_run().add_break()

                if entry.get("aliases"):
                    p.add_run(f"(AKA: {entry['aliases']})")
                    p.add_run().add_break()

                if entry.get("aliases_external"):
                    p.add_run(f"(Also known as: {entry['aliases_external']})")
                    p.add_run().add_break()

                p.add_run(entry["pages"])
                p.add_run().add_break()

            else:
                p.add_run(entry["line"])
                p.add_run().add_break()

    doc.save(DOCX_OUTPUT)

# ---------------- MAIN ---------------- #

def main():
    raw = json.load(open(RAW_JSON))
    curated = json.load(open(CURATED_JSON))

    validate_inputs(raw, curated)

    final = apply_curation(raw, curated)
    index = build_normalized_index(final, curated)

    save_final_json(index)

    alpha = build_alpha(index)
    create_doc(alpha)

    print("\n✅ Index created:", DOCX_OUTPUT)

if __name__ == "__main__":
    main()