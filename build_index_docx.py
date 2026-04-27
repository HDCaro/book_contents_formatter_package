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

    # MERGE
    for k, r in curated.items():
        if r.get("action") == "merge":
            target = r.get("target")
            if k in final and target in final:
                final[target]["pages"].update(final[k]["pages"])
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

# ---------------- CANONICAL BUILD ---------------- #

def build_normalized_index(final, curated):
    """
    Output:
    {
      normalized_name: {
        "pages": set(...),
        "aliases": set(...)
      }
    }
    """
    normalized_index = {}

    for key, v in final.items():
        norm = v["normalized"]

        normalized_index.setdefault(norm, {
            "pages": set(),
            "aliases": set()
        })

        # merge pages
        normalized_index[norm]["pages"].update(v["pages"])

        # original key is also an alias (if different)
        if key != norm:
            normalized_index[norm]["aliases"].add(key)

        # curated aliases
        rule = curated.get(key, {})
        for alias in rule.get("aliases", []):
            if alias != norm:
                normalized_index[norm]["aliases"].add(alias)

    return normalized_index

def sort_normalized_index(index):
    return dict(sorted(index.items(), key=lambda x: x[0].lower()))

# ---------------- SAVE FINAL JSON ---------------- #

def save_final_json(index):
    output = {}

    for norm, v in index.items():
        if not v["pages"]:
            continue

        entry = {
            "aliases": sorted(v["aliases"]),
            "pages": sorted(v["pages"])
        }

        output[norm] = entry

    with open(FINAL_JSON, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2)

    print(f"💾 Saved: {FINAL_JSON}")

# ---------------- LOAD ---------------- #

def load_final_json():
    try:
        with open(FINAL_JSON, "r", encoding="utf-8") as f:
            data = json.load(f)
    except:
        print(f"❌ Missing file: {FINAL_JSON}")
        sys.exit(1)

    for v in data.values():
        v["pages"] = set(v["pages"])

    return data

# ---------------- INDEX ---------------- #

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

    for name, v in index.items():
        if not v["pages"]:
            continue

        pages = compress(list(v["pages"]))
        letter = name[0].upper()

        if v["aliases"]:
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

def apply_spacing(p):
    fmt = p.paragraph_format
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    fmt.line_spacing = 1

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
        r = p.add_run(letter)
        r.bold = True
        r.font.size = Pt(14)

        for entry in entries:

            if entry["type"] == "alias":
                p1 = doc.add_paragraph(entry["name"])
                p1.paragraph_format.left_indent = Inches(0.3)
                p1.paragraph_format.first_line_indent = Inches(-0.3)
                apply_spacing(p1)

                aka = doc.add_paragraph(f"(AKA: {entry['aliases']})")
                aka.paragraph_format.left_indent = Inches(0.3)
                apply_spacing(aka)

                p2 = doc.add_paragraph(entry["pages"])
                p2.paragraph_format.left_indent = Inches(0.3)
                p2.paragraph_format.first_line_indent = Inches(-0.3)
                apply_spacing(p2)

            else:
                p = doc.add_paragraph(entry["line"])
                p.paragraph_format.left_indent = Inches(0.3)
                p.paragraph_format.first_line_indent = Inches(-0.3)
                apply_spacing(p)

    try:
        doc.save(DOCX_OUTPUT)
    except PermissionError:
        print("❌ Close DOCX before saving")
        sys.exit(1)

# ---------------- MAIN ---------------- #

def main():
    print("\n=== BUILD INDEX PIPELINE ===\n")

    raw = json.load(open(RAW_JSON))
    curated = json.load(open(CURATED_JSON))

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