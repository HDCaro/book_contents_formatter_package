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

    # MERGE
    for k, r in curated.items():
        if r.get("action") == "merge":
            target = r.get("target")
            if k in final:
                final.setdefault(target, {
                    "normalized": target,
                    "type": final[k]["type"],
                    "pages": set()
                })
                final[target]["pages"].update(final[k]["pages"])
                del final[k]

    # UPDATE
    for k, r in curated.items():
        if r.get("action") == "update" and k in final:
            if "normalized" in r:
                final[k]["normalized"] = r["normalized"]
            if "type" in r:
                final[k]["type"] = r["type"]

    return final

# ---------------- ALIAS GROUP ---------------- #

def apply_alias_groups(final, curated):
    for key, rule in curated.items():
        if rule.get("action") == "alias_group":
            aliases = rule.get("aliases", [])

            final.setdefault(key, {
                "normalized": rule.get("normalized", key),
                "type": "person",
                "pages": set()
            })

            for alias in aliases:
                if alias in final:
                    final[key]["pages"].update(final[alias]["pages"])
                    del final[alias]

    return final

# ---------------- ADD ENRICH ---------------- #

def enrich_add_entries(pages, curated, final):
    for key, rule in curated.items():
        if rule.get("action") == "add":
            pattern = re.compile(rf"\b{re.escape(key)}\b", re.IGNORECASE)
            pages_found = set()

            for i, page in enumerate(pages, 1):
                if pattern.search(page):
                    pages_found.add(i)

            final[key] = {
                "normalized": rule["normalized"],
                "type": rule.get("type", "unknown"),
                "pages": pages_found
            }

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

def build_alpha(final, curated):
    grouped = {}

    for key, v in final.items():
        if not v["pages"]:
            continue

        name = v["normalized"]
        pages = compress(list(v["pages"]))

        rule = curated.get(key, {})
        aliases = rule.get("aliases")

        letter = name[0].upper()

        if aliases:
            grouped.setdefault(letter, []).append({
                "type": "alias",
                "name": name,
                "aliases": aliases,
                "pages": pages
            })
        else:
            grouped.setdefault(letter, []).append({
                "type": "normal",
                "line": f"{name}, {pages}"
            })

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

def apply_index_spacing(p):
    fmt = p.paragraph_format
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    fmt.line_spacing = 1

def create_doc(alpha):
    doc = Document()

    # Title
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

        # tight but readable spacing
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(1)

        for entry in entries:

            if entry["type"] == "alias":
                p1 = doc.add_paragraph(entry["name"])
                p1.paragraph_format.left_indent = Inches(0.3)
                p1.paragraph_format.first_line_indent = Inches(-0.3)
                apply_index_spacing(p1)

                aka = doc.add_paragraph("(AKA: " + ", ".join(entry["aliases"]) + ")")
                aka.paragraph_format.left_indent = Inches(0.3)
                apply_index_spacing(aka)

                p2 = doc.add_paragraph(entry["pages"])
                p2.paragraph_format.left_indent = Inches(0.3)
                p2.paragraph_format.first_line_indent = Inches(-0.3)
                apply_index_spacing(p2)

            else:
                p = doc.add_paragraph(entry["line"])
                p.paragraph_format.left_indent = Inches(0.3)
                p.paragraph_format.first_line_indent = Inches(-0.3)
                apply_index_spacing(p)

    try:
        doc.save(DOCX_OUTPUT)
    except PermissionError:
        print("\n❌ Cannot save file: it is open.")
        print(f"👉 {DOCX_OUTPUT}")
        sys.exit(1)

# ---------------- MAIN ---------------- #

def main():
    print("\n=== BUILD INDEX DOCX ===\n")

    try:
        with open(RAW_JSON, "r", encoding="utf-8") as f:
            raw = json.load(f)
    except:
        print(f"❌ Missing file: {RAW_JSON}")
        sys.exit(1)

    try:
        with open(CURATED_JSON, "r", encoding="utf-8") as f:
            curated = json.load(f)
    except:
        print(f"❌ Missing file: {CURATED_JSON}")
        sys.exit(1)

    final = apply_curation(raw, curated)
    final = apply_alias_groups(final, curated)

    text = extract_text(DOCX_INPUT)
    pages = split_pages(text)
    enrich_add_entries(pages, curated, final)

    alpha = build_alpha(final, curated)
    create_doc(alpha)

    print(f"\n✅ Index created: {DOCX_OUTPUT}")

if __name__ == "__main__":
    main()