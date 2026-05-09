#!/usr/bin/env python3

import win32com.client
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
import os
import json
from pathlib import Path

# -----------------------------------
# WORD CONSTANTS (define explicitly)
# -----------------------------------
# Page numbering constants
wdRestartContinuous = 0
wdRestartPage = 2
wdRestartSection = 1

# Page number style constants
wdPageNumberStyleLowercaseRoman = 14
wdPageNumberStyleUppercaseRoman = 13

# Alignment constants
wdAlignParagraphCenter = 1

# Field type constants
wdFieldPage = 33

# Break type constants
wdSectionBreakNextPage = 2
wdPageBreak = 7

# Header/Footer constants
wdHeaderFooterPrimary = 1


def find_project_root():
    current = Path(__file__).resolve()
    for parent in [current.parent] + list(current.parents):
        if (parent / "src").exists() and (parent / "data").exists():
            return parent
    raise RuntimeError("Project root not found (expected folders: src and data)")


def find_file_by_name(project_root, filename):
    target = filename.lower()
    matches = []

    for root, _, files in os.walk(project_root):
        for name in files:
            if name.lower() == target:
                matches.append(Path(root) / name)

    # Prefer shallower paths for deterministic selection when multiple exist.
    matches.sort(key=lambda p: (len(p.parts), str(p).lower()))
    return matches


def resolve_existing_path(project_root, label, preferred_relative_paths, filename):
    for rel in preferred_relative_paths:
        candidate = project_root / rel
        if candidate.exists():
            print(f"   ✅ Resolved {label}: {candidate.relative_to(project_root)}")
            return candidate

    matches = find_file_by_name(project_root, filename)
    if matches:
        resolved = matches[0]
        print(f"   ✅ Auto-discovered {label}: {resolved.relative_to(project_root)}")
        return resolved

    search_list = "\n      - " + "\n      - ".join(str(p) for p in preferred_relative_paths)
    raise FileNotFoundError(
        f"Could not find {label} ('{filename}') under project root: {project_root}\n"
        f"Checked preferred paths:{search_list}"
    )


def resolve_config_path(project_root, configured_path):
    path = Path(configured_path)
    if path.is_absolute():
        return path
    return project_root / path


def load_builder_config(project_root):
    config_path = project_root / "src" / "front_matter" / "builders" / "front_matter_builder.config.json"
    if not config_path.exists():
        raise FileNotFoundError(
            f"Configuration file not found: {config_path}\n"
            "Create it and set the input/output paths before running."
        )

    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)

    # Validate required sections
    required_sections = ["inputs", "outputs"]
    missing_sections = [s for s in required_sections if s not in config]
    if missing_sections:
        raise ValueError(f"Missing required config sections: {', '.join(missing_sections)}")

    inputs = config.get("inputs", {})
    outputs = config.get("outputs", {})
    options = config.get("options", {})

    # Validate required input keys
    required_inputs = ["title_file", "copyright_file", "book_body_file"]
    missing_inputs = [k for k in required_inputs if not str(inputs.get(k, "")).strip()]
    if missing_inputs:
        raise ValueError(f"Missing required input keys in config: {', '.join(missing_inputs)}")

    # Validate required output keys
    required_outputs = ["output_dir", "filename", "temp_dir"]
    missing_outputs = [k for k in required_outputs if not str(outputs.get(k, "")).strip()]
    if missing_outputs:
        raise ValueError(f"Missing required output keys in config: {', '.join(missing_outputs)}")

    auto_discover = bool(inputs.get("auto_discover_missing_inputs", True))

    # Resolve input paths
    title_file = resolve_config_path(project_root, inputs["title_file"])
    if not title_file.exists() and auto_discover:
        matches = find_file_by_name(project_root, title_file.name)
        if matches:
            title_file = matches[0]

    copyright_file = resolve_config_path(project_root, inputs["copyright_file"])
    if not copyright_file.exists() and auto_discover:
        matches = find_file_by_name(project_root, copyright_file.name)
        if matches:
            copyright_file = matches[0]

    book_body_file = resolve_config_path(project_root, inputs["book_body_file"])
    if not book_body_file.exists() and auto_discover:
        matches = find_file_by_name(project_root, book_body_file.name)
        if matches:
            book_body_file = matches[0]

    # Validate all input files exist
    missing_files = []
    if not title_file.exists():
        missing_files.append(f"title_file: {title_file}")
    if not copyright_file.exists():
        missing_files.append(f"copyright_file: {copyright_file}")
    if not book_body_file.exists():
        missing_files.append(f"book_body_file: {book_body_file}")
    if missing_files:
        raise FileNotFoundError("Configured input files not found:\n - " + "\n - ".join(missing_files))

    # Resolve output paths
    output_dir = resolve_config_path(project_root, outputs["output_dir"])
    temp_dir = resolve_config_path(project_root, outputs["temp_dir"])
    output_filename = outputs.get("filename", "front_matter.docx")
    metadata_filename = outputs.get("metadata_filename", "front_matter.config.json")

    return {
        "config_path": config_path,
        "title_file": title_file,
        "copyright_file": copyright_file,
        "book_body_file": book_body_file,
        "output_dir": output_dir,
        "output_filename": output_filename,
        "metadata_filename": metadata_filename,
        "temp_dir": temp_dir,
        "delete_temp_toc": bool(options.get("delete_temp_toc", True)),
        "apply_book_layout": bool(options.get("apply_book_layout", True)),
        "page_numbering_style": options.get("page_numbering_style", "roman_lowercase"),
    }


def pretty_path(project_root, path_value):
    path_value = Path(path_value)
    try:
        return str(path_value.relative_to(project_root))
    except ValueError:
        return str(path_value)


# -----------------------------------
# CLEAN TEXT (REMOVE ALL LINE BREAKS AND CLEAN)
# -----------------------------------
def clean_text(text):
    if not text:
        return ""
    # Remove control characters and clean up
    cleaned = "".join(
        c for c in text
        if c >= " " or c in ("\n", "\t")
    ).strip()

    # Remove trailing slashes and extra whitespace
    cleaned = cleaned.rstrip("/\\").strip()

    return cleaned


def clean_title_text(text):
    """
    Special cleaning for titles - removes ALL line breaks and normalizes spaces
    """
    if not text:
        return ""

    # Remove ALL types of line breaks and normalize spaces
    cleaned = text.replace("\r\n", " ").replace("\r", " ").replace("\n", " ")

    # Replace multiple spaces with single space
    import re
    cleaned = re.sub(r'\s+', ' ', cleaned)

    # Remove control characters except spaces
    cleaned = "".join(c for c in cleaned if c >= " ")

    # Remove trailing slashes and extra whitespace
    cleaned = cleaned.rstrip("/\\").strip()

    return cleaned


# -----------------------------------
# WORD APP
# -----------------------------------
def get_word():
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    return word


# -----------------------------------
# EXTRACT HEADINGS (COLLECT ALL CONSECUTIVE HEADING 2 TEXT)
# -----------------------------------
def extract_headings(doc_path):
    print("\n📖 Extracting headings with consecutive Heading 2 collection")

    word = get_word()
    doc = word.Documents.Open(os.path.abspath(doc_path))
    doc.Repaginate()

    headings = []
    total = doc.Paragraphs.Count

    i = 1
    while i <= total:
        try:
            para = doc.Paragraphs(i)
            text = clean_text(para.Range.Text)
            style = para.Style.NameLocal.lower()

            # Look for Heading 1 (Chapter numbers)
            if "heading 1" in style and text:
                print(f"\n   📍 Found Heading 1: '{text}'")

                # ---- Chapter with Heading 2 title(s) ----
                if text.lower().startswith("chapter"):
                    print(f"   🔍 Processing chapter: '{text}'")

                    chapter = text
                    chapter_title_parts = []
                    j = i + 1

                    # Look for ALL consecutive Heading 2 paragraphs
                    print(f"   🔍 Looking for consecutive Heading 2 titles starting from paragraph {j}...")

                    while j <= total:
                        try:
                            p = doc.Paragraphs(j)
                            p_style = p.Style.NameLocal.lower()

                            # Get RAW text from the entire paragraph range
                            raw_text = p.Range.Text
                            cleaned_text = clean_title_text(raw_text)

                            print(f"   📝 Paragraph {j}: '{cleaned_text}' (style: {p_style})")

                            if not raw_text or not raw_text.strip():
                                print(f"   ⏭️ Empty paragraph {j}, skipping")
                                j += 1
                                continue

                            # Found Heading 2 - collect this part of the title
                            if "heading 2" in p_style:
                                chapter_title_parts.append(cleaned_text)
                                print(f"   ✅ Collected Heading 2 part: '{cleaned_text}'")
                                j += 1
                                continue  # Keep looking for more Heading 2 parts

                            # If we hit another Heading 1, stop looking
                            elif "heading 1" in p_style:
                                print(f"   🛑 Hit another Heading 1, stopping search")
                                break

                            # If we hit any other style, stop looking for more title parts
                            else:
                                print(f"   🛑 Hit non-heading style, stopping title collection")
                                break

                        except Exception as e:
                            print(f"   ⚠️ Error processing paragraph {j}: {e}")
                            break

                    # Build the complete title from all parts
                    if chapter_title_parts:
                        # Join all title parts with a space
                        complete_title = " ".join(chapter_title_parts)
                        full = f"{chapter}: {complete_title}"
                        print(f"   🎉 Complete chapter title: '{full}'")
                    else:
                        full = chapter
                        print(f"   ⚠️ No Heading 2 found, using chapter only: '{full}'")

                    # Final cleaning to ensure single line
                    full = clean_title_text(full)
                    page = para.Range.Information(3)

                    headings.append({"text": full, "page": page})
                    print(f"   📌 Added to TOC: '{full}' → page {page}")

                    i = j
                    continue

                # ---- Single-line heading (Discography, Books, etc.) ----
                else:
                    page = para.Range.Information(3)
                    # Clean single-line headings too
                    clean_heading = clean_title_text(text)
                    headings.append({"text": clean_heading, "page": page})
                    print(f"   📌 Single-line heading: '{clean_heading}' → page {page}")

            # Also capture standalone Heading 2 (if not part of a chapter)
            elif "heading 2" in style and text:
                page = para.Range.Information(3)
                # Clean standalone Heading 2
                clean_heading = clean_title_text(para.Range.Text)
                headings.append({"text": clean_heading, "page": page})
                print(f"   📌 Standalone Heading 2: '{clean_heading}' → page {page}")

        except Exception as e:
            print(f"   ⚠️ Error processing paragraph {i}: {e}")

        i += 1

    doc.Close(False)
    word.Quit()

    print(f"\n   ✅ Total headings extracted: {len(headings)}")
    for idx, h in enumerate(headings, 1):
        print(f"   {idx}. {h['text']} → page {h['page']}")

    return headings


# -----------------------------------
# BUILD TOC (GUARANTEED SINGLE LINE)
# -----------------------------------
def build_toc_doc(headings, toc_path):
    print("\n📝 Building TOC with guaranteed single-line titles")

    doc = Document()

    # Title
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = title.add_run("CONTENTS")
    run.font.name = "Georgia"
    run.font.size = Pt(20)
    run.bold = True

    doc.add_paragraph()

    for h in headings:
        para = doc.add_paragraph()

        # Hanging indent
        para.paragraph_format.left_indent = Inches(0.5)
        para.paragraph_format.first_line_indent = Inches(-0.5)

        # Tab stop with dots
        para.paragraph_format.tab_stops.add_tab_stop(
            Inches(6),
            WD_TAB_ALIGNMENT.RIGHT,
            WD_TAB_LEADER.DOTS
        )

        # Get the title text and ensure it's single line
        title_text = h["text"]

        # Final safety check - remove any remaining line breaks
        title_text = title_text.replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
        title_text = " ".join(title_text.split())  # Normalize all whitespace

        print(f"   📝 TOC entry: '{title_text}' → {h['page']}")

        # Add title text (guaranteed single line)
        run = para.add_run(title_text)
        run.font.name = "Georgia"
        run.font.size = Pt(12)

        # Add tab and page number
        page_run = para.add_run(f"\t{h['page']}")
        page_run.bold = True

    doc.save(toc_path)
    print("   ✅ TOC created with guaranteed single-line entries")

    return toc_path


# -----------------------------------
# APPLY ROMAN PAGINATION (FIXED ROMAN FORMAT)
# -----------------------------------
def apply_roman_pagination(doc):
    print("\n🔢 Applying Roman pagination (FIXED FORMAT)")

    wdHeaderFooterPrimary = 1
    wdAlignParagraphCenter = 1
    wdPageNumberStyleLowercaseRoman = 14

    try:
        sections = doc.Sections
        print(f"   📊 Total sections: {sections.Count}")

        if sections.Count < 2:
            print("   ⚠️ Not enough sections")
            return

        # -----------------------------------
        # SECTION 1 → NO PAGE NUMBERS (TITLE)
        # -----------------------------------
        sec1 = sections.Item(1)

        try:
            sec1.Headers.Item(wdHeaderFooterPrimary).Range.Delete()
            sec1.Footers.Item(wdHeaderFooterPrimary).Range.Delete()
            sec1.PageSetup.DifferentFirstPageHeaderFooter = True
            print("   ✅ Section 1: no numbering")
        except Exception as e:
            print(f"   ⚠️ Section 1 cleanup error: {e}")

        # -----------------------------------
        # SECTION 2 → ROMAN NUMBERS (FIXED)
        # -----------------------------------
        sec2 = sections.Item(2)

        try:
            # Unlink from previous section
            sec2.Headers.Item(wdHeaderFooterPrimary).LinkToPrevious = False
            sec2.Footers.Item(wdHeaderFooterPrimary).LinkToPrevious = False
            print("   ✅ Section 2 unlinked")
        except Exception as e:
            print(f"   ⚠️ Unlink failed: {e}")

        try:
            # -----------------------------------
            # 🔧 FIXED: Set restart properties FIRST
            # -----------------------------------
            print("   🔧 Setting page restart properties...")
            sec2.PageSetup.RestartPageNumbering = True
            sec2.PageSetup.PageNumberStart = 1
            print("   ✅ Page restart: True, Start: 1")

            # -----------------------------------
            # 🔧 FIXED: Get footer and clear existing page numbers
            # -----------------------------------
            footer = sec2.Footers.Item(wdHeaderFooterPrimary)
            page_nums = footer.PageNumbers

            # Remove existing page numbers (important)
            print(f"   🧹 Removing {page_nums.Count} existing page numbers...")
            while page_nums.Count > 0:
                page_nums(1).Delete()

            # -----------------------------------
            # 🔧 FIXED: Add page number with Roman format
            # -----------------------------------
            print("   🔧 Adding Roman page numbers...")
            page_nums.Add(PageNumberAlignment=wdAlignParagraphCenter)

            # 🚨 CRITICAL FIX: Set NumberStyle on PageNumbers collection, NOT PageSetup
            page_nums.NumberStyle = wdPageNumberStyleLowercaseRoman
            page_nums.StartingNumber = 1  # Ensure starts at 1 (becomes "i")

            print(f"   ✅ Roman format applied: NumberStyle = {wdPageNumberStyleLowercaseRoman}")
            print("   ✅ Roman numbering configured (i, ii, iii...)")

        except Exception as e:
            print(f"   ❌ Page number setup failed: {e}")

            # Fallback method using field insertion
            try:
                print("   🔧 Trying fallback field method...")
                footer = sec2.Footers.Item(wdHeaderFooterPrimary)
                footer.Range.Delete()

                # Insert Roman numeral field directly
                footer_range = footer.Range
                field = footer_range.Fields.Add(
                    Range=footer_range,
                    Type=33,  # wdFieldPage
                    PreserveFormatting=False
                )
                field.Code.Text = "PAGE \\* ROMAN \\* LOWER"
                field.Update()

                # Center the page number
                footer.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

                print("   ✅ Fallback Roman field inserted")

            except Exception as e2:
                print(f"   ❌ Fallback method also failed: {e2}")

        # -----------------------------------
        # FORCE WORD TO APPLY CHANGES
        # -----------------------------------
        try:
            print("   🔄 Forcing document updates...")
            doc.ActiveWindow.View.ShowFieldCodes = False
            doc.Repaginate()
            doc.Fields.Update()
            doc.Repaginate()
            print("   ✅ Document refreshed")
        except Exception as e:
            print(f"   ⚠️ Refresh error: {e}")

        # -----------------------------------
        # VERIFICATION
        # -----------------------------------
        try:
            print("   🔍 Verifying Roman format...")
            sec2_footer = sec2.Footers.Item(wdHeaderFooterPrimary)
            sec2_page_nums = sec2_footer.PageNumbers
            if sec2_page_nums.Count > 0:
                style = sec2_page_nums(1).NumberStyle
                print(f"   📊 Current NumberStyle: {style} (should be {wdPageNumberStyleLowercaseRoman})")
                if style == wdPageNumberStyleLowercaseRoman:
                    print("   ✅ Roman format verified!")
                else:
                    print("   ⚠️ Roman format not applied correctly")
            else:
                print("   ⚠️ No page numbers found for verification")
        except Exception as e:
            print(f"   ⚠️ Verification failed: {e}")

    except Exception as e:
        print(f"   ❌ Pagination error: {e}")


# -----------------------------------
# ASSEMBLE FINAL DOC (IMPROVED)
# -----------------------------------
def assemble_final(title_doc, copyright_doc, toc_doc, output):
    print("\n📚 Assembling final document with Roman pagination fix")

    word = get_word()
    doc = word.Documents.Add()

    try:
        # Title page (Section 1) - NO page numbering
        if os.path.exists(title_doc):
            print("   📄 Inserting title page (Section 1)...")
            r = doc.Content
            r.Collapse(0)
            r.InsertFile(os.path.abspath(title_doc))

            # Insert section break (creates Section 2 for copyright + TOC)
            r = doc.Content
            r.Collapse(0)
            r.InsertBreak(wdSectionBreakNextPage)
            print("   ✅ Title page + section break inserted")

        # Copyright page (Section 2, page "i")
        if os.path.exists(copyright_doc):
            print("   📄 Inserting copyright page (will be Roman 'i')...")
            r = doc.Content
            r.Collapse(0)
            r.InsertFile(os.path.abspath(copyright_doc))

            # Page break before TOC (stays in same section)
            r = doc.Content
            r.Collapse(0)
            r.InsertBreak(wdPageBreak)
            print("   ✅ Copyright page inserted")

        # TOC (Section 2, pages "ii", "iii", etc.)
        print("   📄 Inserting TOC (will be Roman 'ii', 'iii'...)...")
        r = doc.Content
        r.Collapse(0)
        r.InsertFile(os.path.abspath(toc_doc))
        print("   ✅ TOC inserted")

        # Apply Roman pagination starting from copyright page
        apply_roman_pagination(doc)

        # Final processing
        print("   🔄 Final document processing...")
        try:
            doc.ActiveWindow.View.ShowFieldCodes = False
            doc.Repaginate()
            doc.Fields.Update()
            doc.Repaginate()
        except Exception as e:
            print(f"   ⚠️ Final processing warning: {e}")

        # Save
        print("   💾 Saving document...")
        doc.SaveAs2(os.path.abspath(output))
        doc.Close(False)
        word.Quit()

        print("\n🎉 DOCUMENT COMPLETED!")
        print(f"📁 File: {output}")
        print("\n📋 Expected result:")
        print("   • Title page: No number")
        print("   • Copyright page: Roman 'i'")
        print("   • TOC pages: Roman 'ii', 'iii', 'iv'...")
        print("   • ALL chapter titles on single lines")
        print("\n🔧 Roman format fix applied:")
        print("   • NumberStyle set on PageNumbers collection (not PageSetup)")
        print("   • Fallback field method if PageNumbers fails")

    except Exception as e:
        try:
            word.Quit()
        except:
            pass
        print(f"❌ Assembly error: {e}")


# -----------------------------------
# MAIN
# -----------------------------------
if __name__ == "__main__":
    project_root = find_project_root()

    cfg = load_builder_config(project_root)

    title_file = cfg["title_file"]
    copyright_file = cfg["copyright_file"]
    book_body_file = cfg["book_body_file"]
    
    # Setup temp directory for intermediate files
    temp_dir = cfg["temp_dir"]
    temp_dir.mkdir(parents=True, exist_ok=True)
    toc_temp = temp_dir / "toc.docx"

    # Setup output directory and file
    output_dir = cfg["output_dir"]
    output_dir.mkdir(parents=True, exist_ok=True)
    output_file = output_dir / cfg["output_filename"]

    print(f"\n📂 Project root: {project_root}")
    print(f"🛠️ Config: {pretty_path(project_root, cfg['config_path'])}")
    print(f"📘 Book body (for layout): {pretty_path(project_root, book_body_file)}")
    print(f"📄 Title: {pretty_path(project_root, title_file)}")
    print(f"📄 Copyright: {pretty_path(project_root, copyright_file)}")
    print(f"🧪 Temp TOC: {pretty_path(project_root, toc_temp)}")
    print(f"💾 Output: {pretty_path(project_root, output_file)}")

    headings = extract_headings(book_body_file)
    build_succeeded = False

    if headings:
        build_toc_doc(headings, toc_temp)
        assemble_final(title_file, copyright_file, toc_temp, output_file)
        build_succeeded = True

        if cfg["delete_temp_toc"] and os.path.exists(toc_temp):
            os.remove(toc_temp)
    else:
        print("\n⚠️ No headings were extracted. Front matter output was not generated.")

    output_exists = output_file.exists()
    print("\n================ FINAL SUMMARY ================")
    print(f"Status: {'SUCCESS' if (build_succeeded and output_exists) else 'FAILED'}")
    print(f"Headings extracted: {len(headings)}")
    print(f"Output file: {output_file}")
    print(f"Output exists: {'YES' if output_exists else 'NO'}")
    print("===============================================")

    # Save metadata for downstream tasks
    if build_succeeded and output_exists:
        metadata = {
            "task": "front_matter",
            "status": "success",
            "output_file": str(output_file.relative_to(project_root)),
            "page_count": "unknown",
            "last_page_numbering": "roman_lowercase",
            "next_arabic_page": 1,
            "headings_extracted": len(headings),
            "layout_applied": cfg.get("apply_book_layout", True),
            "timestamp": str(os.path.getmtime(output_file))
        }
        metadata_file = output_dir / cfg["metadata_filename"]
        with open(metadata_file, "w", encoding="utf-8") as f:
            json.dump(metadata, f, indent=2)
        print(f"✅ Metadata saved: {pretty_path(project_root, metadata_file)}")