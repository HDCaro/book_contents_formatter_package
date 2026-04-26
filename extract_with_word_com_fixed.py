#!/usr/bin/env python3

import win32com.client
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
import os

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
# APPLY ROMAN PAGINATION (FIXED FOR COPYRIGHT PAGE START)
# -----------------------------------
def apply_roman_pagination(doc):
    print("\n🔢 Applying Roman pagination (Word-equivalent method)")

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
        # SECTION 2 → ROMAN NUMBERS
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
            # ✅ THIS IS THE UI EQUIVALENT:
            # Page Number → Format → Roman (i, ii, iii)
            # -----------------------------------
            sec2.PageSetup.RestartPageNumbering = True
            sec2.PageSetup.PageNumberStart = 1
            sec2.PageSetup.PageNumberStyle = wdPageNumberStyleLowercaseRoman

            footer = sec2.Footers.Item(wdHeaderFooterPrimary)
            page_nums = footer.PageNumbers

            # Remove existing page numbers (important)
            while page_nums.Count > 0:
                page_nums(1).Delete()

            # Add page number field (centered)
            page_nums.Add(PageNumberAlignment=wdAlignParagraphCenter)

            print("   ✅ Roman numbering applied (i, ii, iii...)")

        except Exception as e:
            print(f"   ⚠️ Page number setup failed: {e}")

        # -----------------------------------
        # FORCE WORD TO APPLY CHANGES
        # -----------------------------------
        try:
            doc.Repaginate()
            doc.Fields.Update()
            doc.Repaginate()
            print("   🔄 Document refreshed")
        except Exception as e:
            print(f"   ⚠️ Refresh error: {e}")

    except Exception as e:
        print(f"   ❌ Pagination error: {e}")
# -----------------------------------
# ASSEMBLE FINAL DOC (IMPROVED)
# -----------------------------------
def assemble_final(title_doc, copyright_doc, toc_doc, output):
    print("\n📚 Assembling final document with single-line titles")

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
    book_file = "HITS AND HAPPINESS FINAL 2 Format MOM Discog.docx"
    title_file = "HH Title.docx"
    copyright_file = "HH Copyright page.docx"
    toc_temp = "temp_toc.docx"
    output_file = "HITS AND HAPPINESS FINAL 2 Format MOM Discog TOC.docx"

    headings = extract_headings(book_file)

    if headings:
        build_toc_doc(headings, toc_temp)
        assemble_final(title_file, copyright_file, toc_temp, output_file)

        if os.path.exists(toc_temp):
            os.remove(toc_temp)
