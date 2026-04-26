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
# CLEAN TEXT (fix XML error AND trailing slashes)
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


# -----------------------------------
# WORD APP
# -----------------------------------
def get_word():
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    return word


# -----------------------------------
# EXTRACT HEADINGS (HEADING 2 APPROACH)
# -----------------------------------
def extract_headings(doc_path):
    print("\n📖 Extracting headings using Heading 1 + Heading 2 approach")

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

                # ---- Chapter with Heading 2 title ----
                if text.lower().startswith("chapter"):
                    print(f"   🔍 Processing chapter: '{text}'")

                    chapter = text
                    chapter_title = None
                    j = i + 1

                    # Look for the next Heading 2 (chapter title)
                    print(f"   🔍 Looking for Heading 2 title starting from paragraph {j}...")

                    while j <= total:
                        try:
                            p = doc.Paragraphs(j)
                            t = clean_text(p.Range.Text)
                            p_style = p.Style.NameLocal.lower()

                            print(f"   📝 Checking paragraph {j}: '{t[:50]}...' (style: {p_style})")

                            if not t:
                                print(f"   ⏭️ Empty paragraph {j}, skipping")
                                j += 1
                                continue

                            # Found Heading 2 - this is our chapter title
                            if "heading 2" in p_style:
                                chapter_title = t
                                print(f"   ✅ Found Heading 2 title: '{chapter_title}'")
                                j += 1  # Move past this heading
                                break

                            # If we hit another Heading 1, stop looking
                            elif "heading 1" in p_style:
                                print(f"   🛑 Hit another Heading 1, stopping search")
                                break

                            # Skip other content
                            else:
                                j += 1
                                continue

                        except Exception as e:
                            print(f"   ⚠️ Error processing paragraph {j}: {e}")
                            break

                    # Build the complete title
                    if chapter_title:
                        full = f"{chapter}: {chapter_title}"
                        print(f"   🎉 Complete chapter: '{full}'")
                    else:
                        full = chapter
                        print(f"   ⚠️ No Heading 2 found, using chapter only: '{full}'")

                    full = clean_text(full)  # Final cleanup
                    page = para.Range.Information(3)

                    headings.append({"text": full, "page": page})
                    print(f"   📌 Added to TOC: '{full}' → page {page}")

                    i = j
                    continue

                # ---- Single-line heading (Discography, Books, etc.) ----
                else:
                    page = para.Range.Information(3)
                    headings.append({"text": text, "page": page})
                    print(f"   📌 Single-line heading: '{text}' → page {page}")

            # Also capture standalone Heading 2 (if not part of a chapter)
            elif "heading 2" in style and text:
                # Check if this Heading 2 follows a Chapter (already processed above)
                # If not, treat it as a standalone heading
                page = para.Range.Information(3)
                headings.append({"text": text, "page": page})
                print(f"   📌 Standalone Heading 2: '{text}' → page {page}")

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
# BUILD TOC (python-docx)
# -----------------------------------
def build_toc_doc(headings, toc_path):
    print("\n📝 Building TOC")

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

        # Hanging indent (restored working behavior)
        para.paragraph_format.left_indent = Inches(0.5)
        para.paragraph_format.first_line_indent = Inches(-0.5)

        # Tab stop with dots
        para.paragraph_format.tab_stops.add_tab_stop(
            Inches(6),
            WD_TAB_ALIGNMENT.RIGHT,
            WD_TAB_LEADER.DOTS
        )

        safe_text = clean_text(h["text"])

        run = para.add_run(safe_text)
        run.font.name = "Georgia"
        run.font.size = Pt(12)

        page_run = para.add_run(f"\t{h['page']}")
        page_run.bold = True

    doc.save(toc_path)
    print("   ✅ TOC created")

    return toc_path


# -----------------------------------
# APPLY ROMAN PAGINATION (FIXED WITH PROPER RESTART)
# -----------------------------------
def apply_roman_pagination(doc):
    print("\n🔢 Applying Roman pagination")

    try:
        sections = doc.Sections
        print(f"   📊 Total sections: {sections.Count}")

        if sections.Count < 2:
            print("   ⚠️ Not enough sections")
            return

        # Section 1 (Title) → no numbering
        print("   🔧 Configuring Section 1 (Title) - NO numbering")
        sec1 = sections.Item(1)
        try:
            # Clear headers and footers for title page
            sec1.Headers.Item(wdHeaderFooterPrimary).Range.Delete()
            sec1.Footers.Item(wdHeaderFooterPrimary).Range.Delete()

            # Ensure no page numbering on title section
            sec1.PageSetup.DifferentFirstPageHeaderFooter = True

            print("   ✅ Section 1 configured with no numbering")

        except Exception as e:
            print(f"   ⚠️ Section 1 cleanup: {e}")

        # Section 2 → Roman numerals starting from i (copyright = i, TOC = ii, iii, etc.)
        print("   🔧 Configuring Section 2 (Copyright + TOC) - Roman numerals starting from i")
        sec2 = sections.Item(2)

        # CRITICAL: Unlink from previous section
        try:
            sec2.Headers.Item(wdHeaderFooterPrimary).LinkToPrevious = False
            sec2.Footers.Item(wdHeaderFooterPrimary).LinkToPrevious = False
            print("   ✅ Section 2 unlinked from previous section")
        except Exception as e:
            print(f"   ⚠️ Unlinking failed: {e}")

        # Get the footer and clear it completely
        footer = sec2.Footers.Item(wdHeaderFooterPrimary)
        footer.Range.Delete()

        # Method 1: Using PageNumbers collection with explicit restart
        try:
            print("   🔧 Method 1: PageNumbers collection with restart...")

            # Configure page setup FIRST (before adding page numbers)
            page_setup = sec2.PageSetup

            # Set restart properties
            try:
                page_setup.RestartPageNumbering = True
                print("   ✅ RestartPageNumbering = True")
            except:
                print("   ⚠️ RestartPageNumbering property not available")

            try:
                page_setup.PageNumberStart = 1
                print("   ✅ PageNumberStart = 1")
            except:
                print("   ⚠️ PageNumberStart property not available")

            # Now add the page numbers
            page_nums = footer.PageNumbers
            page_nums.Add(PageNumberAlignment=wdAlignParagraphCenter)

            # Set the number style to lowercase Roman
            page_nums.NumberStyle = wdPageNumberStyleLowercaseRoman
            page_nums.StartingNumber = 1  # This should make copyright page = i

            print("   ✅ PageNumbers configured: Roman, starting at 1")

            # Force updates
            doc.Fields.Update()
            doc.Repaginate()

            print("   ✅ Method 1 successful - Roman numbering with restart")

        except Exception as e1:
            print(f"   ⚠️ Method 1 failed: {e1}")

            # Method 2: Manual field insertion with restart settings
            try:
                print("   🔧 Method 2: Manual field with restart settings...")
                footer.Range.Delete()

                # Set page setup properties first
                page_setup = sec2.PageSetup
                try:
                    page_setup.RestartPageNumbering = True
                    page_setup.PageNumberStart = 1
                    print("   ✅ Page restart settings applied")
                except Exception as setup_error:
                    print(f"   ⚠️ Page setup failed: {setup_error}")

                # Create the field
                footer_range = footer.Range
                footer_range.Collapse(1)  # Collapse to end

                # Insert field using proper Word field insertion
                field = footer_range.Fields.Add(
                    Range=footer_range,
                    Type=wdFieldPage,
                    PreserveFormatting=False
                )

                # Set the field code for Roman numerals
                field.Code.Text = "PAGE \\* ROMAN \\* LOWER"
                field.Update()

                # Center the footer
                footer.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

                # Update document
                doc.Fields.Update()
                doc.Repaginate()

                print("   ✅ Method 2 successful - Manual field with restart")

            except Exception as e2:
                print(f"   ⚠️ Method 2 failed: {e2}")

                # Method 3: Selection-based with restart
                try:
                    print("   🔧 Method 3: Selection-based with restart...")
                    footer.Range.Delete()

                    # Try to set page setup first
                    try:
                        sec2.PageSetup.RestartPageNumbering = True
                        sec2.PageSetup.PageNumberStart = 1
                    except:
                        pass

                    # Select the footer range
                    footer.Range.Select()
                    selection = doc.Application.Selection

                    # Insert page field using Selection
                    selection.Fields.Add(
                        Range=selection.Range,
                        Type=wdFieldPage,
                        Text="PAGE \\* ROMAN \\* LOWER"
                    )

                    # Center alignment
                    selection.ParagraphFormat.Alignment = wdAlignParagraphCenter

                    # Update fields
                    doc.Fields.Update()
                    doc.ActiveWindow.View.ShowFieldCodes = False
                    doc.Repaginate()

                    print("   ✅ Method 3 successful - Selection-based with restart")

                except Exception as e3:
                    print(f"   ⚠️ Method 3 failed: {e3}")

                    # Method 4: Simple fallback
                    try:
                        print("   🔧 Method 4: Simple fallback...")
                        footer.Range.Delete()
                        footer.Range.Text = "i"
                        footer.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                        print("   ⚠️ Applied simple Roman numeral")

                    except Exception as e4:
                        print(f"   ❌ All methods failed: {e4}")

        # Final comprehensive update
        try:
            print("   🔄 Final updates...")
            doc.ActiveWindow.View.ShowFieldCodes = False
            doc.Repaginate()
            doc.Fields.Update()
            doc.Repaginate()
            print("   ✅ Final updates completed")

        except Exception as e:
            print(f"   ⚠️ Final update failed: {e}")

    except Exception as e:
        print(f"   ❌ Pagination error: {e}")


# -----------------------------------
# ASSEMBLE FINAL DOC (COM) - IMPROVED SECTION HANDLING
# -----------------------------------
def assemble_final(title_doc, copyright_doc, toc_doc, output):
    print("\n📚 Assembling final document")

    word = get_word()
    doc = word.Documents.Add()

    try:
        # Title page (Section 1) - NO page numbering
        if os.path.exists(title_doc):
            print("   📄 Inserting title page (Section 1)...")
            r = doc.Content
            r.Collapse(0)  # Collapse to start
            r.InsertFile(os.path.abspath(title_doc))

            # Insert section break after title (creates Section 2)
            r = doc.Content
            r.Collapse(0)  # Move to end
            r.InsertBreak(wdSectionBreakNextPage)
            print("   ✅ Title page inserted with section break")

        # Copyright page (goes into Section 2) - Should be Roman numeral i
        if os.path.exists(copyright_doc):
            print("   📄 Inserting copyright page (Section 2, page i)...")
            r = doc.Content
            r.Collapse(0)  # Move to end
            r.InsertFile(os.path.abspath(copyright_doc))

            # Add page break before TOC (stays in same section)
            r = doc.Content
            r.Collapse(0)
            r.InsertBreak(wdPageBreak)
            print("   ✅ Copyright page inserted")

        # TOC (also in Section 2) - Should be Roman numerals ii, iii, etc.
        print("   📄 Inserting TOC (Section 2, pages ii, iii, etc.)...")
        r = doc.Content
        r.Collapse(0)
        r.InsertFile(os.path.abspath(toc_doc))
        print("   ✅ TOC inserted")

        # Apply Roman pagination to Section 2 (copyright + TOC)
        apply_roman_pagination(doc)

        # Final comprehensive update
        print("   🔄 Final document processing...")
        try:
            # Make sure we're not showing field codes
            doc.ActiveWindow.View.ShowFieldCodes = False

            # Multiple updates to ensure everything processes correctly
            doc.Repaginate()
            doc.Fields.Update()
            doc.Repaginate()  # Second repagination to ensure restart takes effect

        except Exception as e:
            print(f"   ⚠️ Final processing warning: {e}")

        # Save
        print("   💾 Saving document...")
        doc.SaveAs2(os.path.abspath(output))
        doc.Close(False)
        word.Quit()

        print("\n🎉 DONE")
        print(f"📁 {output}")
        print("\n📋 Expected page numbering:")
        print("   • Title page: No page number")
        print("   • Copyright page: Roman numeral i")
        print("   • TOC pages: Roman numerals ii, iii, iv, etc.")

    except Exception as e:
        try:
            word.Quit()
        except:
            pass
        print(f"❌ Error: {e}")


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
