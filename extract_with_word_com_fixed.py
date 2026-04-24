#!/usr/bin/env python3

import win32com.client
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
import os


# -----------------------------------
# CLEAN TEXT (fix XML error)
# -----------------------------------
def clean_text(text):
    if not text:
        return ""
    return "".join(
        c for c in text
        if c >= " " or c in ("\n", "\t")
    ).strip()


# -----------------------------------
# WORD APP
# -----------------------------------
def get_word():
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    return word


# -----------------------------------
# DETECT TITLE LINE (Georgia ~20 Bold)
# -----------------------------------
def is_title_line(para):
    try:
        f = para.Range.Font
        return (
            "georgia" in (f.Name or "").lower()
            and f.Size >= 18 and f.Size <= 22
            and f.Bold
        )
    except:
        return False


# -----------------------------------
# EXTRACT HEADINGS (FINAL LOGIC)
# -----------------------------------
def extract_headings(doc_path):
    print("\n📖 Extracting headings")

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

            if "heading 1" in style and text:

                # ---- Chapter with multi-line title ----
                if text.lower().startswith("chapter"):

                    chapter = text
                    title_parts = []
                    j = i + 1

                    while j <= total:
                        p = doc.Paragraphs(j)
                        t = clean_text(p.Range.Text)

                        if not t or not is_title_line(p):
                            break

                        title_parts.append(t)
                        j += 1

                    full = f"{chapter}: {' '.join(title_parts)}" if title_parts else chapter
                    page = para.Range.Information(3)

                    headings.append({"text": full, "page": page})
                    print(f"   📍 {full[:60]} → {page}")

                    i = j
                    continue

                # ---- Single-line heading (Discography, Books) ----
                else:
                    page = para.Range.Information(3)

                    headings.append({"text": text, "page": page})
                    print(f"   📍 {text} → {page}")

        except:
            pass

        i += 1

    doc.Close(False)
    word.Quit()

    print(f"   ✅ Total headings: {len(headings)}")
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
# APPLY ROMAN PAGINATION
# -----------------------------------
def apply_roman_pagination(doc):
    print("\n🔢 Applying Roman pagination")

    try:
        sections = doc.Sections

        if sections.Count < 2:
            print("   ⚠️ Not enough sections")
            return

        # Section 1 (Title) → no numbering
        sec1 = sections(1)
        try:
            sec1.Headers(1).Range.Delete()
            sec1.Footers(1).Range.Delete()
        except:
            pass

        # Section 2 → Roman numerals
        sec2 = sections(2)

        sec2.Headers(1).LinkToPrevious = False
        sec2.Footers(1).LinkToPrevious = False

        footer = sec2.Footers(1)
        footer.Range.ParagraphFormat.Alignment = 1

        footer.PageNumbers.Add()
        footer.PageNumbers.NumberStyle = 14  # roman lowercase

        sec2.PageSetup.RestartPageNumbering = True
        sec2.PageSetup.PageNumberStart = 1

        print("   ✅ Roman numbering applied")

    except Exception as e:
        print(f"   ⚠️ Pagination error: {e}")


# -----------------------------------
# ASSEMBLE FINAL DOC (COM)
# -----------------------------------
def assemble_final(title_doc, copyright_doc, toc_doc, output):
    print("\n📚 Assembling final document")

    word = get_word()
    doc = word.Documents.Add()

    try:
        # Title
        if os.path.exists(title_doc):
            r = doc.Content
            r.Collapse(0)
            r.InsertFile(os.path.abspath(title_doc))

            # Section break after title
            r = doc.Content
            r.Collapse(0)
            r.InsertBreak(2)

        # Copyright
        if os.path.exists(copyright_doc):
            r = doc.Content
            r.Collapse(0)
            r.InsertFile(os.path.abspath(copyright_doc))

        # TOC
        r = doc.Content
        r.Collapse(0)
        r.InsertBreak(7)

        r = doc.Content
        r.Collapse(0)
        r.InsertFile(os.path.abspath(toc_doc))

        # Apply Roman pagination
        apply_roman_pagination(doc)

        # Save
        doc.SaveAs2(os.path.abspath(output))
        doc.Close(False)
        word.Quit()

        print("\n🎉 DONE")
        print(f"📁 {output}")

    except Exception as e:
        word.Quit()
        print(f"❌ Error: {e}")


# -----------------------------------
# MAIN
# -----------------------------------
if __name__ == "__main__":
    book_file = "Hits And Happiness Final 2 Discog.docx"
    title_file = "HH Title.docx"
    copyright_file = "HH Copyright page.docx"
    toc_temp = "temp_toc.docx"
    output_file = "Hits And Happiness Final 2 TOC.docx"

    headings = extract_headings(book_file)

    if headings:
        build_toc_doc(headings, toc_temp)
        assemble_final(title_file, copyright_file, toc_temp, output_file)

        if os.path.exists(toc_temp):
            os.remove(toc_temp)