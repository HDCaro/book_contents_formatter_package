#!/usr/bin/env python3

import win32com.client
import os
import time

WD_PAGE_BREAK = 7
WD_SECTION_BREAK_NEXT_PAGE = 2
WD_TAB_ALIGNMENT_RIGHT = 2
WD_TAB_LEADER_DOTS = 1
WD_ACTIVE_END_PAGE_NUMBER = 3


# -----------------------------------
# WORD APP
# -----------------------------------
def get_word():
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    word.ScreenUpdating = False
    return word


# -----------------------------------
# FONT CHECK (CRITICAL FIX)
# -----------------------------------
def is_title_line(para):
    try:
        font = para.Range.Font

        name = (font.Name or "").lower()
        size = font.Size
        bold = font.Bold

        return (
            "georgia" in name and
            size >= 18 and size <= 22 and
            bold
        )
    except:
        return False


# -----------------------------------
# EXTRACT HEADINGS (FONT-AWARE)
# -----------------------------------
def extract_headings(doc_path):
    print("\n📖 Extracting headings (FONT-AWARE VERSION)")

    word = get_word()

    try:
        doc = word.Documents.Open(os.path.abspath(doc_path))
        doc.Repaginate()

        total = doc.Paragraphs.Count
        print(f"   📊 Paragraphs: {total}")

        headings = []

        i = 1
        while i <= total:

            try:
                para = doc.Paragraphs(i)
                text = para.Range.Text.strip()
                style = para.Style.NameLocal.lower()

                # -------------------------
                # CHAPTER DETECTION
                # -------------------------
                if "heading" in style and text.lower().startswith("chapter"):

                    chapter = text
                    title_parts = []

                    j = i + 1

                    print(f"\n   📘 Found {chapter}")

                    # ONLY accept Georgia 20 Bold lines
                    while j <= total:
                        next_para = doc.Paragraphs(j)
                        next_text = next_para.Range.Text.strip()

                        if not next_text:
                            break

                        if not is_title_line(next_para):
                            break

                        title_parts.append(next_text)
                        print(f"      ➕ Title line: {next_text}")

                        j += 1

                    # Combine safely
                    if title_parts:
                        full = f"{chapter}: {' '.join(title_parts)}"
                    else:
                        full = chapter

                    page = para.Range.Information(WD_ACTIVE_END_PAGE_NUMBER)

                    headings.append({
                        "text": full,
                        "page": page
                    })

                    print(f"      ✅ Final: {full} → {page}")

                    i = j
                    continue

                # -------------------------
                # OTHER HEADINGS
                # -------------------------
                elif "heading" in style and text:
                    page = para.Range.Information(WD_ACTIVE_END_PAGE_NUMBER)

                    headings.append({
                        "text": text,
                        "page": page
                    })

            except:
                pass

            i += 1

        doc.Close(False)
        word.Quit()

        print(f"\n   ✅ Total headings: {len(headings)}")
        return headings

    except Exception as e:
        word.Quit()
        print(f"❌ Error: {e}")
        return None


# -----------------------------------
# BUILD DOCUMENT
# -----------------------------------
def build_document(book, title_doc, copyright_doc, output):
    print("\n📚 GENERATING FRONT MATTER (FINAL FIXED VERSION)")

    headings = extract_headings(book)

    if not headings:
        print("❌ No headings")
        return

    word = get_word()

    try:
        doc = word.Documents.Add()

        # TITLE
        if os.path.exists(title_doc):
            print("📄 Inserting title...")
            rng = doc.Content
            rng.Collapse(0)
            rng.InsertFile(os.path.abspath(title_doc))

        # COPYRIGHT
        if os.path.exists(copyright_doc):
            print("📄 Inserting copyright...")

            rng = doc.Content
            rng.Collapse(0)
            rng.InsertBreak(WD_SECTION_BREAK_NEXT_PAGE)

            rng = doc.Content
            rng.Collapse(0)
            rng.InsertFile(os.path.abspath(copyright_doc))

        # TOC
        print("📑 Creating TOC...")

        rng = doc.Content
        rng.Collapse(0)
        rng.InsertBreak(WD_PAGE_BREAK)

        rng = doc.Content
        rng.Collapse(0)
        rng.InsertAfter("CONTENTS\r\n\r\n")

        start = rng.End

        toc_text = ""
        for h in headings:
            toc_text += f"{h['text']}\t{h['page']}\r\n"

        rng.InsertAfter(toc_text)

        toc_range = doc.Range(start, rng.End)

        # FORMAT
        for p in toc_range.Paragraphs:
            try:
                p.Range.Font.Name = "Georgia"
                p.Range.Font.Size = 12

                p.TabStops.ClearAll()
                p.TabStops.Add(
                    Position=450,
                    Alignment=WD_TAB_ALIGNMENT_RIGHT,
                    Leader=WD_TAB_LEADER_DOTS
                )
            except:
                pass

        # SAVE
        doc.SaveAs2(os.path.abspath(output))
        doc.Close(False)
        word.Quit()

        print("\n🎉 DONE")
        print(f"📁 {output}")
        print(f"📊 Headings: {len(headings)}")

    except Exception as e:
        word.Quit()
        print(f"❌ Error: {e}")


# -----------------------------------
# RUN
# -----------------------------------
if __name__ == "__main__":
    build_document(
        "Hits And Happiness Final 2 Discog.docx",
        "HH Title.docx",
        "HH Copyright page.docx",
        "Hits And Happiness Final 2 TOC.docx"
    )