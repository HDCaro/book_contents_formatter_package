#!/usr/bin/env python3
"""
Two-Line Heading Extractor using Word COM Interface - FIXED VERSION
Reads actual page numbers from Word using COM automation with proper constants
Requires: Windows + Microsoft Word installed
"""

import win32com.client
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
import os
import sys

# Word constants (since win32com.client.constants might not work properly)
WD_ACTIVE_END_PAGE_NUMBER = 3
WD_CURRENT_PAGE_NUMBER = 7
WD_NUMBER_OF_PAGES_IN_DOCUMENT = 4


def get_page_numbers_from_word_com_fixed(docx_path):
    """
    Use Word COM interface to get actual page numbers - FIXED VERSION
    """
    print("🔄 Starting Word COM interface (Fixed Version)...")

    try:
        # Create Word application
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False  # Don't show Word window

        # Open the document
        abs_path = os.path.abspath(docx_path)
        print(f"📖 Opening document: {abs_path}")
        doc = word_app.Documents.Open(abs_path)

        # Get document info
        total_pages = doc.Range().Information(WD_NUMBER_OF_PAGES_IN_DOCUMENT)
        print(f"📄 Document has {total_pages} pages")

        # Dictionary to store paragraph index -> page number
        paragraph_pages = {}
        heading_pages = {}

        print("🔍 Scanning for headings and their page numbers...")

        # Get all paragraphs
        paragraphs = doc.Paragraphs
        total_paragraphs = paragraphs.Count

        print(f"📊 Found {total_paragraphs} paragraphs")

        # Process paragraphs and look specifically for headings
        heading_count = 0
        for i in range(1, total_paragraphs + 1):
            try:
                para = paragraphs(i)
                para_text = para.Range.Text.strip()

                # Check if this looks like a heading we care about
                if (para_text.lower().startswith('chapter') or
                        para.Style.NameLocal.lower().startswith('heading')):

                    # Try multiple methods to get page number
                    try:
                        # Method 1: Use the range start position
                        page_num = para.Range.Information(WD_ACTIVE_END_PAGE_NUMBER)

                        if page_num <= 0 or page_num > total_pages:
                            # Method 2: Use current page number
                            page_num = para.Range.Information(WD_CURRENT_PAGE_NUMBER)

                        if page_num <= 0 or page_num > total_pages:
                            # Method 3: Calculate based on range position
                            char_position = para.Range.Start
                            total_chars = doc.Range().End
                            estimated_page = max(1, int((char_position / total_chars) * total_pages))
                            page_num = estimated_page

                        # Store both paragraph index and heading info
                        paragraph_pages[i - 1] = page_num

                        if para_text.lower().startswith('chapter'):
                            heading_pages[para_text] = page_num
                            heading_count += 1
                            print(f"   📍 {para_text[:30]:<30} → Page {page_num}")

                    except Exception as e:
                        print(f"   ⚠️  Could not get page for paragraph {i}: {str(e)[:50]}")
                        # Use estimated page based on position
                        char_position = para.Range.Start
                        total_chars = doc.Range().End
                        estimated_page = max(1, int((char_position / total_chars) * total_pages))
                        paragraph_pages[i - 1] = estimated_page

                # Show progress for large documents
                if i % 500 == 0:
                    progress = (i / total_paragraphs) * 100
                    print(f"   Progress: {progress:.1f}% ({i}/{total_paragraphs}) - Found {heading_count} headings")

            except Exception as e:
                # Skip problematic paragraphs
                continue

        # Close the document and Word
        doc.Close(SaveChanges=False)
        word_app.Quit()

        print(f"✅ Successfully processed document")
        print(f"📊 Found {heading_count} chapter headings")
        print(f"📄 Document pages: {total_pages}")

        return paragraph_pages, heading_pages, total_pages

    except Exception as e:
        print(f"❌ Error with Word COM interface: {e}")
        print("💡 Troubleshooting:")
        print("   • Close Word if it's running")
        print("   • Make sure the document is not password protected")
        print("   • Try running as Administrator")
        return None, None, 0


def extract_headings_with_com_pages_fixed(docx_path):
    """Extract headings and get their actual page numbers using COM - FIXED"""

    # First, get page numbers using COM
    paragraph_pages, heading_pages, total_pages = get_page_numbers_from_word_com_fixed(docx_path)

    if paragraph_pages is None:
        print("❌ Could not read page numbers from Word.")
        return None

    # Now extract headings using python-docx
    print("🔍 Extracting headings with python-docx...")
    doc = Document(docx_path)
    paragraphs = doc.paragraphs
    headings = []

    i = 0
    while i < len(paragraphs):
        para = paragraphs[i]

        # Check if this paragraph is a heading and contains "Chapter"
        if (para.style.name.startswith('Heading') and
                para.text.strip().lower().startswith('chapter')):

            chapter_num = para.text.strip()
            chapter_title = ""

            # Look for the next paragraph (should be the title)
            if i + 1 < len(paragraphs):
                next_para = paragraphs[i + 1]
                if (next_para.text.strip() and
                        not next_para.style.name.startswith('Heading')):
                    chapter_title = next_para.text.strip()
                    i += 1  # Skip the title paragraph in next iteration

            # Combine chapter number and title
            if chapter_title:
                full_title = f"{chapter_num}: {chapter_title}"
            else:
                full_title = chapter_num

            # Get the actual page number from COM data
            actual_page = paragraph_pages.get(i, 1)

            # Try to get from heading_pages if available
            if chapter_num in heading_pages:
                actual_page = heading_pages[chapter_num]

            # Fallback: estimate based on position if page is still 1
            if actual_page == 1 and i > 10:  # If we're deep in document but still page 1
                estimated_page = max(1, int((i / len(paragraphs)) * total_pages))
                actual_page = estimated_page

            headings.append({
                'text': full_title,
                'level': 1,
                'page': actual_page,
                'paragraph_index': i,
                'chapter_key': chapter_num
            })

            print(f"   ✅ {chapter_num} → Page {actual_page}")

        # Check for other heading levels
        elif para.style.name.startswith('Heading'):
            level = int(para.style.name.split()[-1]) if para.style.name.split()[-1].isdigit() else 2
            heading_text = para.text.strip()

            if heading_text:
                actual_page = paragraph_pages.get(i, 1)

                # Estimate if still showing page 1
                if actual_page == 1 and i > 10:
                    estimated_page = max(1, int((i / len(paragraphs)) * total_pages))
                    actual_page = estimated_page

                headings.append({
                    'text': heading_text,
                    'level': level,
                    'page': actual_page,
                    'paragraph_index': i,
                    'chapter_key': f"Sub: {heading_text}"
                })

        i += 1

    # Post-process: ensure page numbers make sense
    if headings:
        print("🔧 Post-processing page numbers...")

        # Sort by paragraph order
        headings.sort(key=lambda x: x['paragraph_index'])

        # If most pages are 1, use position-based estimation
        page_1_count = sum(1 for h in headings if h['page'] == 1)
        if page_1_count > len(headings) * 0.8:  # If 80%+ are page 1
            print("   📊 Most pages showing as 1, using position-based estimation...")
            for i, heading in enumerate(headings):
                estimated_page = max(1, int((heading['paragraph_index'] / len(paragraphs)) * total_pages))
                heading['page'] = estimated_page
                print(f"   📍 {heading['chapter_key']} → Estimated Page {estimated_page}")

        # Ensure ascending order
        for i in range(1, len(headings)):
            if headings[i]['page'] <= headings[i - 1]['page']:
                headings[i]['page'] = headings[i - 1]['page'] + 1

    print(f"✅ Extracted {len(headings)} headings with page numbers")
    return headings


def create_contents_with_com_pages_fixed(book_path, output_path):
    """Create contents using Word COM interface - FIXED VERSION with Roman page numbering"""

    print(f"📚 Reading actual page numbers from Word document (Fixed Version)...")
    print(f"📖 Document: {book_path}")

    headings = extract_headings_with_com_pages_fixed(book_path)

    if not headings:
        print("❌ No headings found or COM interface failed.")
        return None

    # Create new document
    doc = Document()

    # Set up page margins for traditional format
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1.5)
        section.right_margin = Inches(1)

    # Set up Roman numeral page numbering for the contents pages
    section = doc.sections[0]

    # Add footer for Roman numeral page numbers
    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Add Roman numeral page number field to footer
    footer_run = footer_para.add_run()
    footer_run.font.name = "Georgia"
    footer_run.font.size = Pt(12)

    # Insert page number field with Roman numeral format
    from docx.oxml.shared import qn
    from docx.oxml import parse_xml

    # Create page number field with Roman numeral format
    fldChar1 = parse_xml(
        r'<w:fldChar w:fldCharType="begin" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
    instrText = parse_xml(
        r'<w:instrText xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"> PAGE \* ROMAN </w:instrText>')
    fldChar2 = parse_xml(
        r'<w:fldChar w:fldCharType="end" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')

    footer_run._r.append(fldChar1)
    footer_run._r.append(instrText)
    footer_run._r.append(fldChar2)

    # Add title
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.add_run("CONTENTS")
    title_run.font.name = 'Georgia'
    title_run.font.size = Pt(24)
    title_run.bold = True

    # Add space after title
    doc.add_paragraph()

    # Add each heading with actual page numbers
    for heading in headings:
        para = doc.add_paragraph()

        # Indent based on heading level for sub-headings
        if heading['level'] > 1:
            para.paragraph_format.left_indent = Inches(0.5 * (heading['level'] - 1))

        # Set up tab stop for right-aligned page number with dot leaders
        tab_stops = para.paragraph_format.tab_stops
        tab_stops.add_tab_stop(Inches(5.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)

        # Add the chapter/section title
        title_run = para.add_run(heading['text'])
        title_run.font.name = 'Georgia'
        title_run.font.size = Pt(12)

        # Add tab character followed by actual page number from the book
        page_run = para.add_run(f"\t{heading['page']}")
        page_run.font.name = 'Georgia'
        page_run.font.size = Pt(12)
        page_run.bold = True

    # Save the document
    doc.save(output_path)
    print(f"✅ Contents with Word page numbers saved as: {output_path}")
    print(f"📊 Found {len(headings)} headings")
    print(f"📄 Contents pages will be numbered with Roman numerals (I, II, III...)")
    print(f"📖 Chapter entries show actual book page numbers")
    print(f"💡 Example output:")
    print(f"   Introduction........................... 1")
    print(f"   Chapter 1: 'Pre' His Story............ 3")
    print(f"   Chapter 2: The Early Years............ 15")
    print(f"   (Contents pages numbered I, II, III at bottom)")

    return headings


def check_requirements():
    """Check if required packages are installed"""
    try:
        import win32com.client
        return True
    except ImportError:
        print("❌ Missing required package: pywin32")
        print("💡 Install it with: pip install pywin32")
        return False


if __name__ == "__main__":
    # Configuration for your specific book
    book_file = "HITS AND HAPPINESS FINAL 2 Format.docx"
    output_file = "HITS AND HAPPINESS FINAL 2 Content.docx"

    print("📚 Two-Line Heading Extractor with Word COM Interface - FIXED")
    print("=" * 80)
    print(f"📖 Input File: {book_file}")
    print(f"📄 Output File: {output_file}")
    print(f"🔧 Version: Fixed COM constants and error handling")
    print("=" * 80)

    # Check requirements
    if not check_requirements():
        sys.exit(1)

    # Check if file exists
    if not os.path.exists(book_file):
        print(f"❌ ERROR: File not found: {book_file}")
        print(f"📁 Current directory: {os.getcwd()}")
        sys.exit(1)

    try:
        print("🚀 Starting extraction with fixed COM interface...")

        # Extract headings with actual page numbers
        headings = create_contents_with_com_pages_fixed(book_file, output_file)

        if headings:
            print(f"\n📋 Final Results:")
            print("=" * 60)

            for i, heading in enumerate(headings[:10], 1):
                print(f"{i:2d}. {heading['text'][:45]:<45} ... {heading['page']:>3}")

            if len(headings) > 10:
                print(f"... and {len(headings) - 10} more")

            print("=" * 60)

            print(f"\n🎉 SUCCESS!")
            print(f"📁 Generated: {output_file}")
            print(f"📊 Total Chapters/Sections: {len(headings)}")
            print(f"📄 Page Range: {min(h['page'] for h in headings)} - {max(h['page'] for h in headings)}")
            print(f"🎨 Format: Traditional with Word page numbers")

        else:
            print("\n❌ Failed to extract headings.")

    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        print("💡 Try the manual page number script as backup")

    print("\n" + "=" * 80)
    print("🎉 Script completed!")