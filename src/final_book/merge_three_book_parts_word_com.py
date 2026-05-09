"""
merge_three_book_parts_word_com.py

This script combines three separate Word documents (front matter, book body, and index)
into a single professionally formatted book with proper page numbering:
- Front Matter: Roman numerals (i, ii, iii...)
- Book Body: Arabic numerals starting at 1 (1, 2, 3...)
- Index: Roman numerals continuing from front matter

The script uses Word COM automation with improved error handling and connection management.

Requirements: pywin32 (pip install pywin32), Word 365 installed
"""

import win32com.client
import os
import sys
import time
from pathlib import Path

# File paths configuration
from pathlib import Path

# Detect the root folder (book_contents_formatter_package)
ROOT = Path(__file__).resolve().parents[2]  # Go up 2 levels from src/final_book/

# Paths relative to root
FRONT_MATTER_PATH = ROOT / "release" / "v1" / "HITS_AND_HAPPINESS_FINAL_FRONT_MATTER.docx"
BOOK_BODY_PATH    = ROOT / "release" / "v1" / "HITS_AND_HAPPINESS_FINAL_BODY2.docx"
INDEX_PATH        = ROOT / "release" / "v1" / "HITS AND HAPPINESS _FINAL_INDEX.docx"
OUTPUT_PATH       = ROOT / "release" / "v1" / "HITS_AND_HAPPINESS_FULL_BOOK.docx"

# Example usage
print(FRONT_MATTER_PATH)

INTRO_WORD = "INTRODUCTION"


def check_file_exists(file_path, file_description):
    """Check if a file exists and provide helpful error message."""
    if not os.path.exists(file_path):
        print(f"ERROR: {file_description} not found at: {file_path}")
        return False
    print(f"✓ Found {file_description}: {os.path.basename(file_path)}")
    return True


def create_word_app():
    """Create Word application with proper error handling."""
    try:
        print("  Initializing Word COM object...")
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = True  # Make visible for debugging
        word_app.DisplayAlerts = False  # Suppress alerts
        time.sleep(1)  # Give Word time to initialize
        return word_app
    except Exception as e:
        print(f"  ERROR creating Word application: {str(e)}")
        return None


def insert_file_content(word_app, target_doc, source_path):
    """Insert content from source file using Word's built-in InsertFile method."""
    try:
        print(f"  Inserting content from {os.path.basename(source_path)}...")

        # Move to end of document
        selection = word_app.Selection
        selection.EndKey(Unit=6)  # wdStory = 6 (end of document)

        # Use InsertFile method which preserves formatting better
        selection.InsertFile(
            FileName=source_path,
            Range="",
            ConfirmConversions=False,
            Link=False,
            Attachment=False
        )

        print(f"  ✓ Successfully inserted {os.path.basename(source_path)}")
        return True

    except Exception as e:
        print(f"  ERROR inserting {os.path.basename(source_path)}: {str(e)}")
        return False


def insert_section_break(word_app):
    """Insert a section break (next page)."""
    try:
        selection = word_app.Selection
        selection.EndKey(Unit=6)  # Move to end of document
        selection.InsertBreak(Type=7)  # wdSectionBreakNextPage = 7
        print("  ✓ Section break inserted")
        return True
    except Exception as e:
        print(f"  ERROR inserting section break: {str(e)}")
        return False


def setup_page_numbering_simple(word_app, section_index, number_style, start_number=1):
    """Set up page numbering for a specific section."""
    try:
        doc = word_app.ActiveDocument
        section = doc.Sections(section_index)

        # Unlink from previous section
        try:
            section.Footers(1).LinkToPrevious = False
        except:
            pass

        # Clear existing footer content
        footer = section.Footers(1)
        footer.Range.Delete()

        # Insert page number
        footer.Range.ParagraphFormat.Alignment = 1  # Center alignment
        page_num = footer.PageNumbers.Add(
            PageNumberAlignment=1,  # wdAlignPageNumberCenter = 1
            FirstPage=True
        )

        # Set number style and starting number
        footer.PageNumbers.NumberStyle = number_style
        footer.PageNumbers.StartingNumber = start_number
        footer.PageNumbers.RestartNumberingAtSection = True

        style_name = "Roman" if number_style == 2 else "Arabic"
        print(f"  ✓ {style_name} numbering set for section {section_index} (starting at {start_number})")
        return True

    except Exception as e:
        print(f"  ERROR setting up numbering for section {section_index}: {str(e)}")
        return False


def get_section_page_count(word_app, section_index):
    """Get the number of pages in a specific section."""
    try:
        doc = word_app.ActiveDocument
        section = doc.Sections(section_index)

        # Get range of the section
        section_range = section.Range

        # Calculate pages by moving through the section
        start_page = section_range.Information(1)  # wdActiveEndPageNumber = 1
        end_page = section_range.Information(3)  # wdActiveEndPageNumber = 3

        return end_page - start_page + 1
    except:
        return 1


def create_full_book():
    """Main function to create the complete book."""
    print("=== Book Assembly Script Starting ===\n")

    # Verify all source files exist
    files_ok = True
    files_ok &= check_file_exists(FRONT_MATTER_PATH, "Front Matter")
    files_ok &= check_file_exists(BOOK_BODY_PATH, "Book Body")
    files_ok &= check_file_exists(INDEX_PATH, "Index")

    if not files_ok:
        print("\nPlease update the file paths and try again.")
        return False

    print(f"\nOutput will be saved to: {OUTPUT_PATH}\n")

    word_app = None
    doc = None

    try:
        # Initialize Word application
        print("1. Starting Word application...")
        word_app = create_word_app()
        if not word_app:
            return False

        # Create new document
        print("2. Creating master document...")
        doc = word_app.Documents.Add()
        time.sleep(1)

        # Insert Front Matter
        print("3. Inserting Front Matter...")
        if not insert_file_content(word_app, doc, FRONT_MATTER_PATH):
            return False

        # Add section break after front matter
        print("4. Adding section break after Front Matter...")
        if not insert_section_break(word_app):
            return False

        # Insert Book Body
        print("5. Inserting Book Body...")
        if not insert_file_content(word_app, doc, BOOK_BODY_PATH):
            return False

        # Add section break after book body
        print("6. Adding section break after Book Body...")
        if not insert_section_break(word_app):
            return False

        # Insert Index
        print("7. Inserting Index...")
        if not insert_file_content(word_app, doc, INDEX_PATH):
            return False

        # Wait for document to settle
        time.sleep(2)

        # Configure page numbering for each section
        print("8. Configuring page numbering...")

        # Section 1: Front Matter (Roman numerals starting at i)
        print("  Setting up Front Matter numbering (Roman)...")
        setup_page_numbering_simple(word_app, 1, 2, 1)  # wdPageNumberStyleLowercaseRoman = 2

        # Section 2: Book Body (Arabic numerals starting at 1)
        print("  Setting up Book Body numbering (Arabic)...")
        setup_page_numbering_simple(word_app, 2, 0, 1)  # wdPageNumberStyleArabic = 0

        # Section 3: Index (Roman numerals continuing from front matter)
        print("  Setting up Index numbering (Roman continuation)...")
        front_matter_pages = get_section_page_count(word_app, 1)
        index_start = front_matter_pages + 1
        setup_page_numbering_simple(word_app, 3, 2, index_start)  # Continue Roman

        # Save the document
        print("9. Saving final document...")

        # Create output directory if it doesn't exist
        output_dir = os.path.dirname(OUTPUT_PATH)
        os.makedirs(output_dir, exist_ok=True)

        # Remove existing file if it exists
        if os.path.exists(OUTPUT_PATH):
            os.remove(OUTPUT_PATH)

        # Save document
        doc.SaveAs2(OUTPUT_PATH)
        print(f"  ✓ Document saved successfully")

        # Close document
        doc.Close()
        word_app.Quit()

        print(f"\n✓ SUCCESS! Full book created at: {OUTPUT_PATH}")
        print("\n=== Manual Tasks Remaining ===")
        print("1. Review page number positioning and alignment")
        print("2. Check for any formatting issues at section boundaries")
        print("3. Update Table of Contents if present (References > Update Table)")
        print("4. Verify first page numbering preferences")
        print("5. Add running headers if needed")
        print("6. Final formatting review and adjustments")

        return True

    except Exception as e:
        print(f"\nERROR during book assembly: {str(e)}")
        try:
            if doc:
                doc.Close(SaveChanges=False)
            if word_app:
                word_app.Quit()
        except:
            pass
        return False


if __name__ == "__main__":
    print("Book Assembly Tool - Enhanced Version")
    print("====================================")

    success = create_full_book()

    if success:
        print(f"\n🎉 Book assembly completed successfully!")
        print(f"📖 Your complete book is ready at: {OUTPUT_PATH}")
    else:
        print(f"\n❌ Book assembly failed. Please check the error messages above.")
        print("\nTroubleshooting tips:")
        print("- Ensure Word is not already running")
        print("- Check that all source files are not open in Word")
        print("- Verify file paths are correct")
        print("- Try running as administrator")

    input("\nPress Enter to exit...")