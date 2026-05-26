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

import json
import os
import sys
import time
import importlib
from pathlib import Path

import win32com.client

if str(Path(__file__).resolve().parents[2]) not in sys.path:
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from src.utils.book_project import get_active_book_root, resolve_book_path

# Detect the root folder (book_contents_formatter_package)
ROOT = Path(__file__).resolve().parents[2]  # Go up 2 levels from src/final_book/
BOOK_ROOT = get_active_book_root(ROOT)


def load_assembler_config():
    config_path = ROOT / "src" / "final_book" / "full_book_assembler.config.json"
    with open(config_path, "r", encoding="utf-8") as handle:
        config = json.load(handle)

    inputs = config.get("inputs", {})
    outputs = config.get("outputs", {})
    output_dir = resolve_book_path(BOOK_ROOT, outputs["output_dir"])

    return {
        "config_path": config_path,
        "front_matter_file": resolve_book_path(BOOK_ROOT, inputs["front_matter_file"]),
        "book_body_file": resolve_book_path(BOOK_ROOT, inputs["book_body_file"]),
        "index_file": resolve_book_path(BOOK_ROOT, inputs["index_file"]),
        "output_file": output_dir / outputs.get("filename", "full_book.docx"),
        "pdf_file": output_dir / outputs.get("pdf_filename", "full_book.pdf"),
        "metadata_file": output_dir / outputs.get("metadata_filename", "full_book.config.json"),
        "options": config.get("options", {}),
    }


WD_COLLAPSE_END = 0
WD_COLLAPSE_START = 1
WD_SECTION_BREAK_NEXT_PAGE = 2
WD_HEADER_FOOTER_PRIMARY = 1
WD_ALIGN_PARAGRAPH_CENTER = 1
WD_PAGE_NUMBER_STYLE_ARABIC = 0
WD_PAGE_NUMBER_STYLE_LOWERCASE_ROMAN = 2
WD_FORMAT_ORIGINAL_FORMATTING = 16
WD_FORMAT_PDF = 17


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
        word_app.Visible = False
        word_app.DisplayAlerts = False  # Suppress alerts
        time.sleep(1)  # Give Word time to initialize
        return word_app
    except Exception as e:
        print(f"  ERROR creating Word application: {str(e)}")
        return None


def copy_section_layout(source_section, target_section):
    """Copy section page setup from one Word section to another."""
    source_setup = source_section.PageSetup
    target_setup = target_section.PageSetup

    simple_properties = [
        "TopMargin",
        "BottomMargin",
        "LeftMargin",
        "RightMargin",
        "Gutter",
        "HeaderDistance",
        "FooterDistance",
        "PageWidth",
        "PageHeight",
        "Orientation",
        "MirrorMargins",
        "DifferentFirstPageHeaderFooter",
        "OddAndEvenPagesHeaderFooter",
        "SectionStart",
        "VerticalAlignment",
        "SuppressEndnotes",
    ]

    for property_name in simple_properties:
        try:
            setattr(target_setup, property_name, getattr(source_setup, property_name))
        except Exception:
            continue

    try:
        source_columns = source_setup.TextColumns
        target_columns = target_setup.TextColumns
        target_columns.SetCount(source_columns.Count)
        target_columns.EvenlySpaced = source_columns.EvenlySpaced
        target_columns.LineBetween = source_columns.LineBetween
        if source_columns.Count > 1 and source_columns.EvenlySpaced:
            target_columns.Spacing = source_columns.Spacing
    except Exception:
        pass


def copy_section_page_size(source_section, target_section):
    """Copy only page dimensions and orientation from one section to another."""
    source_setup = source_section.PageSetup
    target_setup = target_section.PageSetup

    for property_name in ["PageWidth", "PageHeight", "Orientation"]:
        try:
            setattr(target_setup, property_name, getattr(source_setup, property_name))
        except Exception:
            continue


def copy_page_size_to_sections(document, reference_section_index, section_start, section_end):
    """Copy page size from one section to a consecutive range of sections."""
    for section_index in range(section_start, section_end + 1):
        copy_section_page_size(document.Sections(reference_section_index), document.Sections(section_index))


def ensure_page_numbering(section, start_number, number_style=WD_PAGE_NUMBER_STYLE_ARABIC):
    """Ensure a section has page numbering that restarts at a given number."""
    try:
        footer = section.Footers(WD_HEADER_FOOTER_PRIMARY)
        footer.LinkToPrevious = False
        footer.Range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER

        if footer.PageNumbers.Count == 0:
            footer.PageNumbers.Add(PageNumberAlignment=WD_ALIGN_PARAGRAPH_CENTER, FirstPage=True)

        footer.PageNumbers.RestartNumberingAtSection = True
        footer.PageNumbers.StartingNumber = start_number
        footer.PageNumbers.NumberStyle = number_style
        return True
    except Exception as exc:
        print(f"  WARNING unable to set page numbering: {exc}")
        return False


def apply_font_to_section(section, font_name):
    """Apply a font family to a section's content and footers without changing size or layout."""
    try:
        section.Range.Font.Name = font_name
    except Exception as exc:
        print(f"  WARNING unable to apply {font_name} to section body: {exc}")

    for footer_index in range(1, section.Footers.Count + 1):
        try:
            section.Footers(footer_index).Range.Font.Name = font_name
        except Exception:
            continue


def append_document_as_section(word_app, target_doc, source_path):
    """Append a source document while preserving source formatting and section layout."""
    try:
        print(f"  Appending {os.path.basename(source_path)} as a new section...")

        source_doc = word_app.Documents.Open(str(source_path), ReadOnly=True, AddToRecentFiles=False)
        existing_sections = int(target_doc.Sections.Count)
        source_sections = int(source_doc.Sections.Count)

        insertion_point = max(target_doc.Content.End - 1, 0)
        target_range = target_doc.Range(insertion_point, insertion_point)
        target_range.Collapse(WD_COLLAPSE_END)
        target_range.InsertBreak(Type=WD_SECTION_BREAK_NEXT_PAGE)

        source_doc.Content.Copy()
        time.sleep(0.5)

        target_range = target_doc.Range(max(target_doc.Content.End - 1, 0), max(target_doc.Content.End - 1, 0))
        target_range.Select()
        word_app.Selection.PasteAndFormat(WD_FORMAT_ORIGINAL_FORMATTING)

        expected_new_sections = existing_sections + source_sections
        actual_sections = int(target_doc.Sections.Count)
        if actual_sections < expected_new_sections:
            print(
                f"  WARNING section count after append was {actual_sections}; expected at least {expected_new_sections}"
            )

        copy_count = min(source_sections, max(actual_sections - existing_sections, 0))
        for offset in range(copy_count):
            source_section = source_doc.Sections(offset + 1)
            target_section = target_doc.Sections(existing_sections + offset + 1)
            copy_section_layout(source_section, target_section)

        source_doc.Close(False)

        print(f"  ✓ Successfully appended {os.path.basename(source_path)}")
        return {
            "source_sections": source_sections,
            "target_section_start": existing_sections + 1,
            "target_section_end": existing_sections + copy_count,
        }

    except Exception as e:
        print(f"  ERROR appending {os.path.basename(source_path)}: {str(e)}")
        try:
            source_doc.Close(False)
        except Exception:
            pass
        return None


def prepend_document_as_sections(word_app, target_doc, source_path):
    """Insert a source document before the current content while preserving source formatting."""
    try:
        print(f"  Prepending {os.path.basename(source_path)} before the body...")

        source_doc = word_app.Documents.Open(str(source_path), ReadOnly=True, AddToRecentFiles=False)
        existing_sections = int(target_doc.Sections.Count)
        source_sections = int(source_doc.Sections.Count)

        source_doc.Content.Copy()
        time.sleep(0.5)

        target_range = target_doc.Range(0, 0)
        target_range.Collapse(WD_COLLAPSE_START)
        target_range.Select()
        word_app.Selection.PasteAndFormat(WD_FORMAT_ORIGINAL_FORMATTING)

        inserted_end = word_app.Selection.Range.End
        boundary_range = target_doc.Range(inserted_end, inserted_end)
        boundary_range.Collapse(WD_COLLAPSE_END)
        boundary_range.InsertBreak(Type=WD_SECTION_BREAK_NEXT_PAGE)

        actual_sections = int(target_doc.Sections.Count)
        expected_new_sections = existing_sections + source_sections
        if actual_sections < expected_new_sections:
            print(
                f"  WARNING section count after prepend was {actual_sections}; expected at least {expected_new_sections}"
            )

        copy_count = min(source_sections, max(actual_sections - existing_sections, 0))
        for offset in range(copy_count):
            source_section = source_doc.Sections(offset + 1)
            target_section = target_doc.Sections(offset + 1)
            copy_section_layout(source_section, target_section)

        source_doc.Close(False)

        print(f"  ✓ Successfully prepended {os.path.basename(source_path)}")
        return {
            "source_sections": source_sections,
            "target_section_start": 1,
            "target_section_end": copy_count,
            "following_section_start": copy_count + 1,
        }

    except Exception as e:
        print(f"  ERROR prepending {os.path.basename(source_path)}: {str(e)}")
        try:
            source_doc.Close(False)
        except Exception:
            pass
        return None


def open_source_as_output(word_app, source_path, output_path):
    """Open the first source document and immediately save it as the output document."""
    try:
        print(f"  Using {os.path.basename(source_path)} as the base document...")
        document = word_app.Documents.Open(str(source_path), ReadOnly=False, AddToRecentFiles=False)
        document.SaveAs2(str(output_path))
        print(f"  ✓ Base document saved to {output_path}")
        return document
    except Exception as e:
        print(f"  ERROR preparing base document from {os.path.basename(source_path)}: {str(e)}")
        return None


def update_document_fields(document):
    """Refresh Word fields so TOC and cross-references see the merged content."""
    try:
        document.Repaginate()
        for table in document.TablesOfContents:
            table.Update()
        document.Fields.Update()
        print("  ✓ Updated fields and repaginated document")
        return True
    except Exception as e:
        print(f"  WARNING unable to update all fields: {str(e)}")
        return False


def export_pdf(document, pdf_path):
    """Export the assembled final book as a PDF."""
    try:
        if os.path.exists(pdf_path):
            os.remove(pdf_path)

        document.SaveAs2(str(pdf_path), FileFormat=WD_FORMAT_PDF)
        remove_blank_page_before_index(pdf_path)
        print(f"  ✓ PDF exported successfully: {pdf_path}")
        return True
    except Exception as e:
        print(f"  ERROR exporting PDF: {str(e)}")
        return False


def is_effectively_blank_pdf_page(page):
    """Treat pages with no meaningful extracted text as blank."""
    text = (page.extract_text() or "").strip()
    return not any(character.isalnum() for character in text)


def is_index_start_pdf_page(page):
    """Detect the first index page by its leading heading text."""
    text = (page.extract_text() or "").strip()
    normalized = " ".join(text.split())
    upper_text = normalized.upper()
    return "INDEX" in upper_text and ("0–9" in normalized or "0-9" in normalized)


def remove_blank_page_before_index(pdf_path):
    """Remove a blank PDF page immediately before the index start, if present."""
    pypdf = importlib.import_module("pypdf")
    pdf_reader_cls = pypdf.PdfReader
    pdf_writer_cls = pypdf.PdfWriter

    reader = pdf_reader_cls(str(pdf_path))
    index_page = None

    for page_index, page in enumerate(reader.pages):
        if is_index_start_pdf_page(page):
            index_page = page_index
            break

    if index_page is None or index_page == 0:
        return False

    previous_page = reader.pages[index_page - 1]
    if not is_effectively_blank_pdf_page(previous_page):
        return False

    writer = pdf_writer_cls()
    for page_index, page in enumerate(reader.pages):
        if page_index == index_page - 1:
            continue
        writer.add_page(page)

    temp_pdf_path = Path(pdf_path).with_suffix(".tmp.pdf")
    with open(temp_pdf_path, "wb") as handle:
        writer.write(handle)

    os.replace(temp_pdf_path, pdf_path)
    print("  ✓ Removed blank PDF page before index")
    return True


def collect_part_metadata(parts):
    return [
        {
            "label": label,
            "source_file": str(Path(path).relative_to(BOOK_ROOT)),
        }
        for label, path in parts
    ]


def create_full_book():
    """Main function to create the complete book."""
    print("=== Book Assembly Script Starting ===\n")

    config = load_assembler_config()
    front_matter_path = config["front_matter_file"]
    book_body_path = config["book_body_file"]
    index_path = config["index_file"]
    output_path = config["output_file"]
    pdf_path = config["pdf_file"]
    parts = [
        ("front_matter", front_matter_path),
        ("body", book_body_path),
        ("index", index_path),
    ]

    # Verify all source files exist
    files_ok = True
    files_ok &= check_file_exists(front_matter_path, "Front Matter")
    files_ok &= check_file_exists(book_body_path, "Book Body")
    files_ok &= check_file_exists(index_path, "Index")

    if not files_ok:
        print("\nPlease update the file paths and try again.")
        return False

    print(f"\nOutput will be saved to: {output_path}\n")

    word_app = None
    doc = None

    try:
        # Initialize Word application
        print("1. Starting Word application...")
        word_app = create_word_app()
        if not word_app:
            return False

        # Create output directory if it doesn't exist
        output_dir = os.path.dirname(output_path)
        os.makedirs(output_dir, exist_ok=True)

        # Remove existing file if it exists
        if os.path.exists(output_path):
            os.remove(output_path)

        print("2. Creating master document from Book Body...")
        doc = open_source_as_output(word_app, book_body_path, output_path)
        if not doc:
            return False
        time.sleep(1)

        print("3. Prepending Front Matter...")
        front_matter_prepend_result = prepend_document_as_sections(word_app, doc, front_matter_path)
        if not front_matter_prepend_result:
            return False
        body_section_start = front_matter_prepend_result["following_section_start"]
        ensure_page_numbering(doc.Sections(1), start_number=1, number_style=WD_PAGE_NUMBER_STYLE_LOWERCASE_ROMAN)
        ensure_page_numbering(doc.Sections(body_section_start), start_number=1)
        copy_page_size_to_sections(doc, body_section_start, 1, body_section_start - 1)

        print("4. Appending Index...")
        index_append_result = append_document_as_section(word_app, doc, index_path)
        if not index_append_result:
            return False
        ensure_page_numbering(
            doc.Sections(index_append_result["target_section_start"]),
            start_number=1,
            number_style=WD_PAGE_NUMBER_STYLE_LOWERCASE_ROMAN,
        )
        apply_font_to_section(doc.Sections(index_append_result["target_section_start"]), "Georgia")
        reference_section_index = max(index_append_result["target_section_start"] - 1, 1)
        copy_page_size_to_sections(
            doc,
            reference_section_index,
            index_append_result["target_section_start"],
            index_append_result["target_section_end"],
        )

        # Wait for document to settle
        time.sleep(2)

        print("5. Refreshing pagination and fields...")
        update_document_fields(doc)

        # Save the document
        print("6. Saving final document...")
        doc.SaveAs2(str(output_path))
        print("  ✓ Document saved successfully")

        print("7. Exporting PDF version...")
        if not export_pdf(doc, pdf_path):
            return False

        section_count = int(doc.Sections.Count)

        metadata_file = config["metadata_file"]
        metadata_file.parent.mkdir(parents=True, exist_ok=True)
        with open(metadata_file, "w", encoding="utf-8") as handle:
            json.dump(
                {
                    "status": "success",
                    "config_file": str(config["config_path"].relative_to(ROOT)),
                    "output_file": str(output_path.relative_to(BOOK_ROOT)),
                    "pdf_file": str(pdf_path.relative_to(BOOK_ROOT)),
                    "part_order": [label for label, _ in parts],
                    "parts": collect_part_metadata(parts),
                    "sections_in_output": section_count,
                    "preserve_section_formatting": bool(config["options"].get("preserve_section_formatting", True)),
                    "body_section_start": body_section_start,
                    "index_section_start": index_append_result["target_section_start"],
                },
                handle,
                indent=2,
            )

        # Close document
        doc.Close()
        word_app.Quit()

        print(f"\n✓ SUCCESS! Full book created at: {output_path}")
        print(f"  PDF version created at: {pdf_path}")
        print(f"  Sections detected in output: {section_count}")
        print("\n=== Manual Review ===")
        print("1. Verify front matter, body, and index appear in that order")
        print("2. Check page setup and headers/footers at section boundaries")
        print("3. Confirm TOC and index fields after opening the saved document in Word")

        return True

    except Exception as e:
        print(f"\nERROR during book assembly: {str(e)}")
        try:
            if doc:
                doc.Close(SaveChanges=False)
            if word_app:
                word_app.Quit()
        except Exception:
            pass
        return False


if __name__ == "__main__":
    print("Book Assembly Tool - Enhanced Version")
    print("====================================")

    success = create_full_book()

    if success:
        print("\n🎉 Book assembly completed successfully!")
        print(f"📖 Your complete book is ready at: {load_assembler_config()['output_file']}")
    else:
        print("\n❌ Book assembly failed. Please check the error messages above.")
        print("\nTroubleshooting tips:")
        print("- Ensure Word is not already running")
        print("- Check that all source files are not open in Word")
        print("- Verify file paths are correct")
        print("- Try running as administrator")

