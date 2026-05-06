"""
extract_front_matter_format.py

Simplified script to extract formatting parameters from ONLY the front matter document.
This will serve as our baseline/reference format for the complete book.

Requirements: pywin32 (pip install pywin32), Word 365 installed
"""

import win32com.client
import json
import os
import sys
import time
import gc
from datetime import datetime

# Single file to analyze - front matter as reference
FRONT_MATTER_PATH = r"C:\Projects\Python\book_contents_formatter_package\release\v1\HITS_AND_HAPPINESS_FINAL_FRONT_MATTER.docx"
OUTPUT_JSON = r"C:\Projects\Python\book_contents_formatter_package\src\final_book\front_matter_format_baseline.json"


def points_to_inches(points):
    """Convert points to inches (72 points = 1 inch)."""
    try:
        return round(float(points) / 72, 3)
    except:
        return 0.0


def extract_complete_page_setup(doc):
    """Extract comprehensive page setup from the front matter."""
    print(f"    🔄 Extracting complete page setup...")

    try:
        page_setup = doc.Sections[0].PageSetup

        setup_info = {
            "paper_size": {
                "width_inches": points_to_inches(page_setup.PageWidth),
                "height_inches": points_to_inches(page_setup.PageHeight),
                "width_points": float(page_setup.PageWidth),
                "height_points": float(page_setup.PageHeight),
                "paper_size_name": "Custom" if page_setup.PaperSize == 41 else f"Size_{page_setup.PaperSize}"
            },
            "margins": {
                "top_inches": points_to_inches(page_setup.TopMargin),
                "bottom_inches": points_to_inches(page_setup.BottomMargin),
                "left_inches": points_to_inches(page_setup.LeftMargin),
                "right_inches": points_to_inches(page_setup.RightMargin),
                "gutter_inches": points_to_inches(page_setup.Gutter),
                "header_distance_inches": points_to_inches(page_setup.HeaderDistance),
                "footer_distance_inches": points_to_inches(page_setup.FooterDistance),
                # Raw points for exact matching
                "top_points": float(page_setup.TopMargin),
                "bottom_points": float(page_setup.BottomMargin),
                "left_points": float(page_setup.LeftMargin),
                "right_points": float(page_setup.RightMargin),
                "gutter_points": float(page_setup.Gutter)
            },
            "orientation": "Portrait" if page_setup.Orientation == 0 else "Landscape",
            "orientation_code": int(page_setup.Orientation),
            "layout_settings": {
                "mirror_margins": bool(page_setup.MirrorMargins),
                "different_first_page": bool(page_setup.DifferentFirstPageHeaderFooter),
                "different_odd_even": bool(page_setup.OddAndEvenPagesHeaderFooter),
                "vertical_alignment": int(page_setup.VerticalAlignment),
                "section_start": int(page_setup.SectionStart),
                "suppress_endnotes": bool(page_setup.SuppressEndnotes)
            }
        }

        print(
            f"    ✓ Page setup: {setup_info['paper_size']['width_inches']}\" x {setup_info['paper_size']['height_inches']}\"")
        print(
            f"    ✓ Margins: T:{setup_info['margins']['top_inches']}\", B:{setup_info['margins']['bottom_inches']}\", L:{setup_info['margins']['left_inches']}\", R:{setup_info['margins']['right_inches']}\"")

        return setup_info

    except Exception as e:
        print(f"    ❌ Error extracting page setup: {str(e)}")
        return {"error": str(e)}


def extract_key_styles(doc):
    """Extract the most important styles with full details."""
    print(f"    🔄 Extracting key document styles...")

    styles_info = {}

    # Key styles to extract
    important_styles = [
        "Normal", "Heading 1", "Heading 2", "Heading 3", "Heading 4",
        "Title", "Subtitle", "Body Text", "Body Text Indent",
        "Caption", "Header", "Footer", "Page Number"
    ]

    try:
        for style_name in important_styles:
            try:
                style = doc.Styles[style_name]

                style_info = {
                    "exists": True,
                    "type": int(style.Type),
                    "type_name": "Paragraph" if style.Type == 1 else "Character" if style.Type == 2 else "Other",
                    "built_in": bool(style.BuiltIn),
                    "in_use": bool(style.InUse)
                }

                # Font information
                try:
                    font = style.Font
                    style_info["font"] = {
                        "name": str(font.Name),
                        "size_points": float(font.Size),
                        "bold": bool(font.Bold),
                        "italic": bool(font.Italic),
                        "underline": int(font.Underline),
                        "color_rgb": int(font.Color),
                        "small_caps": bool(font.SmallCaps),
                        "all_caps": bool(font.AllCaps)
                    }
                except Exception as e:
                    style_info["font"] = {"error": str(e)}

                # Paragraph format (for paragraph styles)
                if style.Type == 1:  # Paragraph style
                    try:
                        pf = style.ParagraphFormat
                        style_info["paragraph_format"] = {
                            "alignment": int(pf.Alignment),
                            "alignment_name": ["Left", "Center", "Right", "Justify"][
                                pf.Alignment] if pf.Alignment < 4 else f"Alignment_{pf.Alignment}",
                            "line_spacing": float(pf.LineSpacing),
                            "line_spacing_rule": int(pf.LineSpacingRule),
                            "space_before_inches": points_to_inches(pf.SpaceBefore),
                            "space_after_inches": points_to_inches(pf.SpaceAfter),
                            "space_before_points": float(pf.SpaceBefore),
                            "space_after_points": float(pf.SpaceAfter),
                            "first_line_indent_inches": points_to_inches(pf.FirstLineIndent),
                            "left_indent_inches": points_to_inches(pf.LeftIndent),
                            "right_indent_inches": points_to_inches(pf.RightIndent),
                            "keep_together": bool(pf.KeepTogether),
                            "keep_with_next": bool(pf.KeepWithNext),
                            "widow_control": bool(pf.WidowControl)
                        }
                    except Exception as e:
                        style_info["paragraph_format"] = {"error": str(e)}

                styles_info[style_name] = style_info
                print(
                    f"      ✓ {style_name}: {style_info.get('font', {}).get('name', 'N/A')} {style_info.get('font', {}).get('size_points', 'N/A')}pt")

            except:
                # Style doesn't exist
                styles_info[style_name] = {"exists": False}

    except Exception as e:
        styles_info["extraction_error"] = str(e)

    print(
        f"    ✓ Extracted {len([s for s in styles_info.values() if isinstance(s, dict) and s.get('exists', False)])} styles")
    return styles_info


def extract_section_details(doc):
    """Extract detailed information about each section."""
    print(f"    🔄 Extracting detailed section information...")

    sections_info = []

    try:
        for i in range(doc.Sections.Count):
            section = doc.Sections[i]
            ps = section.PageSetup

            section_info = {
                "section_number": i + 1,
                "page_setup": {
                    "margins": {
                        "top_inches": points_to_inches(ps.TopMargin),
                        "bottom_inches": points_to_inches(ps.BottomMargin),
                        "left_inches": points_to_inches(ps.LeftMargin),
                        "right_inches": points_to_inches(ps.RightMargin),
                        "gutter_inches": points_to_inches(ps.Gutter)
                    },
                    "orientation": "Portrait" if ps.Orientation == 0 else "Landscape",
                    "different_first_page": bool(ps.DifferentFirstPageHeaderFooter),
                    "different_odd_even": bool(ps.OddAndEvenPagesHeaderFooter),
                    "section_start": int(ps.SectionStart)
                },
                "headers_footers": {
                    "header_count": section.Headers.Count,
                    "footer_count": section.Footers.Count,
                    "has_primary_header": section.Headers.Count > 0,
                    "has_primary_footer": section.Footers.Count > 0
                }
            }

            sections_info.append(section_info)
            print(
                f"      ✓ Section {i + 1}: {section_info['page_setup']['orientation']}, margins T:{section_info['page_setup']['margins']['top_inches']}\"")

    except Exception as e:
        sections_info.append({"error": f"Could not extract section details: {str(e)}"})

    return sections_info


def create_format_baseline():
    """Create comprehensive format baseline from front matter."""
    print(f"=== Front Matter Format Baseline Extractor ===\n")

    if not os.path.exists(FRONT_MATTER_PATH):
        print(f"❌ Front matter file not found: {FRONT_MATTER_PATH}")
        return False

    print(f"📄 Analyzing: {os.path.basename(FRONT_MATTER_PATH)}")

    word_app = None
    doc = None

    try:
        print(f"  🔄 Starting Word application...")
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = False
        word_app.ScreenUpdating = False
        print(f"  ✓ Word application ready")

        print(f"  🔄 Opening front matter document...")
        doc = word_app.Documents.Open(
            FileName=FRONT_MATTER_PATH,
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False
        )
        print(f"  ✓ Document opened successfully")

        # Extract basic document statistics
        print(f"  🔄 Gathering document statistics...")
        doc_stats = {
            "file_info": {
                "file_path": FRONT_MATTER_PATH,
                "file_name": os.path.basename(FRONT_MATTER_PATH),
                "analysis_timestamp": datetime.now().isoformat()
            },
            "statistics": {
                "page_count": int(doc.ComputeStatistics(2)),  # wdStatisticPages
                "word_count": int(doc.ComputeStatistics(0)),  # wdStatisticWords
                "character_count": int(doc.ComputeStatistics(3)),  # wdStatisticCharacters
                "paragraph_count": int(doc.Paragraphs.Count),
                "section_count": int(doc.Sections.Count)
            }
        }
        print(
            f"  ✓ Stats: {doc_stats['statistics']['page_count']} pages, {doc_stats['statistics']['section_count']} sections, {doc_stats['statistics']['word_count']} words")

        # Extract comprehensive formatting
        page_setup = extract_complete_page_setup(doc)
        sections = extract_section_details(doc)
        styles = extract_key_styles(doc)

        # Combine all information
        baseline = {
            "baseline_info": {
                "created_from": "front_matter",
                "purpose": "Reference format for complete book assembly",
                "extraction_version": "1.0"
            },
            "document_info": doc_stats,
            "page_setup": page_setup,
            "sections": sections,
            "styles": styles,
            "validation_rules": {
                "required_margins": page_setup.get("margins", {}),
                "required_paper_size": page_setup.get("paper_size", {}),
                "required_orientation": page_setup.get("orientation", "Portrait"),
                "key_fonts": {
                    style_name: style_data.get("font", {}).get("name", "Unknown")
                    for style_name, style_data in styles.items()
                    if isinstance(style_data, dict) and style_data.get("exists", False)
                }
            }
        }

        # Save baseline
        os.makedirs(os.path.dirname(OUTPUT_JSON), exist_ok=True)
        with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
            json.dump(baseline, f, indent=2, ensure_ascii=False)

        print(f"\n✅ SUCCESS! Format baseline created:")
        print(f"📁 {OUTPUT_JSON}")

        # Print key formatting summary
        print(f"\n📋 FORMAT BASELINE SUMMARY:")
        if "error" not in page_setup:
            print(
                f"  📏 Paper Size: {page_setup['paper_size']['width_inches']}\" × {page_setup['paper_size']['height_inches']}\"")
            print(
                f"  📐 Margins: T:{page_setup['margins']['top_inches']}\", B:{page_setup['margins']['bottom_inches']}\", L:{page_setup['margins']['left_inches']}\", R:{page_setup['margins']['right_inches']}\"")
            print(f"  📄 Orientation: {page_setup['orientation']}")

        print(
            f"  📊 Document: {doc_stats['statistics']['page_count']} pages, {doc_stats['statistics']['section_count']} sections")

        active_styles = [name for name, data in styles.items() if isinstance(data, dict) and data.get("exists", False)]
        print(f"  🎨 Active Styles: {len(active_styles)} found")

        return True

    except Exception as e:
        print(f"  ❌ Error creating baseline: {str(e)}")
        return False

    finally:
        print(f"  🔄 Cleaning up...")
        try:
            if doc:
                doc.Close(SaveChanges=False)
                print(f"    ✓ Document closed")
        except:
            pass

        try:
            if word_app:
                word_app.Quit()
                print(f"    ✓ Word application closed")
        except:
            pass

        gc.collect()
        print(f"  ✓ Cleanup completed")


if __name__ == "__main__":
    success = create_format_baseline()

    if success:
        print(f"\n🎉 Format baseline extraction completed!")
        print(f"📋 Next steps:")
        print(f"   1. Review the baseline JSON file")
        print(f"   2. Use this as reference for validating other documents")
        print(f"   3. Create validation script using this baseline")
    else:
        print(f"\n❌ Baseline extraction failed.")

    input("\nPress Enter to exit...")