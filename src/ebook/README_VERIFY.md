# EPUB eBook Builder with Verification System

## Overview

The ebook builder generates production-ready EPUB files from Microsoft Word DOCX inputs (front matter + book body). It includes a comprehensive verification system (`--verify`) to catch KDP compatibility issues before upload.

## Quick Start

### Build EPUB

```bash
python src/ebook/ebook_builder.py
```

Uses configuration from `src/ebook/ebook_builder.config.json`

### Verify EPUB

```bash
python src/ebook/ebook_builder.py --verify
```

Verifies the output EPUB file specified in config

### Verify Specific EPUB

```bash
python src/ebook/ebook_builder.py --verify path/to/MyBook.epub
```

## Configuration

Edit `src/ebook/ebook_builder.config.json`:

```json
{
  "inputs": {
    "front_matter_docx": "path/to/front_matter.docx",
    "book_body_file": "path/to/book_body.docx",
    "cover_image_file": "optional/cover.jpg",
    "back_cover_image_file": "optional/back_cover.jpg",
    "auto_discover_missing_inputs": true
  },
  "outputs": {
    "output_dir": "data/outputs/02_epub",
    "filename": "MyBook.epub",
    "temp_dir": "data/outputs/02_epub/temp"
  },
  "metadata": {
    "title": "My Book Title",
    "author": "Author Name",
    "language": "es",
    "identifier": ""
  },
  "options": {
    "keep_html_exports": true,
    "fallback_from_front_matter_config": true
  }
}
```

## Verification Report

The `--verify` command generates a comprehensive compatibility report:

### Location

```
{output_dir}/{filename}_verification_report.txt
```

### Report Contents

1. **Summary**
   - Pass/Fail status
   - Error, Warning, and Info counts
   - Severity score (0-100+)

2. **Critical Errors**
   - Issues that must be fixed before KDP upload
   - Example: broken image references, missing metadata

3. **Warnings**
   - Potential issues that may cause reader display problems
   - Example: excessive inline styling, character encoding risk

4. **Information**
   - Success confirmations and statistics
   - Example: "✓ Found 38 chapters", "✓ All 320 images embedded"

5. **Statistics**
   - Chapters count
   - Embedded images count
   - TOC entries
   - Inline spans count
   - Character encoding analysis per file

### Example Report Section

```
================================================================================
EBOOK VERIFICATION REPORT
================================================================================
Generated: 2026-05-11 13:14:17
Status: ✓ PASS - No critical issues found

SUMMARY:
  Errors:    0
  Warnings:  2
  Info:      10
  Severity Score: 20/100

INFORMATION:
  ✓ Found 38 chapters (H1 tags)
  ✓ Found 320 embedded images
  ✓ All image references valid
  ✓ TOC has 39 entries
  ✓ Body content: 1.3M characters
```

## Verification Checks

### 1. Character Encoding (mojibake detection)

- Detects smart quotes (`"`, `"`, `'`, `'`)
- Detects fancy dashes (`–`, `—`)
- Detects special spaces (non-breaking, etc.)
- Risk levels: LOW, MEDIUM, HIGH
- Recommendation: KDP may need normalization for compatibility

### 2. HTML Structure (semantic validation)

- Verifies H1 tags for chapter structure
- Detects excessive inline styling (Word bloat)
- Flags headers/footers (not allowed in eBooks)
- Checks for page breaks (will be ignored by KDP)

### 3. Image References

- Verifies all images embedded in EPUB
- Checks image paths are valid and resolvable
- Supports both relative and absolute paths
- Reports broken references

### 4. KDP Compliance

- Validates required metadata (title, author)
- Checks for reflowable CSS
- Verifies Table of Contents exists
- Ensures proper manifest structure

### 5. Content Completeness

- Confirms all chapters present
- Checks body content size (warns if suspiciously small)
- Verifies manuscript structure

## Common Issues & Solutions

### ⚠️ "Found XXX inline-styled spans"

**Issue**: Word exports contain excessive styling
**Solution**:

- Usually harmless for reading
- Can clean with external tools if preferred
- KDP will handle gracefully

### ⚠️ "Potentially problematic characters detected"

**Issue**: Smart quotes, fancy dashes may not render correctly on some devices
**Solution**:

- Option 1: Upload as-is (most Kindle devices handle fine)
- Option 2: Re-export from Word using "Replace" feature
  - Find: `"` (straight quotes)
  - Replace: `"` (smart quotes) - reverse to normalize
- Option 3: Future feature - use `"character_mode": "kdp-safe"` in config

### ❌ "Broken image reference"

**Issue**: HTML references an image that's not embedded
**Solution**:

- Verify images were properly imported in Word
- Re-export Word document using filtered HTML format
- Check that all images appear in the source DOCX

### ❌ "No <h1> tags found"

**Issue**: Chapters not structured with Heading 1 style in Word
**Solution**:

- Format chapter titles in Word with Heading 1 style
- Ensure Heading 1 styles are preserved in export
- Re-export and rebuild EPUB

## Build Process

1. **Word Preflight** - Terminates any running Word instances
2. **Export Front Matter** - DOCX → filtered HTML via Word COM
3. **Export Body** - DOCX → filtered HTML via Word COM
4. **Extract Headings** - Reads chapter structure from Word
5. **Anchor Headings** - Injects navigation anchors in HTML
6. **Collect Images** - Finds and embeds image assets
7. **Create EPUB** - Builds final package with manifest, TOC, spine
8. **Generate Report** - Optional verification report

## Technical Details

### Parser Choice

Uses `html.parser` instead of `lxml` for tolerant HTML handling of Word's messy output.

### Image Handling

- Scans Word HTML for `<img>` tags with relative paths
- Locates actual image files in temporary export folders
- Embeds into EPUB under `EPUB/images/embedded/`
- Rewrites all references for EPUB navigation

### Heading Detection

- Reads Heading 1 for chapter markers
- Collects consecutive Heading 2 for subtitles
- Preserves heading hierarchy in TOC
- Injects anchor IDs for linked navigation

## Advanced Usage

### Automation Script

```bash
# Build and verify in one step
python src/ebook/ebook_builder.py
python src/ebook/ebook_builder.py --verify
```

### Batch Processing

```bash
# Process multiple configs
for config in configs/*.json; do
    echo "Processing $config..."
    python src/ebook/ebook_builder.py --config "$config"
    python src/ebook/ebook_builder.py --verify
done
```

## Output Files

```
data/outputs/02_epub/
├── MyBook.epub                                # Final EPUB file
├── MyBook_verification_report.txt             # Verification report
└── temp/                                      # Temporary exports
    ├── front_matter.html                      # Front matter HTML
    └── book_body.html                         # Body HTML
```

## Troubleshooting

### Word Not Found / COM Errors

- Ensure Microsoft Word is installed
- Check user has permission to launch Word
- Verify `win32com` package is installed: `pip install pywin32`

### Image Extraction Fails

- Confirm images are embedded in Word (not linked)
- Check images are supported formats: JPG, PNG, GIF, WebP
- Try re-inserting images in Word

### Missing Chapters

- Verify all chapters use Heading 1 style in Word
- Check chapter text is not in text boxes or floating objects
- Ensure no complex nested tables
- Try rebuilding with fresh Word export

## Output Formats

### EPUB Structure

```
EPUB/
├── body.xhtml                  # Main manuscript content
├── front.xhtml                 # Front matter
├── back_cover.xhtml            # Back cover (if provided)
├── images/embedded/            # All embedded images
│   ├── 0001_image001.jpg
│   └── ...
├── style/nav.css               # Navigation styling
├── META-INF/container.xml      # EPUB container
└── mimetype                    # EPUB marker file
```

### Metadata Handled

- Title (required)
- Author (recommended)
- Language (default: es)
- Unique Identifier (auto-generated if not provided)
- Cover image (optional)
- Back cover image (optional)

## KDP Upload Recommendations

1. ✓ Run `--verify` before uploading
2. ✓ Fix all CRITICAL ERRORS
3. ⚠️ Review WARNINGS for potential device compatibility
4. 📝 Test final EPUB in Kindle Previewer
5. 🎯 Test on actual Kindle devices if possible
6. 📤 Upload to KDP with confidence

---

**Version**: 1.0  
**Last Updated**: 2026-05-11
