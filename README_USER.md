# Book Contents Formatter Package - User Guide

This README is for using the tools in this repository without needing to understand the internal architecture.

If you want the technical/developer view of the repository, read [README.md](README.md).

## What You Can Do With This Project

This project helps you work with book production files based on Microsoft Word documents.

Main user-facing workflows:

1. Prepare Word source files so headings, front matter, and assets are usable by the automation.
2. Build an index or table-of-contents source from a structured Word manuscript.
3. Generate a styled EPUB from Word book files.

## System Requirements

- Windows is recommended for the Word-based workflows.
- Microsoft Word must be installed for COM-based automation.
- Python and the dependencies from `requirements.txt` must be installed.

Basic setup:

```bash
pip install -r requirements.txt
```

## DOCX File Requirements

Your source Word files need to be structured correctly or the automation will not produce good results.

### Book Body DOCX

The main manuscript file should use Word heading styles consistently.

Recommended structure:

```text
Heading 1: Chapter 1
Heading 2: The Beginning of Everything

Heading 1: Chapter 2
Heading 2: A New Dawn

Heading 1: INTRODUCTION

Heading 1: RICHARD NILES DISCOGRAPHY BY YEAR
```

For index/TOC extraction workflows:

- `Heading 1` for chapters or major sections.
- `Heading 2` for chapter titles or subsections.
- `Heading 3` for deeper subsections when needed.

Avoid:

- putting the full chapter label and subtitle in an inconsistent style,
- using plain bold text instead of heading styles,
- mixing heading levels randomly.

### Front Matter DOCX

For EPUB and front matter workflows, the front matter DOCX should contain the real title page design you want to reuse.

Recommended contents:

- title page,
- author name,
- optional imprint/logo,
- clean layout with embedded images rather than linked remote assets.

### Copyright DOCX

If you provide a separate copyright page DOCX, it should contain the final copyright and publication details you want included in the output.

### Image Assets

If your Word files contain images:

- embed them in the DOCX,
- do not rely on external links,
- use standard formats such as JPG or PNG.

## Main Workflows

### 1. Prepare a Book Body DOCX for Extraction

If your manuscript uses proper heading styles, you can extract headings and generate Roman numeral front matter/index output.

See:

- [book_indexer_guide.md](book_indexer_guide.md)

### 2. Build an EPUB

If you want an EPUB from your Word files:

```bash
python src/ebook/ebook_builder.py
```

To verify the EPUB:

```bash
python src/ebook/ebook_builder.py --verify
```

Detailed EPUB usage:

- [src/ebook/README_VERIFY.md](src/ebook/README_VERIFY.md)

## EPUB Input Expectations

The EPUB builder typically expects:

- a front matter DOCX,
- a main book body DOCX,
- an optional copyright DOCX,
- an optional cover image,
- an optional back cover image.

These are configured in:

- `src/ebook/ebook_builder.config.json`

Typical config input keys:

- `front_matter_docx`
- `book_body_file`
- `copyright_page_docx`
- `cover_image_file`
- `back_cover_image_file`

## Expected Outputs

Depending on the workflow, outputs may include:

- formatted Word contents documents,
- combined Word front matter,
- generated indexes,
- EPUB files,
- EPUB verification reports.

Most generated files are written under `data/outputs/`.

## Troubleshooting

Common problems and likely causes:

1. No headings found.
   Cause: your DOCX is not using Word heading styles correctly.

2. Wrong TOC/index structure.
   Cause: inconsistent Heading 1 / Heading 2 usage in the manuscript.

3. Images missing in EPUB.
   Cause: images were linked instead of embedded in the DOCX, or the input paths are wrong.

4. Word automation failures.
   Cause: Microsoft Word is not installed, or the workflow is not running on Windows.

## Recommended Reading Order

If you are a user and not editing code:

1. Start with [README_USER.md](README_USER.md).
2. For Word TOC/index generation, read [book_indexer_guide.md](book_indexer_guide.md).
3. For EPUB generation, read [src/ebook/README_VERIFY.md](src/ebook/README_VERIFY.md).

## In Short

To get good results, make sure your DOCX files are clean, use real Word heading styles, and keep front matter/copyright/cover assets organized. The tools work best when the source documents already reflect the logical structure of the book.
