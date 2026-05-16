# Modern EPUB Status

This document records the EPUB packaging and content-model changes that still apply after the documentation cleanup.

For commands and verification details, use [src/ebook/README_VERIFY.md](src/ebook/README_VERIFY.md).

## Current Applicable Changes

### OEBPS package root

The EPUB now uses an OEBPS-rooted package structure instead of the older default layout. Validation and verification logic were aligned with that structure.

### Split XHTML content model

The manuscript is no longer treated as a single `body.xhtml` file. The builder generates multiple XHTML files for front matter, visual navigation, and chapter content.

### Visual TOC built from actual chapter entries

The in-book TOC is generated from the chapter data actually emitted by the builder, which avoids stale links after chapter splitting.

### Front matter-driven title page

The title page is derived from the real front matter source rather than synthetic placeholder text.

### Image rewriting and embedding

Image references are rewritten to match the generated EPUB layout, and embedded assets are packaged under the EPUB image directories.

### Shared stylesheet

Generated XHTML files link to a shared stylesheet under the EPUB styles directory.

## What This File No Longer Tries to Be

This file is not the canonical quick-start guide and it does not document every exact generated filename. Those details can drift as the builder evolves.

Use it as a change summary, not as an operational checklist.
