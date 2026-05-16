# EPUB CSS Styling Guide

This file documents the styling assumptions that still matter for the current EPUB builder without hard-coding stale filename examples.

For build and verification commands, use [src/ebook/README_VERIFY.md](src/ebook/README_VERIFY.md).

## Current Styling Model

The EPUB builder applies a shared stylesheet to generated XHTML content inside the EPUB package.

The styling goals are:

- readable long-form body text,
- clear chapter and section hierarchy,
- stable spacing across major readers,
- safe handling of images, tables, and block elements,
- compatibility with reflowable EPUB readers.

## Styling Characteristics

### Typography

- serif body text for long reading,
- sans-serif emphasis where needed for hierarchy,
- readable line height,
- consistent spacing between paragraphs and headings.

### Layout

- chapter and section separation through margins and page-break hints,
- restrained padding for reading comfort,
- indentation and spacing that survive reader-specific overrides.

### Special elements

- blockquotes remain visually distinct,
- tables are styled for legibility,
- images are centered and constrained for reflowable layouts,
- links remain visible without overwhelming body text.

## Asset and Link Assumptions

- generated XHTML files link to a shared stylesheet,
- image references are rewritten relative to the generated XHTML location,
- styling is intended to work with the current OEBPS-based package layout.

## Reader Variability

EPUB readers do not render CSS identically. Treat the stylesheet as a compatibility-oriented baseline, not pixel-perfect presentation.

The main target is stable rendering in common reflowable readers such as Kindle ingestion, Apple Books, and Kobo-class readers.

## What No Longer Applies

This file no longer assumes:

- a single `body.xhtml` content file,
- exact front-matter filenames such as `front.xhtml`,
- a fixed generated file inventory beyond the shared stylesheet and rewritten assets.
