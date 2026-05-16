# Book Body Heading Guide for Index and TOC Extraction

This document describes the manuscript structure expected by the index and contents-extraction workflows. It no longer serves as a command reference for a single root-level script.

## What This Guide Is For

Use this guide when preparing a Word manuscript so the repository can extract:

1. chapter-level headings,
2. section-level headings,
3. a table-of-contents source structure,
4. index-related structural signals.

## Required Heading Structure

Your book DOCX should use real Word heading styles:

- `Heading 1` for chapters or major sections
- `Heading 2` for sections or chapter subtitles
- `Heading 3` for deeper subsections when needed

Do not replace heading styles with manual formatting such as bold text, larger font size, or custom spacing.

## Recommended Manuscript Pattern

```text
Heading 1: Chapter 1
Heading 2: Introduction

Heading 1: Chapter 2
Heading 2: Getting Started

Heading 2: Setting Up Your Environment
Heading 2: Basic Configuration

Heading 1: Chapter 3
Heading 2: Advanced Topics
```

## Why This Matters

The extraction workflows depend on structural signals from Word, not visual appearance alone.

Correct heading usage improves:

- chapter detection,
- section grouping,
- TOC generation,
- downstream page mapping,
- consistency between Word and EPUB outputs.

## Common Problems

### No headings found

Cause: chapter and section titles were styled manually instead of using Word heading styles.

### Wrong hierarchy

Cause: `Heading 1`, `Heading 2`, and `Heading 3` were applied inconsistently.

### Duplicated or confusing chapter titles

Cause: chapter labels and subtitles were merged into a single inconsistent heading pattern.

## Scope Notes

- Page numbers may still need correction later in the pipeline.
- This guide describes input preparation, not the full curated index pipeline.
- For end-to-end index processing, use [README_index.md](README_index.md).
- For general user setup, use [README_USER.md](README_USER.md).
