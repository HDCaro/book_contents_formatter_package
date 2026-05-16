# Data Structure and Configuration Model

## Overview

The repository uses a centralized, task-based data organization with config-driven workflows.

The main goals are:

1. keep source files separate from generated outputs,
2. route workflows through JSON configuration instead of hardcoded paths,
3. allow downstream steps to reuse metadata from earlier stages.

## Folder Structure

```text
data/
  inputs/         source files and assets
  outputs/        generated task outputs
  intermediate/   extracted JSON and working datasets
```

Specialized index data also exists under `data/index/` for the curated index pipeline.

## Main Config Files

| Config File                        | Location                     | Purpose                                            |
| ---------------------------------- | ---------------------------- | -------------------------------------------------- |
| `front_matter_builder.config.json` | `src/front_matter/builders/` | Configure front matter build inputs and outputs    |
| `index_builder.config.json`        | `src/index/build/`           | Configure index build steps and dependencies       |
| `full_book_assembler.config.json`  | `src/final_book/`            | Configure final Word assembly                      |
| `ebook_builder.config.json`        | `src/ebook/`                 | Configure EPUB build inputs, metadata, and outputs |

## Config Pattern

Most config files follow the same structure:

```json
{
  "inputs": {},
  "outputs": {},
  "options": {}
}
```

Many also include task metadata or descriptive fields.

## Metadata Flow

When a workflow emits metadata, downstream steps can use it to avoid re-entering path or pagination information.

Typical uses include:

- front matter metadata feeding later assembly steps,
- index outputs feeding later book assembly,
- config-driven reuse of source and output paths.

## Notes

- Paths are usually relative to project root.
- Generated output folders should be treated as artifacts, not hand-maintained source content.
- Some older docs describe this structure as a new refactor; that language is historical.
- The current canonical overview is this file plus [README.md](README.md).
