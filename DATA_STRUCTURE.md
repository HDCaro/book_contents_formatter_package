# Data Structure Refactor - Configuration and File Organization

## Overview

The project now uses a centralized, task-based data organization with three distinct phases:

1. **Front Matter Builder** (Task 1)
2. **Index Builder** (Task 2)
3. **Full Book Assembler** (Task 3)

Each task reads its config from a JSON file and manages its own inputs and outputs.

## Folder Structure

```
data/
├── inputs/                           (READ-ONLY source files)
│   ├── book_body/                    (Main book content - defines layout)
│   ├── front_matter/                 (Title and copyright pages)
│   ├── index_source/                 (Index source data)
│   └── reference/                    (Format/layout references)
├── outputs/                          (Task outputs, organized by task)
│   ├── 01_front_matter/              (Front matter task output)
│   │   ├── front_matter.docx         (Output file)
│   │   ├── front_matter.config.json  (Metadata for downstream)
│   │   └── temp/                     (Intermediate files)
│   ├── 02_index/                     (Index task output)
│   │   ├── index.docx                (Output file)
│   │   └── index.config.json         (Metadata for downstream)
│   ├── 03_full_book/                 (Full book task output)
│   │   ├── full_book.docx            (Final output)
│   │   └── full_book.config.json     (Summary metadata)
│   └── staging/                      (Intermediate for joining)
└── intermediate/                     (Extracted data, JSON exports)
```

## Config Files Location

| Config File                        | Location                     | Task | Purpose                                              |
| ---------------------------------- | ---------------------------- | ---- | ---------------------------------------------------- |
| `front_matter_builder.config.json` | `src/front_matter/builders/` | 1    | Configure front matter build inputs/outputs          |
| `index_builder.config.json`        | `src/index/build/`           | 2    | Configure index build (reads Task 1 metadata)        |
| `full_book_assembler.config.json`  | `src/final_book/`            | 3    | Configure final assembly (reads Task 1 & 2 metadata) |

## Config Structure

Each config JSON follows this pattern:

```json
{
  "task": "task_name",
  "description": "What this task does",
  "inputs": {
    "input_file": "data/inputs/.../file.docx",
    "metadata_from_previous": "data/outputs/XX_previous_task/metadata.json"
  },
  "outputs": {
    "output_dir": "data/outputs/XX_task_name",
    "filename": "output.docx",
    "metadata_filename": "metadata.json",
    "temp_dir": "data/outputs/XX_task_name/temp"
  },
  "options": {
    "setting_key": "value"
  }
}
```

## Metadata Flow

Each task produces a `.config.json` metadata file that downstream tasks read:

1. **Task 1** (Front Matter) outputs:
   - `front_matter.docx` (the document)
   - `front_matter.config.json` (metadata: page_count, last_page_numbering, etc.)

2. **Task 2** (Index) reads:
   - Task 1 metadata to determine starting page number

3. **Task 3** (Full Book) reads:
   - Task 1 metadata and Task 2 metadata for assembly info

## Execution Flow

```
Task 1: Front Matter Builder
  ├─ Reads: src/front_matter/builders/front_matter_builder.config.json
  ├─ Inputs: data/inputs/front_matter/*, data/inputs/book_body/*
  └─ Outputs: data/outputs/01_front_matter/{front_matter.docx, front_matter.config.json}
                ↓
Task 2: Index Builder (future)
  ├─ Reads: src/index/build/index_builder.config.json
  ├─ Reads metadata: data/outputs/01_front_matter/front_matter.config.json
  ├─ Inputs: data/intermediate/index.json, data/inputs/book_body/*
  └─ Outputs: data/outputs/02_index/{index.docx, index.config.json}
                ↓
Task 3: Full Book Assembler (future)
  ├─ Reads: src/final_book/full_book_assembler.config.json
  ├─ Reads metadata: Task 1 & Task 2 outputs
  ├─ Inputs: data/outputs/01_front_matter/front_matter.docx,
             data/inputs/book_body/*.docx, data/outputs/02_index/index.docx
  └─ Outputs: data/outputs/03_full_book/{full_book.docx, full_book.config.json}
```

## File Naming Convention

- **Inputs**: Original filenames (e.g., `title_page.docx`, `HITS_AND_HAPPINESS_BODY.docx`)
- **Outputs**: Task-specific names (e.g., `front_matter.docx`, `index.docx`, `full_book.docx`)
- **Metadata**: Always `{task_name}.config.json` in the same folder as output document

## Notes

- All paths in config files are **relative to project root** (or absolute)
- `auto_discover_missing_inputs` fallback searches by filename if configured path doesn't exist
- Temp files are cleaned up after use unless `delete_temp_*` is false
- Metadata is generated automatically after each successful task completion
