# Advanced Book Indexer Guide

## What This Does

This script automatically:
1. **Extracts headings** from your book body DOCX file
2. **Creates a formatted index** with Roman numerals (i, ii, iii, iv, v...)
3. **Places it at the front** of your book
4. **Uses your preferred formatting** (Georgia font, custom sizes, alignment)

## Requirements

Your book DOCX file must use **Word heading styles**:
- **Heading 1** for chapters
- **Heading 2** for sections  
- **Heading 3** for subsections

## How to Use

### Step 1: Prepare Your Book
Make sure your book body uses proper heading styles:
```
Heading 1: Chapter 1: Introduction
Heading 1: Chapter 2: Getting Started
Heading 2: Setting Up Your Environment
Heading 2: Basic Configuration
Heading 1: Chapter 3: Advanced Topics
```

### Step 2: Run the Script
```bash
python advanced_book_indexer.py
```

### Step 3: Customize (Optional)
Edit the script to change:
- Font name (default: Georgia)
- Font sizes (default: 12 for numbers, 20 for titles)
- Alignment (default: center)

## Output Example

The script creates an index like this:

```
                    TABLE OF CONTENTS

                        Chapter 1
                    Introduction
                    .................. i

                        Chapter 2
                    Getting Started
                    .................. ii

                Setting Up Your Environment
                    .................. iii

                Basic Configuration
                    .................. iv

                        Chapter 3
                    Advanced Topics
                    .................. v
```

## Integration with Workflow

You can combine this with the "Book Contents Formatter" workflow:

1. **First**: Use this script to extract headings and create basic structure
2. **Then**: Use the workflow to fine-tune formatting, fonts, and styling
3. **Result**: Professional book with perfect index placement

## Advanced Features

- **Roman numeral conversion**: Automatically converts 1→i, 2→ii, 3→iii, etc.
- **Hierarchical structure**: Indents sub-headings appropriately
- **Dot leaders**: Professional dots connecting titles to page numbers
- **Custom typography**: Full control over fonts and sizing
- **Page estimation**: Attempts to estimate actual page numbers

## Limitations

- Page numbers are estimated (for exact pages, you'd need the final formatted book)
- Requires proper heading styles in source document
- Works best with well-structured documents

## Next Steps

After running this script, you can:
1. **Review the generated index** for accuracy
2. **Adjust page numbers** if needed (after final formatting)
3. **Merge with your book body** to create the complete publication
4. **Use the workflow** for additional formatting refinements