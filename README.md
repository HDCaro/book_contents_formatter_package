# Word Document TOC Generator with Roman Pagination
 
A Python script that automatically extracts headings from a Word document and generates a professional Table of Contents with proper Roman numeral pagination for book front matter.
 
## 📚 Overview
 
This tool solves the complex problem of creating professional book front matter by:
- Extracting chapter titles from a Word document using heading styles
- Generating a properly formatted Table of Contents
- Assembling title page, copyright page, and TOC with correct Roman pagination
- Handling multi-line chapter titles and ensuring single-line TOC entries
 
## 🎯 Features
 
- **Automatic Heading Extraction**: Scans Word documents for Heading 1 (chapters) and Heading 2 (chapter titles)
- **Smart Title Consolidation**: Combines multiple consecutive Heading 2 paragraphs into single TOC entries
- **Roman Numeral Pagination**: Proper front matter numbering (i, ii, iii, iv...)
- **Professional TOC Formatting**: Hanging indents, dotted leaders, right-aligned page numbers
- **Single-Line Guarantee**: Removes soft line breaks and ensures all titles fit on one line
- **Error Recovery**: Multiple fallback methods for pagination and formatting
 
## 📋 Requirements
 
### System Requirements
- Windows (required for Word COM automation)
- Microsoft Word installed
- Python 3.7+
 
### Python Dependencies
```bash
pip install python-docx pywin32
```
 
### Document Structure Requirements
 
Your main book document must follow this **exact** heading structure:
 
#### ✅ Correct Structure:
```
Heading 1: Chapter 1
Heading 2: The Beginning of Everything
[chapter content...]
 
Heading 1: Chapter 2  
Heading 2: A New Dawn
[chapter content...]
 
Heading 1: INTRODUCTION
[introduction content...]
 
Heading 1: RICHARD NILES DISCOGRAPHY BY YEAR
[discography content...]
```
 
#### ❌ Incorrect Structure:
```
Chapter 1: The Beginning (all in Heading 1)
Chapter 2: A New Dawn (all in Heading 1)
```
 
### Required Files
1. **Main book document** (e.g., `book.docx`) - with proper heading structure
2. **Title page** (e.g., `title.docx`) - standalone title page
3. **Copyright page** (e.g., `copyright.docx`) - copyright information
 
## 🚀 Usage
 
### Basic Usage
```python
python extract_with_word_com_fixed.py
```
 
### Customizing File Names
Edit the file paths in the script:
```python
book_file = "your-book-file.docx"
title_file = "your-title.docx"
copyright_file = "your-copyright.docx"
output_file = "final-book-with-toc.docx"
```
 
## 🔧 Technical Challenges & Solutions
 
### 1. Roman Numeral Pagination Challenge
**Problem**: Word's COM interface doesn't expose Roman numeral formatting in the obvious way.
 
**Solution**: 
```python
# ❌ This doesn't work:
sec2.PageSetup.PageNumberStyle = wdPageNumberStyleLowercaseRoman
 
# ✅ This works:
page_nums.NumberStyle = wdPageNumberStyleLowercaseRoman
page_nums.StartingNumber = 1
```
 
**Key Insight**: The `NumberStyle` property belongs to the `PageNumbers` collection, not the `PageSetup` object.
 
### 2. Page Numbering Restart Challenge
**Problem**: Getting the copyright page to start at Roman numeral "i" instead of continuing from the title page.
 
**Solution**: 
```python
# Critical sequence:
sec2.PageSetup.RestartPageNumbering = True
sec2.PageSetup.PageNumberStart = 1
sec2.Headers.Item(1).LinkToPrevious = False
sec2.Footers.Item(1).LinkToPrevious = False
```
 
**Key Insight**: Must unlink sections AND set restart properties before adding page numbers.
 
### 3. Multi-Line Chapter Title Challenge
**Problem**: Chapter titles split across multiple Heading 2 paragraphs appeared as separate TOC entries.
 
**Example Problem**:
```
Chapter 7: London 1975 –	73
You CAN Go Home Again!	73
```
 
**Solution**: Consecutive Heading 2 collection:
```python
while j <= total:
    if "heading 2" in p_style:
        chapter_title_parts.append(cleaned_text)
        j += 1
        continue  # Keep collecting
    else:
        break  # Stop at non-heading
 
complete_title = " ".join(chapter_title_parts)
```
 
**Result**:
```
Chapter 7: London 1975 – You CAN Go Home Again!	73
```
 
### 4. Soft Line Break Challenge
**Problem**: Heading 2 paragraphs contained soft line breaks (Shift+Enter) causing TOC entries to wrap.
 
**Solution**: Aggressive text cleaning:
```python
def clean_title_text(text):
    # Remove ALL types of line breaks
    cleaned = text.replace("
", " ").replace("", " ").replace("
", " ")
    # Normalize whitespace
    cleaned = re.sub(r'\s+', ' ', cleaned)
    return cleaned.strip()
```
 
### 5. COM Object Access Challenge
**Problem**: Inconsistent COM object access causing "unknown object" errors.
 
**Solution**: Use `.Item()` method consistently:
```python
# ❌ Unreliable:
sections(1)
 
# ✅ Reliable:
sections.Item(1)
```
 
## 📖 Output Structure
 
The final document will have this structure:
 
```
┌─────────────────────┐
│ Title Page          │ ← No page number
├─────────────────────┤
│ Copyright Page      │ ← Roman "i"
├─────────────────────┤
│ Table of Contents   │ ← Roman "ii", "iii", etc.
│ Page 1              │
│ Page 2 (if needed)  │
└─────────────────────┘
```
 
## 🎨 TOC Formatting Features
 
- **Hanging Indent**: 0.5" left indent with -0.5" first line
- **Dotted Leaders**: Professional dot pattern leading to page numbers
- **Right-Aligned Page Numbers**: Bold page numbers at 6" tab stop
- **Single-Line Entries**: All chapter titles guaranteed to fit on one line
- **Georgia Font**: 12pt for entries, 20pt bold for "CONTENTS" title
 
## 🐛 Troubleshooting
 
### Common Issues
 
1. **"Property RestartPageNumbering cannot be set"**
   - **Cause**: COM interface limitations
   - **Solution**: Script includes multiple fallback methods
 
2. **Chapter titles not detected**
   - **Cause**: Incorrect heading structure
   - **Solution**: Ensure Heading 1 for chapters, Heading 2 for titles
 
3. **Roman numerals showing as regular numbers**
   - **Cause**: NumberStyle not set correctly
   - **Solution**: Fixed in latest version using PageNumbers collection
 
4. **Multi-line TOC entries**
   - **Cause**: Soft line breaks in original headings
   - **Solution**: Aggressive text cleaning removes all line breaks
 
### Debug Output
 
The script provides detailed logging:
```
📖 Extracting headings with consecutive Heading 2 collection
   📍 Found Heading 1: 'Chapter 32'
   🔍 Processing chapter: 'Chapter 32'
   ✅ Collected Heading 2 part: 'California, There I Came'
   🎉 Complete chapter title: 'Chapter 32: California, There I Came'
   📌 Added to TOC: 'Chapter 32: California, There I Came' → page 393
```
 
## 📝 Example Output
 
```
CONTENTS
 
INTRODUCTION .................................................. 1
Chapter 1: 'Pre' His Story .................................... 5
Chapter 2: What is an "Arranger"? ............................ 17
Chapter 3: Kid Stuff ......................................... 27
Chapter 32: California, There I Came ......................... 393
RICHARD NILES DISCOGRAPHY BY YEAR ............................ 425
BOOKS BY RICHARD NILES ....................................... 430
```
 
## 🔄 Workflow
 
1. **Extraction Phase**: Scans main document for headings
2. **Consolidation Phase**: Combines multi-part chapter titles
3. **TOC Generation Phase**: Creates formatted table of contents
4. **Assembly Phase**: Combines title, copyright, and TOC
5. **Pagination Phase**: Applies Roman numerals with restart
6. **Finalization Phase**: Updates fields and saves final document
 
## 📚 Related Files
 
- `extract_with_word_com_fixed.py` - Main extraction and assembly script
- `index.py` - Alternative indexing approach (if available)
 
## 🤝 Contributing
 
This script was developed to solve specific book formatting challenges. Feel free to adapt it for your own document processing needs.
 
## ⚠️ Limitations
 
- **Windows Only**: Requires Word COM interface
- **Word Required**: Microsoft Word must be installed
- **Heading Structure**: Strict requirements for heading styles
- **File Paths**: Uses absolute paths, ensure files exist
 
## 🎉 Success Criteria
 
When working correctly, you should see:
- ✅ Title page with no page number
- ✅ Copyright page with Roman numeral "i"
- ✅ TOC pages with Roman numerals "ii", "iii", etc.
- ✅ All chapter titles on single lines
- ✅ Professional dotted leaders and formatting
- ✅ Proper hanging indents and alignment
 
---
 
*This tool was developed to automate the tedious process of creating professional book front matter while handling the complex edge cases that arise in real-world document processing.*
 

