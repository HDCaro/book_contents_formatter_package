# Local Book Contents Formatter Setup

## Quick Setup

1. **Install Python dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

2. **For existing contents formatting** (like Richard Niles' book):
   ```bash
   python format_contents.py
   ```

3. **For automatic index generation from book body**:
   ```bash
   python advanced_book_indexer.py
   ```

## Files Included

### Core Scripts
- `format_contents.py` - Main contents formatter (Richard Niles' style)
- `update_contents.py` - Simple updater for changes
- `advanced_book_indexer.py` - Automatic index generator with Roman numerals

### Documentation
- `README_Contents_Formatter.md` - Detailed instructions
- `book_indexer_guide.md` - Guide for automatic indexer
- `setup_instructions.md` - This file

### Dependencies
- `requirements.txt` - Python packages needed

## Usage Examples

### Example 1: Format Existing Contents
If you have hand-made contents like Richard Niles:
```bash
python format_contents.py
```
Output: `contents_author_format.docx`

### Example 2: Generate Index from Book Body
If you have a complete book DOCX with heading styles:
```bash
# Edit advanced_book_indexer.py to point to your file
python advanced_book_indexer.py
```
Output: `complete_book_with_index.docx`

### Example 3: Update Existing Contents
If you need to make changes:
```bash
# Edit the contents_text in update_contents.py
python update_contents.py
```

## Customization

### Font and Styling
Edit these variables in the scripts:
```python
font_name = "Georgia"          # Change font family
chapter_num_size = 12          # Chapter number size
chapter_title_size = 20        # Chapter title size
alignment = "center"           # Text alignment
```

### Roman Numerals
The `advanced_book_indexer.py` automatically converts:
- 1 → i
- 2 → ii  
- 3 → iii
- 4 → iv
- 5 → v
- etc.

## Troubleshooting

### Common Issues
1. **"No module named 'docx'"**: Run `pip install python-docx`
2. **"No headings found"**: Make sure your DOCX uses Heading 1, Heading 2, etc. styles
3. **Wrong formatting**: Check the font and size variables in the script

### Getting Help
- Check the README files for detailed instructions
- Look at the example outputs to understand the format
- Modify the scripts to match your specific needs

## Next Steps
1. Test with a small sample file first
2. Adjust formatting parameters as needed
3. Run on your full book
4. Review and fine-tune the output