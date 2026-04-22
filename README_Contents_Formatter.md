# Table of Contents Formatter for "Hits & Happiness – A Musical Memoir"

## What This Does

This Python script automatically formats your hand-made table of contents into a professional Word document using your preferred styling:

- **Georgia 12**: Chapter numbers (centered)
- **Georgia 20 bold**: Chapter titles (centered)  
- **Dot leaders**: Professional dots leading to page numbers
- **Proper spacing**: Clean layout between entries

## Files Created

1. **`contents_author_format.docx`** - Your preferred format with chapter numbers above titles
2. **`contents_traditional_format.docx`** - Alternative traditional format
3. **`format_contents.py`** - Main script with all your contents data
4. **`update_contents.py`** - Simple updater for future changes

## How to Use

### First Time Setup
1. Make sure you have Python installed on your computer
2. Install the required library by running: `pip install python-docx`
3. Run the script: `python format_contents.py`

### Making Updates Later
If you need to change chapter titles or page numbers:

1. **Option A**: Edit the data directly in `format_contents.py` (look for the `contents_data` list)
2. **Option B**: Use the simple updater script `update_contents.py`

### For Quick Updates
Copy your contents in this format:
```
Introduction...............................................................................................................1
Chapter 1: 'Pre' His Story.........................................................................................5
Chapter 2: What is an "Arranger"?.........................................................................17
```

Then paste it into `update_contents.py` and run it.

## Current Contents Included

✅ All 35 chapters from your original document  
✅ Introduction  
✅ Correct page numbers  
✅ Proper formatting with Georgia font  
✅ Centered alignment as requested  

## Features

- **Automatic formatting**: No manual spacing or font changes needed
- **Consistent styling**: Every entry follows the same professional format
- **Easy updates**: Change the data and regenerate instantly
- **Multiple formats**: Choose between your preferred style or traditional
- **Professional quality**: Ready for publication

## Technical Notes

- Uses Microsoft Word format (.docx)
- Georgia font family throughout
- Proper margins and spacing
- Compatible with all modern Word versions
- Can be easily converted to PDF if needed

## Support

If you need to modify the formatting or add new features, the scripts are well-commented and easy to customize. The main formatting logic is in the `create_formatted_contents()` function.