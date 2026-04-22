#!/bin/bash

echo "Book Contents Formatter - FIXED VERSION - Mac/Linux Setup"
echo "=========================================================="
echo

echo "Installing Python dependencies..."
pip install -r requirements.txt

echo
echo "Creating test example..."
python test_example.py

echo
echo "Running FIXED formatter with proper right-aligned page numbers..."
python format_contents_fixed.py

echo
echo "Creating alignment comparison demo..."
python alignment_comparison.py

echo
echo "Setup complete! Check the FIXED files:"
echo "- contents_author_format_FIXED.docx"
echo "- contents_traditional_format_FIXED.docx"
echo "- alignment_comparison_demo.docx"
echo "- advanced_book_indexer_fixed.py (for your Roman numeral index)"
echo