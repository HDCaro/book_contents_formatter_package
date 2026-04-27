#!/usr/bin/env python3
"""
Artist, Band, and Book Index Generator for Richard's Book
Extracts and alphabetizes all mentions of artists, bands, and books from a Word document.
"""

import re
import docx
from collections import defaultdict
import csv
from typing import Dict, List, Set, Tuple


class IndexGenerator:
    def __init__(self):
        # Known patterns for artists/bands
        self.artist_patterns = [
            # Quoted artists/bands
            r'"([A-Z][^"]*(?:and the|&)[^"]*)"',  # "Martha and the Vandellas"
            r'"([A-Z][A-Za-z\s&]+)"',  # General quoted names

            # Specific band patterns
            r'\b(The [A-Z][A-Za-z\s]+?)(?:\s|,|\.|\)|$)',  # The Beatles, The O'Jays
            r'\b([A-Z][A-Za-z]+\s+and\s+the\s+[A-Z][A-Za-z]+)',  # Martha and the Vandellas
            r'\b([A-Z][A-Za-z]+\s+&\s+[A-Z][A-Za-z]+)',  # Simon & Garfunkel

            # Solo artists (Last, First format or single names)
            r'\b([A-Z][a-z]+,\s+[A-Z][a-z]+)',  # Bassey, Shirley
            r'\b([A-Z][a-z]+\s+[A-Z][a-z]+)(?=\s+(?:told|said|wrote|sang|performed|recorded|arranged|produced|composed))',
            # Context clues

            # Music industry specific patterns
            r'\b([A-Z][A-Za-z]+\s+[A-Z][A-Za-z]+)(?=\s+(?:arranger|producer|composer|musician|singer|guitarist|drummer|bassist))',
            r'(?:arranger|producer|composer)\s+([A-Z][A-Za-z]+\s+[A-Z][A-Za-z]+)',
            r'(?:by|with|featuring)\s+([A-Z][A-Za-z]+\s+[A-Z][A-Za-z]+)',

            # Band name patterns
            r'\b([A-Z][A-Za-z]+\s+[A-Z][A-Za-z]+\s+(?:Band|Orchestra|Quartet|Quintet|Trio))',
            r'\b([A-Z][A-Za-z]+\s+(?:Brothers|Sisters))',
        ]

        # Book title patterns
        self.book_patterns = [
            # Italicized books (common in Word docs)
            r'\*([^*]+)\*',  # *Book Title*

            # Books with "book" context
            r'(?:book|novel|biography|autobiography)\s+([A-Z][^,.!?]+?)(?:\s+by|\s+was|\s+is|,|\.)',

            # Specific book title patterns from sample
            r'\b(Six Days of the Condor|Three Days of the Condor|Samson and Delilah)',

            # Quoted titles
            r'"([A-Z][^"]*(?:Days|Story|Life|Biography|History)[^"]*)"',
        ]

        # Known artists/bands from sample (to catch missed ones)
        self.known_entities = {
            'artists': {
                'Martha and the Vandellas', 'Marvin Gaye', 'William Stevenson', 'Ivy Hunter',
                'Paul Riser', 'Steve Rowland', 'Don Was', 'Aretha Franklin', 'Dusty Springfield',
                'George Benson', 'Arif Mardin', 'Bette Midler', 'Chaka Khan', 'Swing Out Sister',
                'Mitch Dalton', 'The O\'Jays', 'The Stylistics', 'The Spinners', 'Bee Gees',
                'Thom Bell', 'Cecil B. DeMille', 'Sidney Pollack', 'James Grady', 'Jesse',
                'The Beatles', 'Shirley Bassey', 'Elvis Presley', 'Frank Sinatra', 'Ray Charles',
                'Duke Ellington', 'Count Basie', 'Ella Fitzgerald', 'John Coltrane', 'Miles Davis',
                'Pat Metheny', 'Weather Report', 'Steely Dan', 'Joni Mitchell', 'Bob Dylan',
                'The Rolling Stones', 'Led Zeppelin', 'Pink Floyd', 'Queen', 'David Bowie',
                'Elton John', 'Paul McCartney', 'George Harrison', 'Ringo Starr', 'John Lennon'
            },
            'books': {
                'Six Days of the Condor', 'Three Days of the Condor', 'Samson and Delilah',
                'Judges 14', 'The Bible', 'Spy in the House of Love', 'Tune In', 'Here There and Everywhere',
                'The Beatles Anthology', 'Popular Music Studies', 'Chronicles', 'Life', 'Just Kids'
            }
        }

        self.artists_found = defaultdict(set)  # artist -> set of page numbers
        self.books_found = defaultdict(set)  # book -> set of page numbers

    def clean_name(self, name: str) -> str:
        """Clean and normalize names"""
        name = name.strip()
        # Remove extra whitespace
        name = re.sub(r'\s+', ' ', name)
        # Remove trailing punctuation
        name = re.sub(r'[,.!?;]+$', '', name)
        return name

    def alphabetize_name(self, name: str) -> str:
        """Convert names to alphabetizable format matching the provided index style"""
        name = self.clean_name(name)

        # Handle "The" prefix for bands - move to end
        if name.startswith('The ') and len(name) > 4:
            return f"{name[4:]}, The"

        # Handle "and the" in band names
        if ' and the ' in name.lower():
            parts = name.split(' and the ')
            if len(parts) == 2:
                return f"{parts[1]}, {parts[0]} and the"

        # Handle common name patterns from the index
        # For names like "John Smith", keep as is for sorting
        # For names already in "Last, First" format, keep as is
        if ',' in name:
            return name

        # For two-word names, check if it should be "Last, First"
        parts = name.split()
        if len(parts) == 2 and not name.startswith(('Dr', 'Mr', 'Mrs', 'Ms', 'Sir', 'Lord')):
            # Common first names that indicate "First Last" format
            first_names = {'John', 'Paul', 'George', 'Ringo', 'Bob', 'Ray', 'Duke', 'Count',
                           'Frank', 'Tony', 'Mike', 'Steve', 'Dave', 'Bill', 'Jim', 'Tom',
                           'Mary', 'Diana', 'Ella', 'Sarah', 'Kate', 'Anne', 'Jane'}

            if parts[0] in first_names:
                return f"{parts[1]}, {parts[0]}"

        return name

    def extract_from_text(self, text: str, page_num: int = 1):
        """Extract artists and books from text"""

        # Extract artists using patterns
        for pattern in self.artist_patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE)
            for match in matches:
                artist = self.clean_name(match.group(1))
                if len(artist) > 2 and not artist.lower() in ['the', 'and', 'of', 'in']:
                    self.artists_found[artist].add(page_num)

        # Extract books using patterns
        for pattern in self.book_patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE)
            for match in matches:
                book = self.clean_name(match.group(1))
                if len(book) > 3:
                    self.books_found[book].add(page_num)

        # Check for known entities
        for artist in self.known_entities['artists']:
            if artist.lower() in text.lower():
                self.artists_found[artist].add(page_num)

        for book in self.known_entities['books']:
            if book.lower() in text.lower():
                self.books_found[book].add(page_num)

    def process_docx(self, file_path: str):
        """Process Word document"""
        try:
            doc = docx.Document(file_path)

            for page_num, paragraph in enumerate(doc.paragraphs, 1):
                if paragraph.text.strip():
                    self.extract_from_text(paragraph.text, page_num)

        except Exception as e:
            print(f"Error processing Word document: {e}")
            return False

        return True

    def process_text_file(self, file_path: str):
        """Process plain text file"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()

            # Split into pages (approximate)
            pages = content.split('\n\n')  # Assume double newlines separate sections

            for page_num, page_text in enumerate(pages, 1):
                if page_text.strip():
                    self.extract_from_text(page_text, page_num)

        except Exception as e:
            print(f"Error processing text file: {e}")
            return False

        return True

    def generate_index(self, output_format='text', separate_books=True):
        """Generate the final index"""

        # Sort artists alphabetically
        sorted_artists = {}
        for artist in self.artists_found:
            alpha_key = self.alphabetize_name(artist)
            sorted_artists[alpha_key] = (artist, sorted(list(self.artists_found[artist])))

        # Sort books alphabetically
        sorted_books = {}
        for book in self.books_found:
            alpha_key = self.clean_name(book)
            sorted_books[alpha_key] = (book, sorted(list(self.books_found[book])))

        if output_format == 'text':
            return self._generate_text_index(sorted_artists, sorted_books, separate_books)
        elif output_format == 'csv':
            return self._generate_csv_index(sorted_artists, sorted_books, separate_books)
        elif output_format == 'word':
            return self._generate_word_index(sorted_artists, sorted_books, separate_books)

    def _generate_text_index(self, artists, books, separate_books):
        """Generate text format index matching the provided format"""
        output = []

        # Main INDEX header
        output.append("INDEX")
        output.append("")

        # Group entries by first letter
        all_entries = {}

        # Add artists
        for alpha_key in sorted(artists.keys()):
            original_name, pages = artists[alpha_key]
            page_list = ", ".join(map(str, pages))
            first_letter = alpha_key[0].upper()
            if first_letter not in all_entries:
                all_entries[first_letter] = []
            all_entries[first_letter].append(f"{alpha_key}, {page_list}")

        # Add books if not separate
        if not separate_books and books:
            for alpha_key in sorted(books.keys()):
                original_name, pages = books[alpha_key]
                page_list = ", ".join(map(str, pages))
                first_letter = alpha_key[0].upper()
                if first_letter not in all_entries:
                    all_entries[first_letter] = []
                all_entries[first_letter].append(f"{alpha_key}, {page_list}")

        # Output by letter sections
        for letter in sorted(all_entries.keys()):
            output.append(letter)
            output.append("")
            for entry in sorted(all_entries[letter]):
                output.append(entry)
            output.append("")

        # Separate books section if requested
        if separate_books and books:
            output.append("BOOKS AND PUBLICATIONS")
            output.append("")
            book_entries = {}
            for alpha_key in sorted(books.keys()):
                original_name, pages = books[alpha_key]
                page_list = ", ".join(map(str, pages))
                first_letter = alpha_key[0].upper()
                if first_letter not in book_entries:
                    book_entries[first_letter] = []
                book_entries[first_letter].append(f"{alpha_key}, {page_list}")

            for letter in sorted(book_entries.keys()):
                output.append(letter)
                output.append("")
                for entry in sorted(book_entries[letter]):
                    output.append(entry)
                output.append("")

        return "\n".join(output)

    def _generate_csv_index(self, artists, books, separate_books):
        """Generate CSV format for Excel import"""
        rows = []

        # Artists section
        rows.append(["Type", "Name", "Alphabetical Key", "Pages"])
        rows.append(["", "", "", ""])  # Separator

        for alpha_key in sorted(artists.keys()):
            original_name, pages = artists[alpha_key]
            page_list = ", ".join(map(str, pages))
            rows.append(["Artist/Band", original_name, alpha_key, page_list])

        if separate_books and books:
            rows.append(["", "", "", ""])  # Separator
            for alpha_key in sorted(books.keys()):
                original_name, pages = books[alpha_key]
                page_list = ", ".join(map(str, pages))
                rows.append(["Book", original_name, alpha_key, page_list])

        return rows


def main():
    """Main function to run the index generator"""
    import sys
    import os

    # Default values
    default_file = "HITS AND HAPPINESS FINAL 2 Format MOM Discog.docx"

    if len(sys.argv) < 2:
        # Use default file if no parameters provided
        file_path = default_file
        output_format = 'text'
        separate_books = False
        print(f"No parameters provided. Using defaults:")
        print(f"File: {file_path}")
        print(f"Format: {output_format}")
        print(f"Separate books: {separate_books}")
    else:
        print("Usage: python artist_book_index_generator.py [file_path] [output_format] [separate_books]")
        print("Output formats: text (default), csv")
        print("separate_books: true/false (default: false - combines artists and books)")
        print("\nExample:")
        print('python artist_book_index_generator.py "HITS AND HAPPINESS FINAL 2 Format MOM Discog.docx" text false')
        print(f"\nDefault (no parameters): Uses '{default_file}' with text format")
        print("=" * 70)

        file_path = sys.argv[1] if len(sys.argv) > 1 else default_file
        output_format = sys.argv[2] if len(sys.argv) > 2 else 'text'
        separate_books = sys.argv[3].lower() == 'true' if len(sys.argv) > 3 else False

    # Generate output filename based on input file
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    output_base = f"{base_name}_index"

    generator = IndexGenerator()

    print(f"Processing file: {file_path}")
    print(f"Output format: {output_format}")
    print(f"Separate books section: {separate_books}")
    print(f"Output base name: {output_base}")
    print("=" * 50)

    # Process file based on extension
    if file_path.lower().endswith('.docx'):
        success = generator.process_docx(file_path)
    else:
        success = generator.process_text_file(file_path)

    if not success:
        print("Failed to process file")
        return

    # Generate index
    if output_format == 'csv':
        rows = generator.generate_index('csv', separate_books)
        output_file = f'{output_base}.csv'
        with open(output_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerows(rows)
        print(f"CSV index saved to: {output_file}")
    else:
        index_text = generator.generate_index('text', separate_books)
        output_file = f'{output_base}.txt'
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(index_text)
        print(f"Text index saved to: {output_file}")
        print("\nPreview (first 1000 characters):")
        print("=" * 50)
        print(index_text[:1000] + "..." if len(index_text) > 1000 else index_text)

    # Print statistics
    print("\n" + "=" * 50)
    print("STATISTICS:")
    print(f"Artists/Bands found: {len(generator.artists_found)}")
    print(f"Books found: {len(generator.books_found)}")
    print(f"Total entries: {len(generator.artists_found) + len(generator.books_found)}")

    # Show some examples
    if generator.artists_found:
        print(f"\nSample artists found:")
        for i, (artist, pages) in enumerate(list(generator.artists_found.items())[:5]):
            print(f"  - {artist}: {sorted(list(pages))}")

    if generator.books_found:
        print(f"\nSample books found:")
        for i, (book, pages) in enumerate(list(generator.books_found.items())[:3]):
            print(f"  - {book}: {sorted(list(pages))}")


if __name__ == "__main__":
    main()