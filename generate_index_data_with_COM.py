import re
import json
import time
from pathlib import Path
import win32com.client as win32

# ---------------- CONFIG ---------------- #

DOCX_INPUT = "HITS AND HAPPINESS FINAL 2 Format MOM Discog.docx"
OUTPUT_JSON = "index_raw.json"

MIN_OCCURRENCES = 1  # 🔥 KEEP SINGLE OCCURRENCES

VERBOSE = True

# ---------------- REGEX ---------------- #

PATTERN = re.compile(
    r"\b([A-Z][a-zA-Z]+(?:[-'][A-Za-z]+)*(?:\s+[A-Z][a-zA-Z]+(?:[-'][A-Za-z]+)*)+)\b"
)

# ---------------- NORMALIZATION ---------------- #

def normalize_caps(text):
    def fix_word(word):
        parts = word.split('-')
        return "-".join(p.capitalize() if p.isupper() else p for p in parts)
    return " ".join(fix_word(w) for w in text.split())

# ---------------- WORD COM ---------------- #

def open_word(path):
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc = word.Documents.Open(str(Path(path).resolve()))
    return word, doc

def close_word(word, doc):
    doc.Close(False)
    word.Quit()

# ---------------- INDEX BUILD ---------------- #

def build_index_fast(doc):
    index = {}

    rng = doc.Content
    text = rng.Text

    # normalize once
    text = normalize_caps(text)

    start = time.time()
    print("[INFO] Scanning document once...\n")

    total_matches = 0

    for match in PATTERN.finditer(text):
        candidate = match.group(1)

        start_pos = match.start()
        end_pos = match.end()

        try:
            word_range = doc.Range(Start=start_pos, End=end_pos)
            page = word_range.Information(3)  # wdActiveEndPageNumber
        except:
            continue

        if candidate not in index:
            index[candidate] = set()

        index[candidate].add(int(page))
        total_matches += 1

        # optional light progress
        if total_matches % 1000 == 0:
            print(f"[PROGRESS] {total_matches} matches processed...")

    print(f"\n[INFO] Scan complete in {time.time() - start:.2f}s")
    print(f"[INFO] Total matches found: {total_matches}")
    print(f"[INFO] Unique entries: {len(index)}\n")

    # ---------------- FINAL FILTER ---------------- #

    result = {}

    for k, pages in index.items():
        pages = sorted(pages)

        # 🔥 KEEP SINGLE ENTRIES BUT FILTER OBVIOUS NOISE
        words = k.split()

        keep = (
            len(words) <= 4 and
            all(w[0].isupper() for w in words)
        )

        if keep and len(pages) >= MIN_OCCURRENCES:
            result[k] = {
                "normalized": k,
                "type": "unknown",
                "pages": pages,
                "action": "keep"
            }

    print(f"[INFO] Final entries kept: {len(result)}\n")

    return result

# ---------------- MAIN ---------------- #

def main():
    print("\n=== FAST INDEX GENERATION (WORD COM) ===\n")

    word, doc = open_word(DOCX_INPUT)

    try:
        index = build_index_fast(doc)

        with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
            json.dump(index, f, indent=2)

        print(f"✅ Generated {OUTPUT_JSON}")

    finally:
        close_word(word, doc)

if __name__ == "__main__":
    main()