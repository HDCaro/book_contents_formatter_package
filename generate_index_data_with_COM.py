import re
import json
import time
from pathlib import Path
import requests
import win32com.client as win32

# ---------------- CONFIG ---------------- #

DOCX_INPUT = "HITS AND HAPPINESS FINAL 2 Format MOM Discog.docx"
OUTPUT_JSON = "index_raw.json"
CACHE_FILE = "discogs_cache.json"

MIN_OCCURRENCES = 1

# Optional: add your Discogs token if you have one
DISCOGS_TOKEN = ""

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

# ---------------- DISCOGS ---------------- #

def load_cache():
    try:
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return {}

def save_cache(cache):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, indent=2)

discogs_cache = load_cache()

def discogs_search(query):
    query = query.lower()

    if query in discogs_cache:
        return discogs_cache[query]

    headers = {"User-Agent": "IndexBuilder/1.0"}
    if DISCOGS_TOKEN:
        headers["Authorization"] = f"Discogs token={DISCOGS_TOKEN}"

    try:
        r = requests.get(
            "https://api.discogs.com/database/search",
            params={"q": query},
            headers=headers,
            timeout=5
        )
        results = r.json().get("results", [])

        found = False
        type_detected = None

        if results:
            found = True
            first = results[0]
            type_detected = first.get("type")  # artist, release, master

        discogs_cache[query] = (found, type_detected)
        return found, type_detected

    except:
        discogs_cache[query] = (False, None)
        return False, None

# ---------------- CLASSIFICATION ---------------- #

def classify_entity(name):
    parts = name.split()

    # PERSON
    if len(parts) == 2 and parts[0] != "The":
        return f"{parts[1]}, {parts[0]}", "person"

    # BAND
    if parts[0] == "The":
        return f"{' '.join(parts[1:])}, The", "band"

    # DISCOGS lookup (only for non-persons)
    found, dtype = discogs_search(name)

    if found:
        if dtype == "artist":
            return name, "band"
        if dtype in ("release", "master"):
            # could be album or song
            return name, "album"

    # fallback
    return name, "work"

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

    text = normalize_caps(doc.Content.Text)

    start = time.time()
    print("[INFO] Scanning document once...\n")

    total_matches = 0

    for match in PATTERN.finditer(text):
        candidate = match.group(1)

        try:
            word_range = doc.Range(Start=match.start(), End=match.end())
            page = word_range.Information(3)
        except:
            continue

        if candidate not in index:
            index[candidate] = set()

        index[candidate].add(int(page))
        total_matches += 1

        if total_matches % 1000 == 0:
            print(f"[PROGRESS] {total_matches} matches...")

    print(f"\n[INFO] Scan complete in {time.time() - start:.2f}s")
    print(f"[INFO] Unique entries: {len(index)}\n")

    # ---------------- FINAL BUILD ---------------- #

    result = {}

    for name, pages in index.items():
        pages = sorted(pages)

        words = name.split()

        keep = (
            len(words) <= 4 and
            all(w[0].isupper() for w in words)
        )

        if not keep or len(pages) < MIN_OCCURRENCES:
            continue

        normalized, typ = classify_entity(name)

        result[name] = {
            "normalized": normalized,
            "type": typ,
            "pages": pages,
            "action": "keep"
        }

    print(f"[INFO] Final entries: {len(result)}\n")

    save_cache(discogs_cache)

    return result

# ---------------- MAIN ---------------- #

def main():
    print("\n=== FAST INDEX + TYPES + DISCOGS ===\n")

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