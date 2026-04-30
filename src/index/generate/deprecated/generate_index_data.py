import re
import json
import sys
import unicodedata
from pathlib import Path

import requests
from docx import Document
from docx.opc.exceptions import PackageNotFoundError

# ---------------- CONFIG ---------------- #

DOCX_INPUT = "HITS AND HAPPINESS FINAL 2 Format MOM Discog.docx"
OUTPUT_JSON = "index_raw.json"
CURATED_JSON = "index_curated_old.json"
CACHE_FILE = "discogs_cache.json"

WORDS_PER_PAGE = 600
MIN_OCCURRENCES = 2

DISCOGS_TOKEN = ""
VERBOSE = True
DEBUG_CHARS = True
DEBUG_PAGES = True

# ---------------- LOGGING ---------------- #

def log(msg):
    if VERBOSE:
        print(msg)

# ---------------- DEBUG ---------------- #

def debug_strange_chars(text):
    seen = set()
    for c in text:
        if ord(c) > 127:
            if c not in seen:
                seen.add(c)
                print(f"[UNICODE] '{c}' (code {ord(c)})")

def debug_page_content(i, page):
    if DEBUG_PAGES and ("Last" in page or "Silver" in page):
        print(f"\n--- PAGE {i} (RAW) ---")
        print(page[:500])
        print("--- END PAGE ---\n")

def debug_candidates(i, candidates):
    if DEBUG_PAGES:
        print(f"[PAGE {i}] Candidates found: {candidates}")

# ---------------- TEXT NORMALIZATION ---------------- #

def normalize_text(text):
    text = unicodedata.normalize("NFKC", text)

    replacements = {
        "\u00A0": " ",
        "\u2007": " ",
        "\u202F": " ",
        "\u2013": "-",
        "\u2014": "-",
        "\u2018": "'",
        "\u2019": "'",
        "\u201C": '"',
        "\u201D": '"',
        "\u2026": "...",
    }

    for k, v in replacements.items():
        text = text.replace(k, v)

    text = re.sub(r"[^\w\s\-']", " ", text)
    text = re.sub(r"\s+", " ", text).strip()

    return text

# ---------------- CAPS NORMALIZATION ---------------- #

def normalize_caps(text):
    def fix_word(word):
        parts = word.split('-')
        fixed = []
        for p in parts:
            if p.isupper() and len(p) > 1:
                fixed.append(p.capitalize())
            else:
                fixed.append(p)
        return "-".join(fixed)

    return " ".join(fix_word(w) for w in text.split())

# ---------------- CACHE ---------------- #

def load_cache():
    try:
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return {}

def save_cache(cache):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, indent=2)

cache = load_cache()

# ---------------- TEXT ---------------- #

def extract_text(docx):
    try:
        doc = Document(docx)
        text = "\n".join(p.text for p in doc.paragraphs)

        if DEBUG_CHARS:
            print("\n=== SCANNING FOR STRANGE CHARACTERS ===")
            debug_strange_chars(text)
            print("=== END SCAN ===\n")

        return text

    except PackageNotFoundError:
        print(f"❌ File not found: {docx}")
        sys.exit(1)

def split_pages(text):
    words = text.split()
    pages = [
        " ".join(words[i:i + WORDS_PER_PAGE])
        for i in range(0, len(words), WORDS_PER_PAGE)
    ]
    log(f"[INFO] Estimated pages: {len(pages)}")
    return pages

# ---------------- NORMALIZATION ---------------- #

def normalize_title(text):
    if "," in text:
        parts = [p.strip() for p in text.split(",")]
        parts.reverse()
        return " ".join(parts)
    return text

def clean_query(text):
    return re.sub(r"[^\w\s]", "", text)

# ---------------- DISCOGS ---------------- #

def discogs_search(query):
    query = clean_query(normalize_title(query))

    if query in cache:
        return cache[query]

    headers = {"User-Agent": "IndexBuilder/1.0"}

    try:
        r = requests.get(
            "https://api.discogs.com/database/search",
            params={"q": query},
            headers=headers,
            timeout=5
        )
        results = r.json().get("results", [])
        found = bool(results)
        cache[query] = found
        return found
    except Exception as e:
        log(f"[ERROR] Discogs request failed: {e}")
        cache[query] = False
        return False

# ---------------- DETECTION ---------------- #

def extract_candidates(text):
    pattern = r"\b([A-Z][a-zA-Z]+(?:[-'][A-Za-z]+)*(?:\s+[A-Z][a-zA-Z]+(?:[-'][A-Za-z]+)*)+)\b"
    return re.findall(pattern, text)

# ---------------- CLASSIFICATION ---------------- #

BAND_KEYWORDS = {
    "Band", "Boys", "Girls", "Group", "Orchestra",
    "Ensemble", "Quartet", "Trio", "Duo"
}

def is_likely_band(parts):
    if any(p in BAND_KEYWORDS for p in parts):
        return True
    if parts[0] == "The":
        return True
    if len(parts) >= 3:
        return True
    return False

def is_likely_person(parts):
    if len(parts) in (2, 3):
        if parts[0] != "The":
            return True
    return False

def invert_person(parts):
    return f"{parts[-1]}, {' '.join(parts[:-1])}"

def normalize_band(parts):
    if parts[0] == "The":
        return f"{' '.join(parts[1:])}, The"
    return " ".join(parts)

def classify_entity(candidate):
    parts = candidate.split()
    normalized = normalize_title(candidate)

    if discogs_search(normalized):
        return normalized, "work"

    if is_likely_band(parts):
        return normalize_band(parts), "band"

    if is_likely_person(parts):
        return invert_person(parts), "person"

    return candidate, "unknown"

# ---------------- INDEX ---------------- #

def build_index(pages):
    index = {}

    for i, page in enumerate(pages, 1):

        debug_page_content(i, page)

        # 🔥 FULL NORMALIZATION PIPELINE
        clean_page = normalize_text(page)
        clean_page = normalize_caps(clean_page)

        candidates = extract_candidates(clean_page)

        debug_candidates(i, candidates)

        for c in candidates:
            c = normalize_caps(c)  # 🔥 ensure merge consistency

            if len(c.split()) < 2:
                continue

            normalized, typ = classify_entity(c)

            if c not in index:
                index[c] = {
                    "normalized": normalized,
                    "type": typ,
                    "pages": set()
                }

            index[c]["pages"].add(i)

    return {
        k: {
            "normalized": v["normalized"],
            "type": v["type"],
            "pages": sorted(v["pages"]),
            "action": "keep"
        }
        for k, v in index.items()
        if len(v["pages"]) >= MIN_OCCURRENCES
    }

# ---------------- DIFF / INIT / MAIN ---------------- #
# (UNCHANGED BELOW)

def compare_indexes(raw, curated):
    raw_keys = set(raw.keys())
    cur_keys = set(curated.keys())
    added = raw_keys - cur_keys
    removed = cur_keys - raw_keys
    common = raw_keys & cur_keys
    changed = []
    for k in common:
        r = raw[k]
        c = curated[k]
        diffs = []
        if r.get("normalized") != c.get("normalized"):
            diffs.append(f"normalized: '{r.get('normalized')}' → '{c.get('normalized')}'")
        if r.get("type") != c.get("type"):
            diffs.append(f"type: '{r.get('type')}' → '{c.get('type')}'")
        if r.get("action") != c.get("action"):
            diffs.append(f"action: '{r.get('action')}' → '{c.get('action')}'")
        raw_pages = set(r.get("pages", []))
        cur_pages = set(c.get("pages", []))
        if raw_pages != cur_pages:
            added_pages = sorted(cur_pages - raw_pages)
            removed_pages = sorted(raw_pages - cur_pages)
            parts = []
            if added_pages:
                parts.append(f"+{added_pages}")
            if removed_pages:
                parts.append(f"-{removed_pages}")
            diffs.append(f"pages: {' '.join(parts)}")
        if diffs:
            changed.append((k, diffs))
    return added, removed, changed

def print_diff_report(raw, curated):
    added, removed, changed = compare_indexes(raw, curated)
    print("\n=== DIFF REPORT (RAW vs CURATED) ===\n")
    if added:
        print("➕ New entries:")
        for k in sorted(added):
            print(f"   + {k}")
    if removed:
        print("\n➖ Removed entries:")
        for k in sorted(removed):
            print(f"   - {k}")
    if changed:
        print("\n🔄 Changed entries:")
        for k, fields in changed:
            print(f"\n   * {k}")
            for f in fields:
                print(f"      - {f}")
    print("\n=== END DIFF ===\n")

def initialize_curated_file(raw_data):
    curated_path = Path(CURATED_JSON)
    if curated_path.exists():
        with open(CURATED_JSON, "r", encoding="utf-8") as f:
            curated = json.load(f)
        print_diff_report(raw_data, curated)
        return
    with open(CURATED_JSON, "w", encoding="utf-8") as f:
        json.dump(raw_data, f, indent=2)

def main():
    print("\n=== GENERATE INDEX DATA ===\n")
    text = extract_text(DOCX_INPUT)
    pages = split_pages(text)
    index = build_index(pages)
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(index, f, indent=2)
    save_cache(cache)
    print(f"\n✅ Generated {OUTPUT_JSON}")
    initialize_curated_file(index)

if __name__ == "__main__":
    main()