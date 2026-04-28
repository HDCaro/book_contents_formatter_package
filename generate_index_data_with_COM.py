import re
import json
import time
from pathlib import Path
import requests
import win32com.client as win32

# ---------------- CONFIG ---------------- #

DOCX_INPUT = "HITS AND HAPPINESS FINAL 2 Format MOM Discog.docx"
OUTPUT_JSON = "index_raw.json"
CURATED_JSON = "index_curated.json"
CACHE_FILE = "discogs_cache.json"

MIN_OCCURRENCES = 1
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

# ---------------- DISCOGS CACHE ---------------- #

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
    key = query.lower()

    if key in discogs_cache:
        return discogs_cache[key]

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
        dtype = None

        if results:
            found = True
            dtype = results[0].get("type")

        discogs_cache[key] = (found, dtype)
        return found, dtype

    except:
        discogs_cache[key] = (False, None)
        return False, None

# ---------------- CLASSIFICATION ---------------- #

def classify_entity(name):
    parts = name.split()

    # person
    if len(parts) == 2 and parts[0] != "The":
        return f"{parts[1]}, {parts[0]}", "person"

    # band
    if parts[0] == "The":
        return f"{' '.join(parts[1:])}, The", "band"

    # discogs
    found, dtype = discogs_search(name)

    if found:
        if dtype == "artist":
            return name, "band"
        if dtype in ("release", "master"):
            return name, "album"

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

    print("[INFO] Scanning document once...\n")
    start = time.time()

    count = 0

    for match in PATTERN.finditer(text):
        name = match.group(1)

        try:
            rng = doc.Range(Start=match.start(), End=match.end())
            page = int(rng.Information(3))
        except:
            continue

        index.setdefault(name, set()).add(page)

        count += 1
        if count % 1000 == 0:
            print(f"[PROGRESS] {count} matches...")

    print(f"\n[INFO] Scan complete in {time.time() - start:.2f}s")
    print(f"[INFO] Unique entries: {len(index)}\n")

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

# ---------------- MERGE ---------------- #

def merge_curated(raw_data):
    path = Path(CURATED_JSON)

    if not path.exists():
        print("🆕 Creating curated file...\n")

        curated = {
            k: {**v, "action": "keep"}
            for k, v in raw_data.items()
        }

        with open(path, "w", encoding="utf-8") as f:
            json.dump(curated, f, indent=2)

        print("✅ Created curated file")
        return

    with open(path, "r", encoding="utf-8") as f:
        curated = json.load(f)

    raw_keys = set(raw_data.keys())
    cur_keys = set(curated.keys())

    added = raw_keys - cur_keys
    missing = cur_keys - raw_keys
    common = raw_keys & cur_keys

    print("\n=== MERGE REPORT ===\n")

    # ADD
    print("➕ ADD:")
    for k in sorted(added):
        print(f"  + {k}")
        curated[k] = {**raw_data[k], "action": "keep"}
    if not added:
        print("  none")

    # PROCESS
    print("\n🔄 PROCESS:")

    for k in sorted(common):
        action = curated[k].get("action", "keep")

        raw_pages = set(raw_data[k]["pages"])
        cur_pages = set(curated[k]["pages"])

        if action == "remove":
            print(f"  [REMOVE] {k}")
            continue

        if action == "update":
            print(f"  [UPDATE] {k}")
            print(f"     old: {sorted(cur_pages)}")
            print(f"     new: {sorted(raw_pages)}")
            curated[k]["pages"] = sorted(raw_pages)
            continue

        if raw_pages != cur_pages:
            print(f"  [KEEP*] {k}")
        else:
            print(f"  [KEEP] {k}")

    # MISSING
    print("\n➖ MISSING:")
    for k in sorted(missing):
        print(f"  - {k}")
    if not missing:
        print("  none")

    print("\n=== END MERGE ===\n")

    with open(path, "w", encoding="utf-8") as f:
        json.dump(curated, f, indent=2)

    print("✅ Curated updated")

# ---------------- MAIN ---------------- #

def main():
    print("\n=== GENERATE INDEX DATA (COM) ===\n")

    word, doc = open_word(DOCX_INPUT)

    try:
        raw = build_index_fast(doc)

        with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
            json.dump(raw, f, indent=2)

        print("✅ Raw index generated")

        merge_curated(raw)

    finally:
        close_word(word, doc)

if __name__ == "__main__":
    main()