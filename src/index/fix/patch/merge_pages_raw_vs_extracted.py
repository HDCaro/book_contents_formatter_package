import json
from pathlib import Path
import re

# ---------------------------------------------------------------------------
# PATHS
# ---------------------------------------------------------------------------

BASE_PATH = Path(__file__).resolve().parents[4]

RAW_JSON = BASE_PATH / "data/index/intermediate/index_raw.json"
EXTRACTED_JSON = BASE_PATH / "data/index/intermediate/index_curated_extracted.json"

OUTPUT_JSON = BASE_PATH / "data/index/intermediate/index_raw_pages_merged.json"

# ---------------------------------------------------------------------------
# NORMALIZATION
# ---------------------------------------------------------------------------

def normalize(s):
    s = s.lower().strip()
    s = re.sub(r"^the\s+", "", s)   # remove leading "the"
    s = re.sub(r"[^\w\s]", "", s)   # remove punctuation
    return s

# ---------------------------------------------------------------------------
# LOAD
# ---------------------------------------------------------------------------

print("\n=== MERGE RAW vs EXTRACTED (ENHANCED) ===\n")

raw = json.load(open(RAW_JSON, encoding="utf-8"))
extracted = json.load(open(EXTRACTED_JSON, encoding="utf-8"))

# ---------------------------------------------------------------------------
# BUILD LOOKUP
# ---------------------------------------------------------------------------

extracted_lookup = {normalize(k): v for k, v in extracted.items()}

# ---------------------------------------------------------------------------
# MERGE
# ---------------------------------------------------------------------------

merged = {}
matched = 0
alias_matched = 0
missing = []

for key, entry in raw.items():

    norm_key = normalize(key)

    extracted_entry = extracted_lookup.get(norm_key)

    # --- direct match ---
    if extracted_entry:
        matched += 1

    # --- alias match ---
    else:
        extracted_entry = None

        for alias in entry.get("aliases", []):
            alias_norm = normalize(alias)

            if alias_norm in extracted_lookup:
                extracted_entry = extracted_lookup[alias_norm]
                alias_matched += 1
                break

    # --- apply result ---
    if extracted_entry:
        merged[key] = {
            **entry,
            "pages": extracted_entry.get("pages", [])
        }
    else:
        merged[key] = entry
        missing.append(key)

# ---------------------------------------------------------------------------
# EXTRA IN EXTRACTED
# ---------------------------------------------------------------------------

raw_norm_keys = set(normalize(k) for k in raw.keys())

extra = [
    k for k in extracted.keys()
    if normalize(k) not in raw_norm_keys
]

# ---------------------------------------------------------------------------
# SAVE
# ---------------------------------------------------------------------------

OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)

json.dump(merged, open(OUTPUT_JSON, "w", encoding="utf-8"), indent=2)

# ---------------------------------------------------------------------------
# REPORT
# ---------------------------------------------------------------------------

print("\n=== RESULT ===\n")

print(f"✔ Direct matches:      {matched}")
print(f"🔁 Alias matches:      {alias_matched}")
print(f"⚠ Still missing:      {len(missing)}")
print(f"➕ Extra extracted:    {len(extra)}\n")

print("Sample still missing:")
for e in missing[:10]:
    print(f" - {e}")

print("\nSample extra extracted:")
for e in extra[:10]:
    print(f" - {e}")

print("\n💾 Saved merged file\n")