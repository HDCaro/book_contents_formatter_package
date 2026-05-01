"""
===============================================================================
FILE: map_extracted_to_curated.py
LOCATION: src/index/fix/patch/
-------------------------------------------------------------------------------
Maps extracted index (from DOCX) back to curated keys using normalized names

FIXES:
- Handles missing "normalized" in extracted (fallback to key)
- Correct mapping direction: curated.normalized -> extracted.key
- Preserves structure for next steps (page fixing)
===============================================================================
"""

import json
from pathlib import Path

# ---------------------------------------------------------------------------
# PATHS
# ---------------------------------------------------------------------------

BASE_PATH = Path(__file__).resolve().parents[4]

INPUT_EXTRACTED = BASE_PATH / "data/index/intermediate/index_curated_extracted.json"
INPUT_CURATED = BASE_PATH / "data/index/intermediate/index_curated_old.json"

OUTPUT = BASE_PATH / "data/index/intermediate/index_curated_mapped.json"

# ---------------------------------------------------------------------------
# LOAD
# ---------------------------------------------------------------------------

print("\n=== MAP EXTRACTED TO CURATED ===\n")

print(f"📥 EXTRACTED: {INPUT_EXTRACTED}")
print(f"📥 CURATED:  {INPUT_CURATED}")
print(f"📤 OUTPUT:   {OUTPUT}\n")

if not INPUT_EXTRACTED.exists():
    print("❌ Missing extracted file")
    exit()

if not INPUT_CURATED.exists():
    print("❌ Missing curated file")
    exit()

with open(INPUT_EXTRACTED, "r", encoding="utf-8") as f:
    extracted = json.load(f)

with open(INPUT_CURATED, "r", encoding="utf-8") as f:
    curated = json.load(f)

# ---------------------------------------------------------------------------
# BUILD NORMALIZED LOOKUP FROM CURATED
# ---------------------------------------------------------------------------

norm_to_curated_key = {}

for key, entry in curated.items():
    norm = entry.get("normalized")
    if norm:
        norm_to_curated_key[norm.strip().lower()] = key

print(f"🔎 Loaded {len(norm_to_curated_key)} normalized mappings from curated\n")

# ---------------------------------------------------------------------------
# MAP
# ---------------------------------------------------------------------------

mapped = {}
unmatched = []

for key, entry in extracted.items():

    # ✅ CRITICAL FIX: fallback to key if normalized missing
    lookup_norm = (entry.get("normalized") or key).strip().lower()

    curated_key = norm_to_curated_key.get(lookup_norm)

    if not curated_key:
        unmatched.append(key)
        curated_key = key  # fallback (keeps entry alive)

    mapped[curated_key] = {
        "normalized": entry.get("normalized") or key,
        "aliases": entry.get("aliases", []),
        "aliases_external": entry.get("aliases_external", []),
        "pages": entry.get("pages", [])
    }

# ---------------------------------------------------------------------------
# SAVE
# ---------------------------------------------------------------------------

OUTPUT.parent.mkdir(parents=True, exist_ok=True)

with open(OUTPUT, "w", encoding="utf-8") as f:
    json.dump(mapped, f, indent=2, ensure_ascii=False)

# ---------------------------------------------------------------------------
# REPORT
# ---------------------------------------------------------------------------

print("\n=== RESULT ===\n")
print(f"✅ Mapped entries: {len(mapped)}")
print(f"⚠️ Unmatched entries: {len(unmatched)}\n")

if unmatched:
    print("Sample unmatched:")
    for e in unmatched[:20]:
        print(f"  - {e}")

print("\n🎉 Mapping complete\n")