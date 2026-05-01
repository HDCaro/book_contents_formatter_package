"""
===============================================================================
FILE: rebuild_curated_from_extracted.py
LOCATION: src/index/fix/patch/
-------------------------------------------------------------------------------
Rebuild curated structure using extracted normalized entries

✔ Maps extracted normalized names → curated entries
✔ Restores original keys (editorial names)
✔ Injects corrected pages
✔ Preserves aliases, types, actions

OUTPUT
------
index_curated_rebuilt.json
===============================================================================
"""

import json
from pathlib import Path

# ---------------- ROOT ---------------- #

def find_project_root():
    current = Path(__file__).resolve()
    for parent in [current] + list(current.parents):
        if (parent / "src").exists() and (parent / "data").exists():
            return parent
    raise RuntimeError("Project root not found")

BASE_DIR = find_project_root()

# ---------------- PATHS ---------------- #

EXTRACTED_JSON = BASE_DIR / "data/index/intermediate/index_curated_fixed_pages.json"
CURATED_JSON = BASE_DIR / "data/index/intermediate/index_curated_old.json"

OUTPUT_JSON = BASE_DIR / "data/index/intermediate/index_curated_rebuilt.json"

# ---------------- MAIN ---------------- #

def main():
    print("\n=== REBUILD CURATED FROM EXTRACTED ===\n")

    print(f"📥 Extracted: {EXTRACTED_JSON}")
    print(f"📥 Curated:   {CURATED_JSON}")
    print(f"📤 Output:    {OUTPUT_JSON}\n")

    if not EXTRACTED_JSON.exists():
        print("❌ Missing extracted JSON")
        return

    if not CURATED_JSON.exists():
        print("❌ Missing curated JSON")
        return

    extracted = json.load(open(EXTRACTED_JSON, encoding="utf-8"))
    curated = json.load(open(CURATED_JSON, encoding="utf-8"))

    # Build normalized → pages map
    normalized_map = {}

    for name, entry in extracted.items():
        normalized_map[name.strip().lower()] = entry["pages"]

    rebuilt = {}

    missing_matches = []
    matched = 0

    for key, entry in curated.items():

        normalized = entry.get("normalized", "").strip().lower()

        pages = normalized_map.get(normalized)

        if pages:
            matched += 1
        else:
            pages = entry.get("pages", [])
            missing_matches.append(key)

        rebuilt[key] = {
            "normalized": entry.get("normalized"),
            "type": entry.get("type"),
            "aliases": entry.get("aliases", []),
            "aliases_external": entry.get("aliases_external", []),
            "pages": pages,
            "action": entry.get("action", "keep")
        }

    # Save
    OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(rebuilt, f, indent=2)

    print(f"💾 Saved → {OUTPUT_JSON}")
    print(f"✅ Matched entries: {matched}")
    print(f"⚠️ Missing matches: {len(missing_matches)}")

    if missing_matches:
        print("\nSample missing matches:")
        for m in missing_matches[:10]:
            print(f" - {m}")

    print()


if __name__ == "__main__":
    main()