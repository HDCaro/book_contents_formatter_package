"""
===============================================================================
FILE: map_extracted_to_curated.py
LOCATION: src/index/fix/patch/
-------------------------------------------------------------------------------
Maps extracted index to curated keys

✔ Preserves curated structure
✔ Marks unmatched entries as "new"
✔ Generates clean HTML report (no fuzzy matching)

OUTPUT
------
index_curated_mapped.json
unmatched_entries.html
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

OUTPUT_JSON = BASE_PATH / "data/index/intermediate/index_curated_mapped.json"
OUTPUT_HTML = BASE_PATH / "data/index/output/unmatched_entries.html"

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

def norm(s):
    return s.strip().lower()

# ---------------------------------------------------------------------------
# LOAD
# ---------------------------------------------------------------------------

print("\n=== MAP EXTRACTED TO CURATED ===\n")

with open(INPUT_EXTRACTED, "r", encoding="utf-8") as f:
    extracted = json.load(f)

with open(INPUT_CURATED, "r", encoding="utf-8") as f:
    curated = json.load(f)

# ---------------------------------------------------------------------------
# BUILD LOOKUP
# ---------------------------------------------------------------------------

norm_to_key = {}

for key, entry in curated.items():
    n = entry.get("normalized")
    if n:
        norm_to_key[norm(n)] = key

# ---------------------------------------------------------------------------
# MAP
# ---------------------------------------------------------------------------

mapped = {}
unmatched = []

for key, entry in extracted.items():

    lookup_norm = norm(entry.get("normalized") or key)
    curated_key = norm_to_key.get(lookup_norm)

    # ---------------- MATCHED ---------------- #

    if curated_key:
        curated_entry = curated[curated_key]

        mapped[curated_key] = {
            "normalized": curated_entry.get("normalized"),
            "type": curated_entry.get("type"),
            "aliases": curated_entry.get("aliases", []),
            "aliases_external": curated_entry.get("aliases_external", []),
            "pages": entry.get("pages", []),
            "action": curated_entry.get("action", "keep")
        }

    # ---------------- UNMATCHED ---------------- #

    else:
        unmatched.append({
            "key": key,
            "normalized": entry.get("normalized") or key,
            "pages": entry.get("pages", [])
        })

        mapped[key] = {
            "normalized": entry.get("normalized") or key,
            "type": "unknown",
            "aliases": entry.get("aliases", []),
            "aliases_external": [],
            "pages": entry.get("pages", []),
            "action": "new"
        }

# ---------------------------------------------------------------------------
# SAVE JSON
# ---------------------------------------------------------------------------

OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)

with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
    json.dump(mapped, f, indent=2, ensure_ascii=False)

# ---------------------------------------------------------------------------
# GENERATE HTML REPORT
# ---------------------------------------------------------------------------

unmatched_sorted = sorted(unmatched, key=lambda x: x["key"].lower())

rows = ""

for e in unmatched_sorted:
    pages_str = ", ".join(map(str, e["pages"][:20]))
    if len(e["pages"]) > 20:
        pages_str += "..."

    rows += f"""
    <tr>
        <td>{e['key']}</td>
        <td>{e['normalized']}</td>
        <td>{pages_str}</td>
    </tr>
    """

html = f"""
<html>
<head>
<meta charset="utf-8">
<title>Unmatched Index Entries</title>
<style>
body {{ font-family: Arial; margin: 20px; }}
table {{ border-collapse: collapse; width: 100%; }}
th, td {{ border: 1px solid #ccc; padding: 6px; }}
th {{ background: #333; color: white; }}
tr:nth-child(even) {{ background: #f2f2f2; }}
</style>
</head>
<body>

<h2>Unmatched Entries ({len(unmatched_sorted)})</h2>

<table>
<tr>
<th>Key</th>
<th>Normalized</th>
<th>Pages</th>
</tr>

{rows}

</table>

</body>
</html>
"""

OUTPUT_HTML.parent.mkdir(parents=True, exist_ok=True)

with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
    f.write(html)

# ---------------------------------------------------------------------------
# REPORT
# ---------------------------------------------------------------------------

print("\n=== RESULT ===\n")
print(f"✅ Mapped entries: {len(mapped)}")
print(f"⚠️ Unmatched entries: {len(unmatched)}")
print(f"📄 HTML report: {OUTPUT_HTML}\n")

print("🎉 Mapping complete\n")