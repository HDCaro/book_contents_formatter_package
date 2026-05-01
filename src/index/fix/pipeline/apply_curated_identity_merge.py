"""
===============================================================================
FILE: apply_curated_identity_merge.py (FIXED VERSION)
-------------------------------------------------------------------------------
Correct identity resolution using GROUPS (not flat mapping)

Fixes:
✔ alias propagation
✔ destination propagation
✔ Pat Silver loss
✔ multi-source merging

===============================================================================
"""

import json
from pathlib import Path
from collections import defaultdict

# ---------------- CONFIG ---------------- #

BASE_DIR = Path(__file__).resolve().parents[3]

TRANSACTION_JSON = BASE_DIR / "data/index/intermediate/index_transaction_edit.json"
CURATED_JSON = BASE_DIR / "data/index/intermediate/index_curated_old_filtered.json"

OUTPUT_JSON = BASE_DIR / "data/index/intermediate/index_curated_final.json"

# ---------------- LOAD ---------------- #

def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

# ---------------- UNION-FIND (DISJOINT SET) ---------------- #

class UnionFind:
    def __init__(self):
        self.parent = {}

    def find(self, x):
        if x not in self.parent:
            self.parent[x] = x
        if self.parent[x] != x:
            self.parent[x] = self.find(self.parent[x])
        return self.parent[x]

    def union(self, a, b):
        rootA = self.find(a)
        rootB = self.find(b)
        if rootA != rootB:
            self.parent[rootB] = rootA

# ---------------- BUILD GROUPS ---------------- #

def build_identity_groups(curated):
    uf = UnionFind()

    for key, entry in curated.items():

        # Always connect key to itself
        uf.find(key)

        # --- destination ---
        dest = entry.get("destination")
        if entry.get("action") == "merge" and dest:
            uf.union(key, dest)

        # --- aliases ---
        for alias in entry.get("aliases", []):
            uf.union(key, alias)

        for alias in entry.get("aliases_external", []):
            uf.union(key, alias)

    # --- Build groups ---
    groups = defaultdict(set)

    for node in uf.parent:
        root = uf.find(node)
        groups[root].add(node)

    return groups

# ---------------- CHOOSE CANONICAL ---------------- #

def choose_canonical(group, curated):
    # Prefer entry that exists in curated and has no merge action
    for name in group:
        entry = curated.get(name)
        if entry and entry.get("action") != "merge":
            return name

    # fallback: any curated entry
    for name in group:
        if name in curated:
            return name

    # fallback: deterministic
    return sorted(group)[0]

# ---------------- MAIN ---------------- #

def main():
    print("\n=== APPLY CURATED IDENTITY MERGE (FIXED) ===\n")

    tx = load_json(TRANSACTION_JSON)
    curated = load_json(CURATED_JSON)

    groups = build_identity_groups(curated)

    final = {}

    print("=== GROUPS ===\n")

    for root, group in groups.items():

        canonical = choose_canonical(group, curated)

        print(f"🔗 {canonical}")
        for name in sorted(group):
            print(f"   - {name}")

        # --- MERGE PAGES ---
        pages = set()

        for name in group:
            if name in tx:
                pages.update(tx[name].get("pages", []))

        pages = sorted(pages)

        # --- BUILD ENTRY ---
        cur = curated.get(canonical, {})

        entry = {
            "pages": pages,
            "normalized": cur.get("normalized", canonical),
            "type": cur.get("type", "unknown")
        }

        # preserve editorial fields
        if "action" in cur:
            entry["action"] = cur["action"]

        if "aliases" in cur:
            entry["aliases"] = cur["aliases"]

        if "aliases_external" in cur:
            entry["aliases_external"] = cur["aliases_external"]

        final[canonical] = entry

    print(f"\n✔ Final entries: {len(final)}")

    # ---------------- SAVE ---------------- #

    OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(final, f, indent=2)

    print(f"\n💾 Saved → {OUTPUT_JSON}\n")


if __name__ == "__main__":
    main()