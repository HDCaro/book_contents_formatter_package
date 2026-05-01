"""
===============================================================================
FILE: diff_index.py
-------------------------------------------------------------------------------
Human-friendly diff viewer for index_curated_final.json

✔ Shows added / removed / changed entries
✔ Field-level comparison
✔ Saves snapshot for next run (like WinMerge baseline)

USAGE
-----
python diff_index.py
or
python diff_index.py old.json new.json
===============================================================================
"""

import json
import sys
from pathlib import Path

# ---------------- ROOT DETECTION ---------------- #

def find_project_root():
    current = Path(__file__).resolve()
    for parent in [current] + list(current.parents):
        if (parent / "src").exists() and (parent / "data").exists():
            return parent
    raise RuntimeError("Project root not found")

BASE_DIR = find_project_root()

# ---------------- FILES ---------------- #

DEFAULT_OLD = BASE_DIR / "data/index/output/index_curated_final_prev.json"
DEFAULT_NEW = BASE_DIR / "data/index/intermediate/index_curated_final.json"

# ---------------- LOAD ---------------- #

def load(path):
    if not path.exists():
        return {}
    with open(path, encoding="utf-8") as f:
        return json.load(f)

# ---------------- COMPARE ---------------- #

def compare(old, new):
    old_keys = set(old.keys())
    new_keys = set(new.keys())

    added = sorted(new_keys - old_keys)
    removed = sorted(old_keys - new_keys)
    common = sorted(old_keys & new_keys)

    changed = []

    for k in common:
        if old[k] != new[k]:
            changed.append(k)

    return added, removed, changed

# ---------------- PRINT ---------------- #

def print_diff(old, new, added, removed, changed):
    print("\n================ DIFF VIEW ================\n")

    if added:
        print("🟢 ADDED\n")
        for k in added[:30]:
            print(f" + {k}")
        print(f"\nTotal added: {len(added)}\n")

    if removed:
        print("🔴 REMOVED\n")
        for k in removed[:30]:
            print(f" - {k}")
        print(f"\nTotal removed: {len(removed)}\n")

    if changed:
        print("🟡 CHANGED\n")

        for k in changed[:20]:
            print(f"\n🔹 {k}")
            o = old.get(k, {})
            n = new.get(k, {})

            fields = set(o.keys()) | set(n.keys())

            for f in sorted(fields):
                ov = o.get(f)
                nv = n.get(f)

                if ov != nv:
                    print(f"   {f}:")
                    print(f"      OLD → {ov}")
                    print(f"      NEW → {nv}")

        print(f"\nTotal changed: {len(changed)}\n")

    if not added and not removed and not changed:
        print("✅ No differences\n")

# ---------------- SAVE SNAPSHOT ---------------- #

def save_snapshot(new_data):
    DEFAULT_OLD.parent.mkdir(parents=True, exist_ok=True)
    with open(DEFAULT_OLD, "w", encoding="utf-8") as f:
        json.dump(new_data, f, indent=2)

# ---------------- MAIN ---------------- #

def main():

    if len(sys.argv) == 3:
        old_path = Path(sys.argv[1])
        new_path = Path(sys.argv[2])
    else:
        old_path = DEFAULT_OLD
        new_path = DEFAULT_NEW

    print(f"\nComparing:\nOLD → {old_path}\nNEW → {new_path}\n")

    old = load(old_path)
    new = load(new_path)

    added, removed, changed = compare(old, new)

    print_diff(old, new, added, removed, changed)

    save_snapshot(new)

    print("💾 Snapshot saved\n")


if __name__ == "__main__":
    main()