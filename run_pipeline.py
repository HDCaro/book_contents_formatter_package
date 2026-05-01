"""
===============================================================================
FILE: run_pipeline.py
-------------------------------------------------------------------------------
Pipeline runner with:
✔ Auto root detection
✔ Human checkpoints
✔ Diff viewer integration
===============================================================================
"""

import subprocess
import sys
from pathlib import Path

# ---------------- ROOT DETECTION ---------------- #

def find_project_root():
    current = Path(__file__).resolve()
    for parent in [current] + list(current.parents):
        if (parent / "src").exists() and (parent / "data").exists():
            return parent
    print("❌ Could not detect project root")
    sys.exit(1)

BASE_DIR = find_project_root()

print(f"\n📁 Project root: {BASE_DIR}\n")

# ---------------- SCRIPT PATHS ---------------- #

SCRIPTS = {
    "generate": BASE_DIR / "src/index/generate/generate_index_batch.py",
    "verify": BASE_DIR / "src/index/verify/verify_raw_index_with_word_com.py",
    "fix_raw": BASE_DIR / "src/index/fix/fix_raw_with_discrepancies.py",
    "revalidate_curated": BASE_DIR / "src/index/fix/revalidate_curated_pages.py",
    "compare": BASE_DIR / "src/index/fix/compare_curated_vs_raw.py",
    "merge": BASE_DIR / "src/index/fix/apply_final_merge.py",
    "diff": BASE_DIR / "src/index/fix/diff_index.py",
    "build": BASE_DIR / "src/index/build/build_index_docx.py",
}

# ---------------- UTIL ---------------- #

def run_step(name, path):
    print(f"\n🚀 STEP: {name}")
    print(f"📄 {path}\n")

    if not path.exists():
        print(f"❌ Script not found: {path}")
        sys.exit(1)

    result = subprocess.run([sys.executable, str(path)])

    if result.returncode != 0:
        print(f"\n❌ FAILED at step: {name}")
        sys.exit(result.returncode)

    print(f"✅ DONE: {name}\n")


def pause(message):
    print("\n" + "="*60)
    print("🛑 HUMAN CHECKPOINT")
    print(message)
    print("="*60)
    input("Press ENTER to continue...")


# ---------------- MAIN ---------------- #

def main():
    args = set(sys.argv[1:])

    print("=====================================")
    print("📚 INDEX PIPELINE START")
    print("=====================================")

    if "--skip-generate" not in args:
        run_step("Generate RAW index", SCRIPTS["generate"])

    if "--skip-verify" not in args:
        run_step("Verify RAW index", SCRIPTS["verify"])

    run_step("Fix RAW with discrepancies", SCRIPTS["fix_raw"])

    pause("""
Review RAW FIXED:
data/index/intermediate/index_raw_fixed.json
""")

    run_step("Revalidate curated pages", SCRIPTS["revalidate_curated"])

    pause("""
Review CURATED PAGES:
data/index/intermediate/index_curated_old_pages.json
""")

    if "--skip-compare" not in args:
        run_step("Compare curated vs raw", SCRIPTS["compare"])

        pause("""
Review comparison output:
✔ missing entries
✔ missing pages (should be 0)
""")

    run_step("Final merge", SCRIPTS["merge"])

    pause("""
Review FINAL JSON:
data/index/intermediate/index_curated_final.json
✔ merges correct
✔ no duplicates
✔ no missing key entries
""")

    run_step("Diff final index vs previous", SCRIPTS["diff"])

    pause("""
Review DIFF:
✔ unexpected removals
✔ unexpected additions
✔ page changes
""")

    run_step("Build DOCX", SCRIPTS["build"])

    print("=====================================")
    print("✅ PIPELINE COMPLETE")
    print("=====================================")


if __name__ == "__main__":
    main()