# 📚 Book Index Pipeline — Final Architecture

This pipeline builds a **fully curated, publication-ready book index** using a combination of:

* Word COM (pagination truth)
* Regex discovery (coverage)
* Editorial curation (identity + intent)

---

# 🧭 Pipeline Overview

```
DOCX → RAW → FIX → CURATED PAGES → MERGE → DOCX INDEX
```

---

# 🧩 Step-by-Step Pipeline

---

## 1️⃣ generate_index_batch.py

### Purpose

Discover index candidates from the book and assign page numbers using Word COM.

### Input

```
data/index/input/HITS AND HAPPINESS FINAL 2 Format MOM Discog.docx
```

### Output

#### 📄 index_raw.json

```
Raw detected entries with page numbers.
```

**Structure**

```json
{
  "Paul McCartney": {
    "normalized": "McCartney, Paul",
    "type": "person",
    "pages": [2, 57, 101],
    "action": "keep"
  }
}
```

#### 📄 index_raw_excluded.json

```
Entries with fewer than minimum pages (e.g. single mentions).
```

#### 📄 index_raw_filtered_out.json

```
Entries with too many occurrences (noise / overly frequent terms).
```

---

## 2️⃣ fix_raw_with_discrepancies.py

### Purpose

Correct raw index pages using verification discrepancies.

### Input

```
index_raw.json
index_discrepancies.json
```

### Output

#### 📄 index_transaction_edit.json

```
Editable working dataset with corrected page numbers.
```

**Structure**

```json
{
  "Paul McCartney": {
    "normalized": "McCartney, Paul",
    "type": "person",
    "pages": [2, 57, 101, 224, 305],
    "action": "keep"
  }
}
```

---

## 3️⃣ revalidate_curated_pages.py ⭐ CRITICAL

### Purpose

Recalculate page numbers for curated entries using Word COM.

This ensures:

* correct pagination
* alias coverage
* no dependency on previous raw errors

### Input

```
index_curated_old_filtered.json
DOCX file
```

### Output

#### 📄 index_curated_old_pages.json

```
Curated entries with CORRECT page numbers.
```

**Structure**

```json
{
  "Pat Silver": {
    "normalized": "Silver-Lasky, Pat",
    "type": "person",
    "aliases": ["Barbara Carleton", "Barbara Hayden"],
    "aliases_external": ["Barbara Romano"],
    "pages": [2, 3, 5, 6, 7, 11, 12],
    "action": "merge"
  }
}
```

---

## 4️⃣ apply_curated_identity_merge.py ⭐ CORE

### Purpose

Final merge combining:

* Transaction data → full page coverage
* Curated data → identity, aliases, editorial intent

### Input

```
index_transaction_edit.json
index_curated_old_pages.json
```

### Output

#### 📄 index_curated_final.json

```
Final index dataset ready for rendering.
```

**Structure**

```json
{
  "Pat Silver-Lasky": {
    "normalized": "Silver-Lasky, Pat",
    "type": "person",
    "pages": [2, 3, 5, 6, 7, 11, 12, 15, 16],
    "aliases": ["Barbara Carleton", "Barbara Hayden"],
    "aliases_external": ["Barbara Romano"]
  }
}
```

---

## 5️⃣ build_index_docx.py

### Purpose

Render final index into Word document format.

### Input

```
index_curated_final.json
```

### Output

#### 📄 index.docx

```
Final formatted book index ready for publishing (KDP / print).
```

---

# 🧠 Data Model Overview

---

## Entry Fields

```json
{
  "normalized": "Display name (used for sorting)",
  "type": "person | band | work | unknown",
  "pages": [list of page numbers],
  "action": "keep | merge | remove",
  "aliases": ["internal aliases found in book"],
  "aliases_external": ["external/known aliases"],
  "destination": "target entry if merged"
}
```

---

# 🔑 Key Design Principles

---

## 1. Pages = Word COM truth

```
All page numbers are derived from Word, never manually edited.
```

---

## 2. Curated = identity truth

```
Curated files define:
- who is who
- what merges
- what gets removed
```

---

## 3. Raw = coverage

```
Raw discovery ensures nothing is missed.
```

---

## 4. Merge = controlled combination

```
Final dataset is built by combining:
- raw coverage
- curated structure
```

---

# ⚠️ Important Rules

---

### ❌ Never merge using normalized name

```
Causes data loss (e.g. Pat Silver issue)
```

---

### ✅ Always merge using identity groups

```
Aliases + destination define identity
```

---

### ❌ Do not trust raw pagination alone

```
Must be corrected or revalidated
```

---

### ✅ Always revalidate curated pages

```
This guarantees correctness before final merge
```

---

# 🧹 Removed / Deprecated Scripts

---

### ❌ revalidate_missing_entries.py

```
Replaced by full curated revalidation
```

### ❌ merge_curated_into_transaction.py

```
Used once for migration only
```

---

# 🚀 Final Pipeline (Minimal)

```
1. generate_index_batch.py
2. fix_raw_with_discrepancies.py
3. revalidate_curated_pages.py
4. apply_curated_identity_merge.py
5. build_index_docx.py
```

---

# 🎯 Outcome

```
✔ Accurate page numbers
✔ Correct identity merges
✔ No lost entries
✔ Editorial control preserved
✔ Production-ready index
```

---

# 📌 Notes

This pipeline is designed for:

* Books with complex references
* Music / film / historical indexing
* Alias-heavy datasets
* Professional publishing workflows

---

# 🧠 Final Insight

This system works because it separates:

```
DATA        → Word (pages)
DISCOVERY   → Regex (coverage)
EDITORIAL   → Curated (identity)
OUTPUT      → DOCX (presentation)
```

---

# ✅ Ready for Production
