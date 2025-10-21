# Scripts Documentation
## Robinson Phase 2 - Product Data Processing Pipeline

This document explains all scripts in the `scripts/` folder in sequential order.

---

## Script 1: `1-0-handle-extractor.py`

### What It Does
Searches for specific terms in collection data and extracts matching product handles.

**Process:**
1. Prompts you for search term(s) (e.g., "socks", "womens")
2. Searches **Column D** for rows containing those terms (case-insensitive)
3. Extracts values from **Column D** and **Column W** for matching rows
4. Outputs to a new Excel file with just those two columns

**Use Case:**  
*"Find all collections related to 'socks' and give me their handles"*

### Files & Paths Referenced

| Type | Path | Description |
|------|------|-------------|
| **Input** | `raw/collections-raw.xlsx` | Source collections file (hardcoded) |
| **Output** | `data/{search-term}-handles.xlsx` | Results file (dynamic based on search term) |

**Example:**
- Search for: `womens`
- Output: `data/womens-handles.xlsx`

---

## Script 2: `2-0-matrixify-segmenter.py`

### What It Does
Extracts full product data from the master products file based on a list of handles.

**Process:**
1. Prompts you for a reference file containing product handles in **Column B**
2. Reads all handles from that reference file
3. Searches the master products file for matching handles
4. Extracts **all columns** for matching products
5. Uses Unicode normalization to handle special characters (™, ®, ©, accents)
6. Preserves the original handles from the reference file in the output

**Use Case:**  
*"I have a list of 50 product handles - give me all the product data for those handles from the master file"*

### Files & Paths Referenced

| Type | Path | Description |
|------|------|-------------|
| **Input (Master)** | `raw/products-raw.xlsx` | Master products database (hardcoded) |
| **Input (Reference)** | User-provided path | File containing handles to extract |
| **Output** | `{same-dir-as-reference}/{reference-filename}-extracted-products.xlsx` | Full product data for matched handles |

**Example:**
- Reference file: `data/womens-handles.xlsx`
- Output: `data/womens-handles-extracted-products.xlsx`

**Special Features:**
- Normalizes handles for matching (removes ™, ®, accents, etc.)
- Preserves original handles in output
- Handles Unicode variations robustly

---

## Script 3: `3-0-size-extractor.py`

### What It Does
Analyzes an Excel file and extracts all unique size tags from product data.

**Process:**
1. Prompts you for an Excel file to analyze
2. Reads **Column H** (which contains comma-separated tags)
3. Extracts only tags that start with `size_`
4. Deduplicates and sorts them alphabetically
5. Outputs to a text file, one tag per line

**Use Case:**  
*"Show me all the unique size tags used in this product file"*

### Files & Paths Referenced

| Type | Path | Description |
|------|------|-------------|
| **Input** | User-provided path | Excel file to analyze |
| **Output** | `{same-dir-as-input}/size_tags.txt` | Sorted list of unique size tags |

**Example:**
- Input: `data/robinsons-socks.xlsx`
- Output: `data/size_tags.txt`

**Output Format:**
```
size_10-11
size_12-2
size_3-5.5
size_6-7
size_8-9
```

---

## Script 4: `4-0-brand-size-extractor.py`

### What It Does
Filters products by brand and size, updates their gender tags, and outputs only the matched rows in Shopify Matrixify format.

**Process:**
1. Prompts for **Brand Name** (exact case match)
2. Prompts for **Size Label** (must exist in comma-separated list)
3. Prompts for **Gender** to add/update
4. Searches all 30,000+ rows for matches
5. Updates gender tags using smart replacement logic
6. Outputs **only matched rows** to a filtered file
7. Applies Shopify Matrixify format (gender only on first handle occurrence)

**Gender Update Logic:**
| Current Gender | New Gender | Result |
|---|---|---|
| `["Female"]` | `Male` | `["Male"]` (opposite replaced) |
| `["Male"]` | `Male` | `["Male"]` (no change) |
| `["Unisex"]` | `Male` | `["Unisex", "Male"]` (appended) |
| `[]` | `Male` | `["Male"]` (added) |

**Use Case:**  
*"Find all Corgi products in size 6-7, tag them as Male, and give me a clean file ready for Shopify upload"*

### Files & Paths Referenced

| Type | Path | Description |
|------|------|-------------|
| **Input** | `raw/products-raw.xlsx` | Master products database (hardcoded, full path) |
| **Output** | `data/products-updated-{brand}-{size}.xlsx` | Filtered products with updated gender tags |

**Example:**
- Brand: `Corgi`, Size: `6-7`, Gender: `Male`
- Output: `data/products-updated-corgi-6-7.xlsx`

**Columns Used:**
- **Column B**: Handle (for uniqueness tracking)
- **Column F**: Brand (exact match filter)
- **Column H**: Size tags (comma-separated, substring match)
- **Column CQ**: Gender (JSON list format)

**Shopify Matrixify Format:**
```
Handle                  Brand   Size      Gender
corgi-city-socks       Corgi   6-7       ["Male"]   ← First occurrence
corgi-city-socks       Corgi   6-7                   ← Variant (blank)
corgi-city-socks       Corgi   6-7                   ← Variant (blank)
```

---

## Typical Workflow

Here's how these scripts work together in sequence:

```
1. Handle Extractor
   └─> Find collections by term (e.g., "womens socks")
       └─> Output: womens-socks-handles.xlsx

2. Matrixify Segmenter
   └─> Extract full product data for those handles
       └─> Output: womens-socks-handles-extracted-products.xlsx

3. Size Extractor (Optional)
   └─> Analyze what sizes exist in the extracted products
       └─> Output: size_tags.txt

4. Brand-Size Extractor (Optional)
   └─> Filter by specific brand/size and update gender tags
       └─> Output: products-updated-corgi-6-7.xlsx
```

---

## Key Design Principles

### 1. **Raw Data Protection**
- `raw/` folder files are **NEVER modified**
- All scripts read from `raw/`, output to `data/`
- Original data always preserved

### 2. **File Format**
- Always use `.xlsx` (Excel format)
- Never use CSV (data integrity issues with special characters)
- Preserves formatting and data types

### 3. **Shopify Matrixify Compliance**
- Product-level fields (like gender) only on first row per handle
- Variant rows have blank product-level fields
- Ready for direct Shopify import

### 4. **User-Friendly**
- Interactive prompts (no command-line arguments to remember)
- Progress indicators for large files
- Clear summary statistics after completion
- Error handling with helpful messages

---

## Column Reference Guide

| Column | Letter | Name | Used By |
|--------|--------|------|---------|
| 2 | B | Product: Handle | Script 2, 4 |
| 4 | D | Collection: Title | Script 1 |
| 6 | F | Vendor/Brand | Script 4 |
| 8 | H | Tags (comma-separated) | Script 3, 4 |
| 23 | W | Product: Handle | Script 1 |
| 95 | CQ | Metafield: Gender | Script 4 |

---

## Dependencies

All scripts require:
- Python 3.x
- `openpyxl` library

Install with:
```bash
pip install -r requirements.txt
```

---

## Notes

- **Case Sensitivity**: Script 4 uses exact case matching for brand names
- **Search Terms**: Script 1 uses case-insensitive substring matching
- **Unicode Handling**: Script 2 handles special characters (™, ®, ©, accents) intelligently
- **Large Files**: All scripts use `read_only=True` mode for memory efficiency with 30K+ row files

---

*Last Updated: October 21, 2025*

