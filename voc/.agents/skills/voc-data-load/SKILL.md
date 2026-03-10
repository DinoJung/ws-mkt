---
name: voc-data-load
description: Load and preprocess VOC data from Excel workbook, extracting target rows and building image index
triggers:
  - voc 데이터 로드
  - excel preprocessing
  - 이미지 인덱싱
  - 엑셀 전처리
  - data load
  - build image index
---

## Purpose

The `voc-data-load` skill handles the data ingestion stage of the VOC classification pipeline. It:

1. Opens the Excel workbook using openpyxl
2. Filters rows in the `TARGET_SHEET` (`2026 list`) where column `TARGET_COLUMN` (7) equals `TARGET_MARKER` (`반응_제품반응`)
3. Extracts title and link from the target rows
4. Preprocesses titles to normalize whitespace and remove nulls
5. Builds an image index from VOC reaction sheets (`BP_반응`, `PK_반응`) to enable image-based classification fallback

This skill is the first stage before hybrid classification. Reference: [VOC Classification Workflow §2, §4, §6](https://github.com/code-yeongyu/voc/blob/main/docs/voc-classification-workflow.md).

## When to Activate

- **Loading a new VOC Excel file** for classification
- **Debugging target row extraction** (verify row count and title preprocessing)
- **Building image index** for linking images to VOC entries
- **Running dry-run validation** before full classification
- **Inspecting title normalization** for clean LLM input

## Workflow

### Step 1: Load Workbook & Identify Target Rows

The `read_target_rows()` function:
- Opens sheet `TARGET_SHEET` = `"2026 list"`
- Scans rows `START_ROW` (5) through `END_ROW` (220)
- Filters where `ws.cell(row, TARGET_COLUMN=7).value == TARGET_MARKER` (`"반응_제품반응"`)
- Extracts title from column `TITLE_COLUMN` (10) and link from column `LINK_COLUMN` (11)
- Returns list of dicts: `[{"row": int, "title": str, "link": str}, ...]`

### Step 2: Preprocess Titles

The `preprocess_title()` function:
- Removes newlines and extra whitespace
- Returns `None` if input is `None` or empty after cleaning
- Normalizes multiple spaces to single space
- Output used as LLM classification input

### Step 3: Build Image Index

The `build_image_index()` function:
- Iterates over reaction sheets in `REACTION_SHEETS` (`["BP_반응", "PK_반응"]`)
- Locates card boundaries by finding `"Voice Of Customer"` marker in row 2
- Extracts link from row 4, column `card_start + 4`
- Collects body images (row > 6) within card range using helper functions:
  - `_image_anchor_col_row()` — extracts column/row position from image anchor
  - `_extract_image_bytes()` — extracts image bytes from openpyxl image object
- Returns mapping: `{link: [image_bytes, ...], ...}`

### Step 4: Extract Link Text

The `_extract_link_text()` function:
- Extracts hyperlink target from openpyxl cell
- Falls back to cell value if no hyperlink present
- Used to obtain canonical link for image matching

## Verification

Run a dry-run to validate data loading without API calls:

```bash
cd /home/ws-mkt/voc
python3 classify_voc.py --dry-run /path/to/input.xlsx
```

Expected output: `dry-run: {N} target rows` where N > 0

For detailed inspection, load and print first target row:

```python
import openpyxl
from classify_voc import read_target_rows, build_image_index

wb = openpyxl.load_workbook("input.xlsx")
targets = read_target_rows(wb)
print(f"Found {len(targets)} targets")
print(f"First: {targets[0]}")

image_index = build_image_index(wb)
print(f"Image index has {len(image_index)} links")
```

Verify constants are set correctly:
- `TARGET_SHEET = "2026 list"`
- `TARGET_MARKER = "반응_제품반응"`
- `TARGET_COLUMN = 7` (filter column)
- `TITLE_COLUMN = 10` (classification input)
- `LINK_COLUMN = 11` (image matching key)
- `START_ROW = 5`, `END_ROW = 220` (scan range)
- `REACTION_SHEETS = ["BP_반응", "PK_반응"]` (image sources)
