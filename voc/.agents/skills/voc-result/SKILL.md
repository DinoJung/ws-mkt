---
name: voc-result
description: VOC classification result writing and output saving to Excel workbook
triggers:
  - voc 결과 저장
  - result writing
  - 분류 결과
  - output save
  - 결과 기록
  - classification result
  - 엑셀 저장
---

## Purpose

Handle the result-writing stage of the VOC classification pipeline. This skill manages:
- Writing computed category values to column 7 (TARGET_COLUMN) in the classified worksheet
- Preserving the original input file (never modified in-place)
- Saving the completed workbook to `output/<filename>`
- Maintaining all sheets in the workbook during output
- Printing classification summary with category counts

Reference: [`docs/voc-classification-workflow.md §2`](docs/voc-classification-workflow.md) "입력/출력 규칙" for I/O path rules and [`§6`](docs/voc-classification-workflow.md) "실행 순서" steps 6–8 for execution sequence.

## When to Activate

Activate this skill when:
- Classification results need to be written to Excel cells
- Output must be saved to a separate file in the `output/` directory
- You need to ensure the original file remains untouched
- A summary report of classification counts is required

## Workflow

### Writing Results
The `write_classification_result(wb, row, category)` function writes a validated category string to the target cell at `ws.cell(row, TARGET_COLUMN)` where TARGET_COLUMN = 7. The input category is validated against VALID_CATEGORIES before writing.

### Processing Flow
`run_classification(wb, api_key, output_path)` orchestrates the full result-writing pipeline:
1. Builds image index from VOC sheets for vision-based fallback
2. Collects all target rows where col 7 contains "반응_제품반응"
3. Iterates through each row (1 to total count)
4. Calls `write_classification_result()` to write category to col 7
5. Increments category count in summary dictionary
6. Prints progress: `[idx/total] row=N (method) -> category`
7. After all rows processed, saves workbook to output_path via `wb.save(output_path)`
8. Outputs summary report showing counts per category

### Output Structure
- **Input file**: Original `.xlsx` file (never modified)
- **Output path**: `output/<input_filename>` (created by `--output-dir` argument, defaults to `output/`)
- **Sheets preserved**: All sheets from input workbook are present in output
- **Target sheet**: `2026 list` with classifications written to col 7

### Top-level Entry Point
`classify_workbook(wb, api_key, output_path)` is the wrapper that delegates to `run_classification()`. In typical usage, the CLI loads the workbook, creates output directory, and calls this function.

## Verification

To confirm result writing completed successfully:

```bash
# Check output file exists
test -f output/input_filename.xlsx && echo "Output file created" || echo "FAIL"

# Verify workbook has classifications
python3 -c "
import openpyxl
wb = openpyxl.load_workbook('output/input_filename.xlsx')
ws = wb['2026 list']
classified_count = sum(1 for r in range(5, 221) if ws.cell(r, 7).value and ws.cell(r, 7).value != '반응_제품반응')
print(f'Classifications written: {classified_count}')
"

# Verify original file untouched
python3 -c "
import openpyxl
wb = openpyxl.load_workbook('original_input.xlsx')
ws = wb['2026 list']
marker_count = sum(1 for r in range(5, 221) if ws.cell(r, 7).value == '반응_제품반응')
print(f'Original file markers: {marker_count} (should be unchanged)')
"
```
