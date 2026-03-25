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

Use this skill when the user asks this AI to complete the result-writing stage of the VOC pipeline. In the current runtime model, this AI is responsible for:

- collecting deterministic workbook updates from classification results
- preserving the original input workbook in-place
- writing the production result workbook to `output/<yymmdd_main_input_filename>` through the preservation writer
- keeping QA artifact paths separate under `output/qa/<artifact_name>/<input_filename>`
- reporting where the production result lives and whether any OMV mirror copy was performed as a procedural follow-up

Reference: [`docs/voc-classification-workflow.md §2`](docs/voc-classification-workflow.md) for I/O path rules and [`§7`](docs/voc-classification-workflow.md) for execution order.

## When to Activate

Activate this skill when:
- the user asks this AI to write VOC classification results into the workbook output
- output locations need to be explained or verified
- production output and QA artifact boundaries must be kept distinct
- you need to confirm that the preservation writer, not legacy openpyxl save behavior, is the active output path

## Workflow

### Result Collection
This AI should treat classification output as a collected update set, not as “save whatever the in-memory workbook currently looks like.”

- `collect_classification_updates(wb, api_key)` gathers:
  - `2026 list` target-cell writes
  - linked-card fan-out updates
  - summary counts per final category

### Output Writing
When the user asks for a real output workbook, the runtime path is:

1. `run_classification(...)` collects updates and summary data
2. `save_workbook_with_preserved_media(source_path, output_path, updates)` selects the supported writer contract
3. `select_output_writer(...)` validates workbook-family support and writable-surface rules
4. `_write_workbook_value_only_copy(...)` copies the source `.xlsx` archive and patches only allowed worksheet values

This means the authoritative production output is the preserved copy at `output/<yymmdd_main_input_filename>`, not a legacy `wb.save(output_path)` write.

### Output Boundaries
- **Input file**: original `.xlsx`, never overwritten in place
- **Production output**: `output/<yymmdd_main_input_filename>`
- **QA artifact output**: `output/qa/<artifact_name>/<input_filename>`
- **OMV mirror**: `/mnt/omv/.workspace/ws-mkt/voc/result/<input_filename>` only when this AI performs the extra procedural copy after the production output exists

Important boundary: OMV mirroring is currently a workflow/procedure step performed by this AI after output creation. It is not yet an automatic runtime write inside `classify_voc.py`.

### Top-level Entry Point
`classify_workbook(wb, api_key, output_path, source_path=...)` delegates to `run_classification(...)`. When the user asks this AI to complete a real run, this AI should validate:

- source workbook path
- output path chosen by `build_output_workbook_path(...)`
- whether any QA artifact path is intentionally requested via `build_qa_artifact_workbook_path(...)`
- whether OMV mirroring should be performed as the post-run copy step

## Verification

To confirm result writing completed successfully, this AI should verify all applicable layers:

```bash
# Production output exists
test -f output/$(date +%y%m%d)_main_input_filename.xlsx && echo "Production output created" || echo "FAIL"

# Output workbook is readable and contains written classifications
python3 -c "
import openpyxl
wb = openpyxl.load_workbook('output/250326_main_input_filename.xlsx')
ws = wb['2026 list']
classified_count = sum(1 for r in range(5, 221) if ws.cell(r, 7).value and ws.cell(r, 7).value != '반응_제품반응')
print(f'Classifications written: {classified_count}')
"

# Optional: QA artifact path stays segregated
python3 -c "
from pathlib import Path
print(Path('output/qa/manual-QA/input_filename.xlsx'))
"
```

Expected verification outcome:
- production output is present under `output/`
- original input workbook remains untouched
- QA artifacts, when used, stay under `output/qa/`
- OMV mirror, when requested by workflow, is reported separately from production output rather than being confused with the runtime writer itself
