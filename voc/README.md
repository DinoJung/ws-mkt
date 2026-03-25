# VOC Classification Pipeline

This directory contains the VOC classification workflow for reading a source Excel workbook, classifying target rows, and saving a processed workbook copy.

## Directory roles

- `input/`: place source `.xlsx` workbooks here for local runs. These files are treated as raw inputs and are now git-ignored so source data is not added to the repository by mistake.
- `output/`: generated result workbooks are written here. Production outputs follow `yymmdd_main_<input-filename>.xlsx`, and these generated artifacts are git-ignored so they stay out of the repository.
- `docs/`: workflow notes and classification logic reference documents.

## Basic run flow

1. Put the source workbook in `input/` or tell this AI which local `.xlsx` file to use.
2. Ask this AI to run the VOC workflow.
3. This AI performs dry-run, full execution, verification, and result reporting.
4. Review the generated production workbook in `output/` using the `yymmdd_main_<input-filename>` naming rule and, when requested by workflow, the mirrored copy under `/mnt/omv/.j2nu/ws-mkt/voc/result/`.

## Windows Compatibility & Release Gating

The `.xlsx` output is generated on Linux for performance. To ensure compatibility:
- All releases require a `Windows Excel Validation Gate` signoff.
- This gate uses a self-hosted Windows runner with real Excel (via COM) to verify no repair warnings exist.
- The gate can resolve artifacts via explicit path, optional downloaded artifact name, or newest `.xlsx` discovery from configured roots.
- Validation now records concrete repair evidence signals (new repair-log artifacts and normal-open vs repair-open behavior), not just single-cell access.
- Reference: `.sisyphus/plans/xlsx-output-compat-performance.md` Task 8.


## Reference command example

Normal operation does not require the user to run the command below directly. This is only a reference for what this AI executes internally when needed.

```bash
cd /home/ws-mkt/voc
set -a && . ./.env && set +a
python3 classify_voc.py "input/your-file.xlsx"
```

For the full classification rules and processing order, see `docs/voc-classification-workflow.md`.
