# VOC Classification Pipeline

This directory contains the VOC classification workflow for reading a source Excel workbook, classifying target rows, and saving a processed workbook copy.

## Directory roles

- `input/`: place source `.xlsx` workbooks here for local runs. These files are treated as raw inputs and are now git-ignored so source data is not added to the repository by mistake.
- `output/`: generated result workbooks are written here. These files are produced by the pipeline, represent processed copies of the input workbook, and are git-ignored so generated artifacts stay out of the repository.
- `docs/`: workflow notes and classification logic reference documents.

## Basic run flow

1. Put the source workbook in `input/` or point the script at another local `.xlsx` path.
2. Run `classify_voc.py` with the workbook path.
3. Review the generated workbook in `output/`.

## Example

```bash
cd /home/ws-mkt/voc
set -a && . ./.env && set +a
python3 classify_voc.py "input/your-file.xlsx"
```

For the full classification rules and processing order, see `docs/voc-classification-workflow.md`.
