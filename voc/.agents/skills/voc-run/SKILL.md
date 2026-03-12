---
name: voc-run
description: One-click orchestrator for complete VOC classification pipeline with dry-run validation and test verification
triggers:
  - voc-run
  - 원클릭 voc
  - voc 파이프라인 실행
  - voc 전체 실행
  - 분류 파이프라인
  - run voc pipeline
  - one-click voc
  - classify all
  - full pipeline
---

## Purpose

The `voc-run` skill activates the complete VOC classification pipeline orchestration. It:

1. **Loads environment variables** via `set -a && . ./.env && set +a` to populate `GITHUB_TOKEN`
2. **Validates dependencies** via `pip install -r requirements.txt`
3. **Executes dry-run** via `python3 classify_voc.py --dry-run` (auto-detects xlsx from `input/` directory) to validate data loading
4. **Runs full classification** via `python3 classify_voc.py` with hybrid rule+LLM+image fallback
5. **Verifies output** via `ls -la output/` to confirm result file generation
6. **Runs regression tests** via `python3 -m pytest test_classify_voc.py -v` to ensure 24/24 tests pass

This skill coordinates all four sub-skills (`voc-data-load`, `voc-classify`, `voc-result`, `voc-test`) into a single end-to-end execution.

**목적 (Korean):** VOC 분류 파이프라인 전체를 순차적으로 실행하여 데이터 로드 → 하이브리드 분류 → 결과 저장 → 테스트 검증을 한 번에 완료합니다.

## When to Activate

- **Processing a new VOC Excel file** end-to-end (input → classified output → verified)
- **Validating pipeline integrity** before production deployment
- **Troubleshooting classification failures** (dry-run → full run → test validation)
- **Automating batch VOC processing** in scheduled jobs or CI/CD
- **Verifying all 4 sub-skills work together** (orchestration verification)

## Workflow

### Step 1: Environment & Dependencies

Load `.env` and ensure all Python packages are available:

```bash
cd /home/ws-mkt/voc
set -a && . ./.env && set +a
pip install -r requirements.txt
```

This populates `GITHUB_TOKEN` (required by `run_classification()`).

### Step 2: Dry-Run Validation

**Sub-skill:** `voc-data-load`

Validate that target rows are loaded correctly without making API calls:

```bash
python3 classify_voc.py --dry-run
```

Expected output: `dry-run: N target rows` (N > 0)

The script auto-detects the xlsx file from the `input/` directory. You can also specify explicitly: `python3 classify_voc.py --dry-run input/file.xlsx`

This invokes `_build_parser()` to register `--dry-run` flag and `main()` to dispatch.

### Step 3: Full Classification Run

**Sub-skill:** `voc-classify`

Execute the hybrid classification pipeline:
- `main()` calls `run_classification(wb, api_key, output_path)`
- `run_classification()` iterates target rows and applies `_classify_item_with_method()`
- Workflow: rule-based (keywords) → LLM (GPT-4.1) → image OCR → fallback (GPT-4o)
- Reference: [VOC Classification Workflow §3, §6](https://github.com/code-yeongyu/voc/blob/main/docs/voc-classification-workflow.md)

```bash
python3 classify_voc.py [--output-dir output/]
```

### Step 4: Verify Output

**Sub-skill:** `voc-result`

Confirm that the classified Excel file was generated in `output/`:

```bash
ls -la output/
```

Expected: `output/<input-filename>` exists with updated classifications in column 7.

### Step 5: Regression Testing

**Sub-skill:** `voc-test`

Run all 24 unit + integration tests to verify pipeline correctness:

```bash
python3 -m pytest test_classify_voc.py -v
```

Expected: `24 passed` with zero failures. Tests validate rule classification, API adapters, image encoding, and error fallback.

## Complete Pipeline Example

```bash
#!/bin/bash
cd /home/ws-mkt/voc

# Step 1: Load environment
set -a && . ./.env && set +a

# Step 2: Install dependencies
pip install -r requirements.txt

# Step 3: Dry-run (data validation)
python3 classify_voc.py --dry-run

# Step 4: Full classification (auto-detects xlsx from input/)
python3 classify_voc.py

# Step 5: Verify output
ls -la output/

# Step 6: Run tests
python3 -m pytest test_classify_voc.py -v
```

## Verification

### Success Criteria

1. **Dry-run succeeds:** Output shows `dry-run: N target rows` with N > 0
2. **Classification completes:** No HTTP/timeout errors; classifications applied to column 7
3. **Output file created:** `output/<input-filename>` present and readable
4. **All tests pass:** `pytest` returns `24 passed in X.XXs`
5. **No console errors:** No `GITHUB_TOKEN` undefined, no module import errors

### Troubleshooting

| Issue | Fix |
|-------|-----|
| `KeyError: GITHUB_TOKEN` | Load `.env`: `set -a && . ./.env && set +a` |
| `ModuleNotFoundError: openpyxl` | Install deps: `pip install -r requirements.txt` |
| `HTTPError: 429` | Rate limit — wait 60s and retry (see `voc-classify` skill) |
| Tests fail | Check `test_classify_voc.py` fixtures and markers |

<!-- Sub-skills: voc-data-load, voc-classify, voc-result, voc-test | Core Functions: main(), _build_parser(), run_classification() -->
