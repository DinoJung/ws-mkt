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

Use this skill when the user asks this AI to run the VOC pipeline end-to-end. The normal operating model is:

1. The **user requests** a VOC run from this AI.
2. This AI **loads the environment and validates dependencies**.
3. This AI **runs dry-run** to confirm target-row discovery.
4. This AI **runs full classification** with the existing hybrid rule+LLM+image fallback flow.
5. This AI **stores the main result workbook** under `output/` and then mirrors that same main artifact to `/mnt/omv/.j2nu/ws-mkt/voc/result/`.
6. This AI **verifies outputs and test results** and reports back what happened.

This skill coordinates all four sub-skills (`voc-data-load`, `voc-classify`, `voc-result`, `voc-test`) into a single AI-executed workflow after the user's request.

**목적 (Korean):** 사용자가 VOC 실행을 요청하면, 이 AI가 데이터 검증 → 하이브리드 분류 → 결과 저장 → OMV result 경로 추가 저장 → 테스트 검증까지 순차적으로 수행하고 결과를 보고합니다.

## When to Activate

- **Processing a new VOC Excel file** end-to-end (input → classified output → verified)
- **Validating pipeline integrity** before production deployment
- **Troubleshooting classification failures** (dry-run → full run → test validation)
- **Automating batch VOC processing** in scheduled jobs or CI/CD
- **Verifying all 4 sub-skills work together** (orchestration verification)

## Workflow

### Step 1: Environment & Dependencies

When the user asks for a run, this AI first loads `.env` and ensures the required Python packages are available.

Reference command executed by this AI:

```bash
cd /home/ws-mkt/voc
set -a && . ./.env && set +a
pip install -r requirements.txt
```

This populates `GITHUB_TOKEN` (required by `run_classification()`).

### Step 2: Dry-Run Validation

**Sub-skill:** `voc-data-load`

This AI validates that target rows are loaded correctly without making API calls.

Reference command executed by this AI:

```bash
python3 classify_voc.py --dry-run
```

Expected output: `dry-run: N target rows` (N > 0)

The script auto-detects the xlsx file from the `input/` directory. This AI can also specify an explicit path when the user points to a particular workbook.

This invokes `_build_parser()` to register `--dry-run` flag and `main()` to dispatch.

### Step 3: Full Classification Run

**Sub-skill:** `voc-classify`

This AI executes the hybrid classification pipeline after the user's request:
- `main()` calls `run_classification(wb, api_key, output_path)`
- `run_classification()` iterates target rows and applies `_classify_item_with_method()`
- Workflow: rule-based (keywords) → LLM (GPT-4.1) → image OCR → fallback (GPT-4o)
- Reference: [VOC Classification Workflow §3, §6](https://github.com/code-yeongyu/voc/blob/main/docs/voc-classification-workflow.md)

```bash
python3 classify_voc.py [--output-dir output/]
```

### Step 4: Mirror Main Output to OMV Result Directory

**Sub-skill:** `voc-result`

After the main workbook is written to `output/`, this AI copies that exact production result workbook to the OMV result directory with the same basename.

Reference command executed by this AI:

```bash
INPUT_NAME="$(basename "$(ls input/*.xlsx)")"
RESULT_DIR="/mnt/omv/.j2nu/ws-mkt/voc/result"
mkdir -p "$RESULT_DIR"
cp "output/$INPUT_NAME" "$RESULT_DIR/$INPUT_NAME"
```

This keeps the repo-local output contract (`output/<input-filename>`) while also ensuring the final main artifact is available in the shared OMV destination. This is currently a procedural step performed by this AI, not a runtime auto-mirror built into `classify_voc.py`.

### Step 5: Verify Output

**Sub-skill:** `voc-result`

This AI confirms that the classified Excel file exists in both `output/` and the OMV result directory.

Reference commands executed by this AI:

```bash
ls -la output/
ls -la /mnt/omv/.j2nu/ws-mkt/voc/result/
```

Expected: the same main workbook filename exists in both locations, with `output/<input-filename>` as the primary repo-local artifact and `/mnt/omv/.j2nu/ws-mkt/voc/result/<input-filename>` as the mirrored copy.

### Step 6: Regression Testing

**Sub-skill:** `voc-test`

This AI runs the current unit + integration suite to verify pipeline correctness.

Reference command executed by this AI:

```bash
python3 -m pytest test_classify_voc.py -v
```

Expected: the current pytest suite passes with zero failures. Tests validate rule classification, API adapters, image encoding, error fallback, and workbook preservation behavior.

## Complete AI-Executed Example

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

# Step 5: Mirror final main output to OMV result
INPUT_FILE="$(basename "$(ls input/*.xlsx)")"
RESULT_DIR="/mnt/omv/.j2nu/ws-mkt/voc/result"
mkdir -p "$RESULT_DIR"
cp "output/$INPUT_FILE" "$RESULT_DIR/$INPUT_FILE"

# Step 6: Verify output
ls -la output/
ls -la "$RESULT_DIR"

# Step 7: Run tests
python3 -m pytest test_classify_voc.py -v
```

## Verification

### Success Criteria

1. **Dry-run succeeds:** Output shows `dry-run: N target rows` with N > 0
2. **Classification completes:** No HTTP/timeout errors; classifications applied to column 7
3. **Main output created locally:** `output/<input-filename>` present and readable
4. **Main output mirrored to OMV:** `/mnt/omv/.j2nu/ws-mkt/voc/result/<input-filename>` present and readable
5. **All tests pass:** `pytest` returns the current suite with zero failures
6. **No console errors:** No `GITHUB_TOKEN` undefined, no module import errors
7. **Clear report back to the user:** this AI reports dry-run result, main output path, OMV mirror path, and verification outcome

### Troubleshooting

| Issue | Fix |
|-------|-----|
| `KeyError: GITHUB_TOKEN` | Load `.env`: `set -a && . ./.env && set +a` |
| `ModuleNotFoundError: openpyxl` | Install deps: `pip install -r requirements.txt` |
| OMV result copy missing | Create the directory with `mkdir -p /mnt/omv/.j2nu/ws-mkt/voc/result` and re-copy the final `output/<input-filename>` workbook |
| `HTTPError: 429` | Rate limit — wait 60s and retry (see `voc-classify` skill) |
| Tests fail | Check `test_classify_voc.py` fixtures and markers |

<!-- Sub-skills: voc-data-load, voc-classify, voc-result, voc-test | Core Functions: main(), _build_parser(), run_classification() -->
