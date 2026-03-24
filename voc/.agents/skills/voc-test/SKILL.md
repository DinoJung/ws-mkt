---
name: voc-test
description: VOC classification pipeline QA verification and regression testing workflow
triggers:
  - voc-test
  - voc 테스트
  - run tests
  - pytest
  - voc qa
  - 테스트 실행
  - 단위 테스트
  - 통합 테스트
  - regression
---

## Purpose

Use this skill when the user asks this AI to verify the VOC pipeline. The expected operating model is that this AI executes the relevant checks, captures the output, and reports the result back to the user instead of telling the user to run pytest manually.

**목적 (Korean):** 사용자가 검증을 요청하면, 이 AI가 VOC 분류 파이프라인 테스트와 관련 검증을 직접 실행하고 결과를 보고합니다.

Reference: [`docs/voc-classification-workflow.md §9`](docs/voc-classification-workflow.md) for verification guidance and failure handling notes.

## When to Activate

- after modifying `classify_voc.py` or test logic
- before finalizing a VOC change set
- when validating API retry/fallback behavior
- when checking that workbook preservation/output behavior still matches expectations
- when the user asks whether a VOC change is safe to proceed with

## Workflow

### 1. Full VOC Test Suite
When the user asks for full regression, this AI runs the repository's current VOC test suite and reports the real outcome.

Reference command executed by this AI:

```bash
cd /home/ws-mkt/voc
python3 -m pytest test_classify_voc.py -v
```

Do not hardcode an expected fixed test count in the skill text. The authoritative requirement is that the current suite passes with zero failures.

### 2. Targeted Verification
When a change only affects part of the workflow, this AI can run focused subsets and explain why that subset is sufficient before moving to broader verification.

Reference commands executed by this AI:

```bash
python3 -m pytest test_classify_voc.py -v -m integration
python3 -m pytest test_classify_voc.py -v -k "rule"
python3 -m pytest test_classify_voc.py -v -k "image"
python3 -m pytest test_classify_voc.py -v -k "fallback"
```

### 3. What This AI Must Report
After running verification, this AI should report:

- which command or subset was executed
- whether it passed or failed
- any relevant failure output or mismatch
- whether broader verification is still needed

### 4. Expected Test Behavior
- no real API calls in mocked tests
- no required user action during the normal verification flow
- original input workbooks remain untouched
- targeted checks are acceptable only when they are explicitly tied to the changed scope

## Verification

Success means the checks actually executed by this AI complete with zero unexpected failures.

Example evidence this AI should be able to report:

```text
- Full suite command executed successfully
- Targeted subset command executed successfully
- No new failures introduced by the current change
```

## Troubleshooting

### 429 Too Many Requests (Rate Limit)

**Symptom:** `HTTPError: 429 Client Error`

**Root Cause:** GitHub Models API rate limit exceeded during a real run or a non-mocked path.

**Fix:**
1. Check retry policy in `_classify_with_adapter()`.
2. Reduce repeated live calls before retrying.
3. Confirm that mocked tests are actually using `@patch("classify_voc.urllib.request.urlopen")`.

### Missing GITHUB_TOKEN Environment Variable

**Symptom:** `KeyError: 'GITHUB_TOKEN'` in `classify_voc.py`

**Fix reference command executed by this AI:**

```bash
cd /home/ws-mkt/voc
set -a && . ./.env && set +a
echo $GITHUB_TOKEN
python3 -m pytest test_classify_voc.py -v
```

### openpyxl Not Installed

**Symptom:** `ModuleNotFoundError: No module named 'openpyxl'`

**Fix reference command executed by this AI:**

```bash
pip install -r /home/ws-mkt/voc/requirements.txt
```

### Reporting Rule
This skill is not complete until this AI reports actual command output or actual pass/fail status. "This should pass" is not evidence.
