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

Activate this skill to run comprehensive QA verification for the VOC classification pipeline. The `voc-test` workflow validates both unit and integration tests, verifies API fallback mechanisms, and ensures the hybrid classification pipeline works correctly with mock-based assertions.

**목적 (Korean):** VOC 분류 파이프라인의 모든 테스트를 자동으로 실행하고 검증합니다. 규칙 기반 분류, LLM 제목 분류, 이미지 OCR 재분류, 그리고 API 에러 폴백 로직을 검증합니다.

## When to Activate

- After modifying `classify_voc.py` or test logic
- Before committing changes to the VOC pipeline
- When validating API retry/fallback behavior (429 rate limit scenarios)
- When verifying hybrid classification output (rule → LLM → image OCR fallback)
- During CI/CD pipeline as regression validation
- To ensure `openpyxl` workbook mutations don't corrupt original files

## Workflow

### 1. Full Test Suite

**Command:**
```bash
cd /home/ws-mkt/voc
python3 -m pytest test_classify_voc.py -v
```

Runs all 24 tests (19 unit + 5 integration):
- `TestVocClassifier::test_rule_classify_*` — rule-based classification
- `TestVocClassifier::test_github_gpt*_adapter_*` — API adapter format/parsing
- `TestVocClassifier::test_classify_pipeline_with_mock_api` — full integration
- `TestVocClassifier::test_hybrid_pipeline_integration` — rule → LLM hybrid flow
- `TestVocClassifier::test_api_error_fallback_full_pipeline` — GPT-4.1 → GPT-4o fallback

### 2. Integration Tests Only

**Command:**
```bash
python3 -m pytest test_classify_voc.py -v -m integration
```

Validates 5 integration markers:
- `test_classify_pipeline_with_mock_api`
- `test_output_preserves_all_sheets`
- `test_classification_summary`
- `test_api_error_fallback_full_pipeline`
- `test_hybrid_pipeline_integration`

### 3. Rule-Based Classification Tests

**Command:**
```bash
python3 -m pytest test_classify_voc.py -v -k "rule"
```

Tests keywords `[교환]` and `[구해요]` are caught before API calls.

### 4. Image & Vision Tests

**Command:**
```bash
python3 -m pytest test_classify_voc.py -v -k "image"
```

Validates image encoding, base64 conversion, and vision payload structure.

### 5. Fallback & Error Handling Tests

**Command:**
```bash
python3 -m pytest test_classify_voc.py -v -k "fallback"
```

Confirms HTTP errors trigger GPT-4.1 → GPT-4o retry mechanism.

## Verification

### Success Criteria

All tests pass with zero failures:
```
collected 24 items
test_classify_voc.py::TestVocClassifier::test_rule_classify_exchange PASSED
test_classify_voc.py::TestVocClassifier::test_rule_classify_seeking PASSED
...
test_classify_voc.py::test_hybrid_pipeline_integration PASSED
=================== 24 passed in X.XXs ====================
```

### Expected Test Behavior

- **No real API calls** — all tests use mock (unittest.mock.patch)
- **No GITHUB_TOKEN required** — mocks intercept urllib.request.urlopen
- **Workbook mutations are in-memory** — original input not modified
- **Fast execution** — all 18 tests complete in < 2 seconds

## Troubleshooting

### 429 Too Many Requests (Rate Limit)

**Symptom:** `HTTPError: 429 Client Error`

**Root Cause:** GitHub Models API rate limit exceeded (e.g., running real pipeline against live API)

**Fix:**
1. Check retry policy in `_classify_with_adapter()` — should have 3 retries with exponential backoff (1s → 2s → 4s)
2. Add delay between classifications: `import time; time.sleep(1)` between API calls
3. For tests, verify mocks are applied: `@patch("classify_voc.urllib.request.urlopen")`

### Missing GITHUB_TOKEN Environment Variable

**Symptom:** `KeyError: 'GITHUB_TOKEN'` in classify_voc.py

**Root Cause:** API key not loaded from `.env` file

**Fix:**
```bash
cd /home/ws-mkt/voc
set -a && . ./.env && set +a
echo $GITHUB_TOKEN  # verify loaded
python3 -m pytest test_classify_voc.py -v
```

### openpyxl Not Installed

**Symptom:** `ModuleNotFoundError: No module named 'openpyxl'`

**Fix:**
```bash
pip install -r /home/ws-mkt/voc/requirements.txt
# or
pip install openpyxl
```

---

**마지막 수정:** 2026-03-10 | **테스트 클래스:** TestVocClassifier | **통합 테스트 마커:** @pytest.mark.integration
