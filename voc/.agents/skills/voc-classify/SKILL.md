---
name: voc-classify
description: Hybrid classification pipeline for VOC items using rules, LLM, and image OCR
triggers:
  - voc 분류
  - hybrid classification
  - 하이브리드 분류
  - rule classify
  - 규칙 분류
  - gpt classify
  - image ocr
  - 이미지 분류
---

## Purpose

The `voc-classify` skill documents the hybrid classification pipeline that intelligently routes VOC (Voice of Customer) titles and images through a 4-step decision tree:
1. Keyword rule matching (no API call)
2. LLM text classification (GPT-4.1)
3. Image OCR fallback (GPT-4.1 vision → GPT-4o)
4. Error recovery

This balances cost (rules-first), accuracy (LLM), and comprehensiveness (image analysis).

## When to Activate

- Need to understand the VOC classification flow in `/home/ws-mkt/voc/classify_voc.py`
- Implementing or debugging `_classify_item_with_method()` logic
- Configuring retry policies, model selection, or image matching
- Optimizing classification accuracy for edge cases
- Tracing classification decisions for a specific VOC item

## Workflow

### Step A: Rule-Based Classification (cost: $0)

**Function:** `rule_classify(title: str) → str | None`

Checks title against `EXCHANGE_KEYWORDS` and `SEEKING_KEYWORDS`:
- If title contains `[교환]` or similar → return `반응_교환` immediately
- If title contains `[구해요]` or similar → return `반응_구해요` immediately
- Otherwise → return `None` (proceed to Step B)

**Cost:** No API calls. ~90% of routine posts match here.

---

### Step B: LLM Title Classification (cost: ~0.1¢ per call)

**Function:** `_classify_with_adapter(adapter, title) → str | None`

Uses `GithubModelsAdapter` with:
- **Model:** `openai/gpt-4.1` (text-only)
- **Prompt:** `build_classification_prompt(title)` includes 15 few-shot examples, 5 category definitions, and importance of tags
- **Retry policy:** 3 attempts, exponential backoff (1s → 2s → 4s), 0.5s inter-call delay
- **Response parsing:** `adapter.parse_response()` → `validate_result()` maps text to `VALID_CATEGORIES`
- **Returns:** Exact category string (e.g., `반응_문의`) or `None` on HTTP/network error

**When to advance to Step C:**
- Result is `None` (API failure, triggering fallback)
- Result is `반응_기타` AND image matching succeeds

---

### Step C: Image OCR Classification (cost: ~1¢ per call)

**Function:** `classify_with_image(image_bytes_list, api_key) → str | None`

Triggered only if Step B returns `반응_기타` AND `image_index[link]` has images.

1. Extract image bytes from VOC sheets (`BP_반응`, `PK_반응`)
2. Call `_classify_with_image_model()` with GPT-4.1 vision first
3. If GPT-4.1 fails → fallback to GPT-4o
4. Vision prompt: `_build_vision_request_body()` (includes base64 images + classification instruction)
5. **Returns:** Category or `None` on both model failures

**Image selection:** Only body images (`row > 6`) within card bounds

---

### Step D: Error-Only Model Fallback (cost: ~0.1¢)

**Function:** `_classify_with_adapter()` with `openai/gpt-4o`

Fallback **only** if:
- Step B returned `None` (HTTP/network error, not classification result)
- Image step (if attempted) returned `None`

**Does NOT trigger on:**
- `반응_기타` classification result (this is valid output)
- Parsing errors (returns default `반응_기타`)

**Final orchestrator:** `_classify_item_with_method()` calls steps sequentially and returns `(category, method_name)` tuple for audit.

## Verification

Verify classification accuracy and pipeline behavior:

1. **Unit tests:** `pytest test_classify_voc.py -v`
2. **Live dry-run:** `python3 classify_voc.py input.xlsx --dry-run`
3. **Full run:** Set `GITHUB_TOKEN` env var, then `python3 classify_voc.py input.xlsx`
4. **Output check:** Verify `output/input.xlsx` column 7 contains only `VALID_CATEGORIES` values
5. **Audit trail:** Check stderr logs for `[row] (method) → category` per item

**Success criteria:**
- All 5 categories present in results
- No invalid category strings
- Classification method trace matches expected flow (rule → llm → image → fallback)
