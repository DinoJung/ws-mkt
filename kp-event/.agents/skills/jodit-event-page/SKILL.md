---
name: jodit-event-page
description: Generate Jodit Editor-ready event-page HTML from a fixed Korean question flow and save it to output with deterministic filenames.
triggers:
  - 이벤트 페이지만들자
  - 이벤트 페이지 만들어줘
  - 이벤트 페이지 만들자
  - 이벤트페이지 만들어줘
  - 이벤트 페이지 생성해줘
---

# Jodit Event Page

## Purpose
Create a single neutral event-page HTML document for Jodit Editor input/paste. The skill asks a fixed sequence of questions, validates required answers, renders inline-styled HTML, and writes the result to `output/yymmdd_title.html` with numeric suffixes for duplicates.

## Interaction Contract
- When a supported trigger fires, ask the user for the event content first using an ask-question style flow.
- Ask one field at a time in the documented order before generating any HTML.
- Never auto-fill event content from fixtures, examples, or prior sample outputs.
- Fixtures under `fixtures/*.json` are verification-only and MUST NOT be used as default answers in a user-triggered run.
- Do not generate a file with sample titles such as `봄 이벤트` or `봄 혜택 이벤트` unless the user explicitly provided those values.

## When to Activate
- The user asks to make an event page using Korean phrasing that matches one of the supported trigger aliases.
- The requested output is HTML content intended for Jodit Editor.
- The page needs these required sections: event title, event period, event point usage deadline, and notices.

## Trigger Aliases
V1 uses explicit alias matching only. Supported aliases:

- `이벤트 페이지만들자`
- `이벤트 페이지 만들어줘`
- `이벤트 페이지 만들자`
- `이벤트페이지 만들어줘`
- `이벤트 페이지 생성해줘`

Unlisted fuzzy phrasing, English-only requests, and generalized intent matching are out of scope for v1.

## Question Flow
Ask these questions in this exact order.

1. **이벤트 타이틀** — required
2. **이벤트 설명** — optional; if blank, omit the description block under the title
3. **이벤트 기간** — required
4. **이벤트 포인트 사용기한** — required
5. **유의사항** — required; multiline input is allowed and rendered as a bullet list
6. **CTA 문구** — optional
7. **CTA 링크** — optional

CTA rule: if either CTA text or CTA link is blank, omit the CTA section entirely.
If a CTA is rendered, the CTA link must use `http` or `https`.

## Workflow
### 1. Ask Questions Interactively
- After a supported trigger alias is detected, pause generation and collect answers from the user.
- Use the ordered metadata in `.agents/skills/jodit-event-page/skill_runtime.py` as the single source of truth for question order.
- Ask each field separately in ask-question form using the exact labels from `QUESTION_METADATA`.
- For `유의사항`, allow multiline input and split line breaks into bullet items.
- Do not substitute fixture/example values when the user has not answered yet.

### 2. Load Runtime Entry Points
- Use `.agents/skills/jodit-event-page/skill_runtime.py` as the concrete runtime implementation.
- Use `.agents/skills/jodit-event-page/verify_event_page.py` as the reusable verification runner.

### 3. Validate Answers
- Fail before file creation if any required field is blank.
- Trim outer whitespace from all answers before validation.
- Treat whitespace-only title values as invalid.

### 4. Build Safe Output Filename
- Prefix with local system date in `yymmdd` format.
- Preserve Korean characters in the title.
- Remove filesystem-illegal characters: `/ \ : * ? " < > |`
- Collapse repeated spaces and convert internal spaces to underscores.
- If the base filename already exists, write `_2`, `_3`, and so on.

### 5. Render Jodit-Safe HTML
- Fill `.agents/skills/jodit-event-page/template.inline.html`.
- Use inline styles only.
- Do not emit `<style>`, `<script>`, `<button>`, `iframe`, or form-related tags.
- Render CTA as a styled `<a>` element only when both CTA fields are present.
- Add `target="_blank"` and `rel="noopener noreferrer"` to CTA links.

### 6. Write Output
- Ensure `output/` exists before writing.
- Write generated HTML into `output/`.
- Never overwrite an existing file with the same base name.

## Output Contract
- Required-only runs create a non-empty HTML file containing title, period, point usage deadline, and notices.
- Full-input runs additionally contain description and exactly one CTA anchor.
- Notices are rendered as bullet list items.

## Runtime Commands
Use these concrete commands for implementation-time verification:

```bash
python3 .agents/skills/jodit-event-page/verify_event_page.py question-flow
python3 .agents/skills/jodit-event-page/verify_event_page.py filename
python3 .agents/skills/jodit-event-page/verify_event_page.py filename --invalid-title
python3 .agents/skills/jodit-event-page/verify_event_page.py required-only
python3 .agents/skills/jodit-event-page/verify_event_page.py required-only --missing period
python3 .agents/skills/jodit-event-page/verify_event_page.py full-input
python3 .agents/skills/jodit-event-page/verify_event_page.py full-input --duplicate
```

## Verification
- Confirm the skill contract explicitly requires ask-question collection before HTML generation.
- Confirm fixtures/examples are documented as verification-only, not default content.
- Confirm the trigger alias list is documented consistently in both the frontmatter and the `## Trigger Aliases` section.
- Confirm the question order matches the runtime question metadata.
- Confirm the required/optional field flags match the runtime question metadata.
- Confirm generated filenames preserve Korean text and use suffixes for duplicates.
- Confirm optional description and CTA blocks are omitted when their inputs are blank.
- Confirm CTA links are rendered as anchors, not buttons.

## Examples
Example values in this section are documentation samples only. They must never be auto-applied to a real user request.

### Required-only Example
- Trigger: `이벤트 페이지만들자`
- Answers:
  - 이벤트 타이틀: `봄 이벤트`
  - 이벤트 설명: ``
  - 이벤트 기간: `2026.03.25 ~ 2026.03.31`
  - 이벤트 포인트 사용기한: `2026.04.07까지`
  - 유의사항:
    - `이벤트 참여는 기간 내에만 가능합니다.`
    - `포인트는 사용기한 이후 자동 소멸됩니다.`
  - CTA 문구: ``
  - CTA 링크: ``
- Expected output filename: `output/260325_봄_이벤트.html`

### Full-input Example
- Trigger: `이벤트 페이지 만들어줘`
- Answers:
  - 이벤트 타이틀: `봄 혜택 이벤트`
  - 이벤트 설명: `따뜻한 봄 시즌을 맞아 포인트 적립과 사용 혜택을 한번에 확인할 수 있는 이벤트입니다.`
  - 이벤트 기간: `2026.03.25 ~ 2026.04.15`
  - 이벤트 포인트 사용기한: `2026.04.22까지`
  - 유의사항:
    - `ID당 1회만 참여 가능합니다.`
    - `부정 참여가 확인되면 지급 대상에서 제외됩니다.`
    - `포인트는 사용기한 이후 자동 소멸됩니다.`
  - CTA 문구: `이벤트 참여하기`
  - CTA 링크: `https://example.com/events/spring-benefit`
- Expected output filename: `output/260325_봄_혜택_이벤트.html`

## Troubleshooting
### Invalid title
- If the title becomes empty after trimming and illegal-character removal, stop and return a clear validation error.

### Missing required field
- If title, period, point usage deadline, or notices is blank, stop and do not create an output file.

### Duplicate output file
- If `output/yymmdd_title.html` already exists, create `output/yymmdd_title_2.html` instead.

### Output directory missing
- Create `output/` automatically before writing the file.
