from __future__ import annotations

from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass
from datetime import date
from html import escape
from pathlib import Path
from typing import cast
from urllib.parse import urlparse
import json
import re


SKILL_DIR = Path(__file__).resolve().parent
REPO_ROOT = SKILL_DIR.parents[2]
TEMPLATE_PATH = SKILL_DIR / "template.inline.html"
DEFAULT_OUTPUT_DIR = REPO_ROOT / "output"
INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|]')
WHITESPACE_RE = re.compile(r"\s+")
IF_BLOCK_RE = re.compile(r"{%\s*if\s+(.+?)\s*%}(.*?){%\s*endif\s*%}", re.DOTALL)
PLACEHOLDER_RE = re.compile(r"{{\s*(\w+)\s*}}")


@dataclass(frozen=True)
class QuestionMetadata:
    order: int
    key: str
    label: str
    required: bool
    prompt: str
    multiline: bool = False


@dataclass(frozen=True)
class EventPageInput:
    title: str
    description: str
    period: str
    point_usage_deadline: str
    notices: tuple[str, ...]
    cta_text: str
    cta_link: str


class ValidationError(ValueError):
    pass


TRIGGER_ALIASES: tuple[str, ...] = (
    "이벤트 페이지만들자",
    "이벤트 페이지 만들어줘",
    "이벤트 페이지 만들자",
    "이벤트페이지 만들어줘",
    "이벤트 페이지 생성해줘",
)

QUESTION_METADATA: tuple[QuestionMetadata, ...] = (
    QuestionMetadata(
        order=1,
        key="title",
        label="이벤트 타이틀",
        required=True,
        prompt="이벤트 타이틀을 입력해 주세요.",
    ),
    QuestionMetadata(
        order=2,
        key="description",
        label="이벤트 설명",
        required=False,
        prompt="이벤트 설명을 입력해 주세요. 없으면 비워 두셔도 됩니다.",
    ),
    QuestionMetadata(
        order=3,
        key="period",
        label="이벤트 기간",
        required=True,
        prompt="이벤트 기간을 입력해 주세요.",
    ),
    QuestionMetadata(
        order=4,
        key="point_usage_deadline",
        label="이벤트 포인트 사용기한",
        required=True,
        prompt="이벤트 포인트 사용기한을 입력해 주세요.",
    ),
    QuestionMetadata(
        order=5,
        key="notices",
        label="유의사항",
        required=True,
        prompt="유의사항을 입력해 주세요. 여러 줄이면 줄바꿈마다 항목으로 처리합니다.",
        multiline=True,
    ),
    QuestionMetadata(
        order=6,
        key="cta_text",
        label="CTA 문구",
        required=False,
        prompt="CTA 문구를 입력해 주세요. 없으면 비워 두셔도 됩니다.",
    ),
    QuestionMetadata(
        order=7,
        key="cta_link",
        label="CTA 링크",
        required=False,
        prompt="CTA 링크를 입력해 주세요. 없으면 비워 두셔도 됩니다.",
    ),
)

REQUIRED_FIELDS: tuple[str, ...] = (
    "title",
    "period",
    "point_usage_deadline",
    "notices",
)
RAW_TEMPLATE_KEYS = {"notices_list_items"}
QuestionAnswerProvider = Callable[[QuestionMetadata], object]


def load_json_payload(json_path: str | Path) -> dict[str, object]:
    with Path(json_path).open("r", encoding="utf-8") as handle:
        payload = cast(object, json.load(handle))
    if not isinstance(payload, dict):
        raise ValidationError("payload must be a JSON object.")
    payload_dict = cast(dict[object, object], payload)
    normalized_payload: dict[str, object] = {}
    for key, value in payload_dict.items():
        normalized_payload[str(key)] = value
    return normalized_payload


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_notices(value: object) -> tuple[str, ...]:
    if value is None:
        return ()
    if isinstance(value, str):
        raw_items = value.splitlines()
    elif isinstance(value, Sequence) and not isinstance(value, (bytes, bytearray, str)):
        raw_items = [str(item) for item in value]
    else:
        raise ValidationError("notices must be a list or multiline string.")

    notices = tuple(item.strip() for item in raw_items if str(item).strip())
    return notices


def sanitize_title_for_filename(title: str) -> str:
    without_illegal_chars = INVALID_FILENAME_CHARS.sub("", title)
    collapsed_spaces = WHITESPACE_RE.sub(" ", without_illegal_chars).strip()
    return collapsed_spaces.replace(" ", "_")


def validate_cta_link(cta_link: str) -> None:
    if not cta_link:
        return
    parsed = urlparse(cta_link)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValidationError("cta_link must use http or https.")


def prepare_event_page_input(payload: Mapping[str, object]) -> EventPageInput:
    title = normalize_text(payload.get("title"))
    description = normalize_text(payload.get("description"))
    period = normalize_text(payload.get("period"))
    point_usage_deadline = normalize_text(payload.get("point_usage_deadline"))
    notices = normalize_notices(payload.get("notices"))
    cta_text = normalize_text(payload.get("cta_text"))
    cta_link = normalize_text(payload.get("cta_link"))

    if not title:
        raise ValidationError("title is required.")
    if not period:
        raise ValidationError("period is required.")
    if not point_usage_deadline:
        raise ValidationError("point_usage_deadline is required.")
    if not notices:
        raise ValidationError("notices is required.")

    sanitized_title = sanitize_title_for_filename(title)
    if not sanitized_title:
        raise ValidationError("title is invalid after sanitization.")

    if cta_text and cta_link:
        validate_cta_link(cta_link)
    else:
        cta_text = ""
        cta_link = ""

    return EventPageInput(
        title=title,
        description=description,
        period=period,
        point_usage_deadline=point_usage_deadline,
        notices=notices,
        cta_text=cta_text,
        cta_link=cta_link,
    )


def build_question_flow() -> tuple[QuestionMetadata, ...]:
    return QUESTION_METADATA


def build_question_flow_payload() -> tuple[dict[str, object], ...]:
    payload: list[dict[str, object]] = []
    for item in QUESTION_METADATA:
        guidance = "required" if item.required else "optional"
        payload.append(
            {
                "order": item.order,
                "key": item.key,
                "header": item.label,
                "question": item.prompt,
                "required": item.required,
                "multiline": item.multiline,
                "guidance": guidance,
            }
        )
    return tuple(payload)


def collect_event_page_answers(
    answer_provider: QuestionAnswerProvider,
) -> dict[str, object]:
    payload: dict[str, object] = {}
    for item in QUESTION_METADATA:
        answer = answer_provider(item)
        if item.key == "notices":
            payload[item.key] = list(normalize_notices(answer))
            continue
        payload[item.key] = normalize_text(answer)
    return payload


def write_event_page_from_answers(
    answer_provider: QuestionAnswerProvider,
    output_dir: str | Path | None = None,
    template_path: str | Path = TEMPLATE_PATH,
    today: date | None = None,
) -> Path:
    payload = collect_event_page_answers(answer_provider)
    return write_event_page(
        payload, output_dir=output_dir, template_path=template_path, today=today
    )


def build_notices_list_items(notices: Sequence[str]) -> str:
    return "\n".join(
        f'<li style="margin-bottom: 8px;">{escape(notice, quote=True)}</li>'
        for notice in notices
    )


def build_render_context(event_input: EventPageInput) -> dict[str, str]:
    return {
        "title": event_input.title,
        "description": event_input.description,
        "event_period": event_input.period,
        "point_usage_deadline": event_input.point_usage_deadline,
        "notices_list_items": build_notices_list_items(event_input.notices),
        "cta_text": event_input.cta_text,
        "cta_link": event_input.cta_link,
    }


def evaluate_condition(expression: str, context: Mapping[str, str]) -> bool:
    variables = [part.strip() for part in expression.split("and")]
    if not variables:
        return False

    for variable in variables:
        if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", variable):
            raise ValueError(f"Unsupported template condition: {expression}")
        if not context.get(variable, ""):
            return False
    return True


def render_conditionals(template_text: str, context: Mapping[str, str]) -> str:
    rendered = template_text
    while True:
        updated, count = IF_BLOCK_RE.subn(
            lambda match: (
                match.group(2) if evaluate_condition(match.group(1), context) else ""
            ),
            rendered,
        )
        rendered = updated
        if count == 0:
            return rendered


def render_placeholders(template_text: str, context: Mapping[str, str]) -> str:
    def replace(match: re.Match[str]) -> str:
        key = match.group(1)
        value = context.get(key, "")
        if key in RAW_TEMPLATE_KEYS:
            return value
        return escape(value, quote=True)

    return PLACEHOLDER_RE.sub(replace, template_text)


def render_template(template_text: str, context: Mapping[str, str]) -> str:
    with_blocks = render_conditionals(template_text, context)
    return render_placeholders(with_blocks, context)


def read_template(template_path: str | Path = TEMPLATE_PATH) -> str:
    return Path(template_path).read_text(encoding="utf-8")


def current_date_stamp(today: date | None = None) -> str:
    current = today or date.today()
    return current.strftime("%y%m%d")


def ensure_output_dir(output_dir: str | Path | None = None) -> Path:
    destination = Path(output_dir) if output_dir is not None else DEFAULT_OUTPUT_DIR
    destination.mkdir(parents=True, exist_ok=True)
    return destination


def build_output_path(
    title: str,
    output_dir: str | Path | None = None,
    today: date | None = None,
) -> Path:
    destination = ensure_output_dir(output_dir)
    sanitized_title = sanitize_title_for_filename(title)
    if not sanitized_title:
        raise ValidationError("title is invalid after sanitization.")

    prefix = current_date_stamp(today)
    base_name = f"{prefix}_{sanitized_title}"
    candidate = destination / f"{base_name}.html"
    suffix = 2
    while candidate.exists():
        candidate = destination / f"{base_name}_{suffix}.html"
        suffix += 1
    return candidate


def render_event_page_html(
    payload: Mapping[str, object],
    template_path: str | Path = TEMPLATE_PATH,
) -> str:
    event_input = prepare_event_page_input(payload)
    template_text = read_template(template_path)
    context = build_render_context(event_input)
    return render_template(template_text, context)


def write_event_page(
    payload: Mapping[str, object],
    output_dir: str | Path | None = None,
    template_path: str | Path = TEMPLATE_PATH,
    today: date | None = None,
) -> Path:
    event_input = prepare_event_page_input(payload)
    output_path = build_output_path(
        event_input.title, output_dir=output_dir, today=today
    )
    html = render_template(
        read_template(template_path), build_render_context(event_input)
    )
    _ = output_path.write_text(html, encoding="utf-8")
    return output_path


def write_event_page_from_json(
    json_path: str | Path,
    output_dir: str | Path | None = None,
    template_path: str | Path = TEMPLATE_PATH,
    today: date | None = None,
) -> Path:
    payload = load_json_payload(json_path)
    return write_event_page(
        payload, output_dir=output_dir, template_path=template_path, today=today
    )


__all__ = [
    "DEFAULT_OUTPUT_DIR",
    "QUESTION_METADATA",
    "REQUIRED_FIELDS",
    "SKILL_DIR",
    "TEMPLATE_PATH",
    "TRIGGER_ALIASES",
    "ValidationError",
    "build_question_flow",
    "build_question_flow_payload",
    "build_output_path",
    "build_render_context",
    "build_notices_list_items",
    "collect_event_page_answers",
    "current_date_stamp",
    "ensure_output_dir",
    "load_json_payload",
    "prepare_event_page_input",
    "render_event_page_html",
    "sanitize_title_for_filename",
    "write_event_page",
    "write_event_page_from_answers",
    "write_event_page_from_json",
]
