from __future__ import annotations

from collections.abc import Callable, Mapping
import argparse
from datetime import date
from html import escape
import importlib.util
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path
from types import ModuleType
from typing import Literal, NoReturn, Protocol, cast

CURRENT_DIR = Path(__file__).resolve().parent


def load_runtime_module() -> ModuleType:
    spec = importlib.util.spec_from_file_location(
        "jodit_event_page_skill_runtime", CURRENT_DIR / "skill_runtime.py"
    )
    if spec is None or spec.loader is None:
        raise RuntimeError("Unable to load skill_runtime.py")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


class HasLabel(Protocol):
    label: str
    required: bool


RUNTIME_MODULE = load_runtime_module()
QUESTION_METADATA = cast(
    tuple[HasLabel, ...], getattr(RUNTIME_MODULE, "QUESTION_METADATA")
)
TRIGGER_ALIASES = cast(tuple[str, ...], getattr(RUNTIME_MODULE, "TRIGGER_ALIASES"))
ValidationError = cast(type[ValueError], getattr(RUNTIME_MODULE, "ValidationError"))


FIXTURES_DIR = CURRENT_DIR / "fixtures"
SKILL_MD_PATH = CURRENT_DIR / "SKILL.md"
VERIFY_ROOT = CURRENT_DIR / ".verify_tmp"
VERIFY_OUTPUT_DIR = VERIFY_ROOT / "output"
FORBIDDEN_PATTERNS = (
    "<style",
    "<script",
    "<button",
    "<iframe",
    "<form",
    "<link",
    "class=",
)
Scenario = Literal["filename", "required-only", "full-input"]
MissingField = Literal["title", "period", "point_usage_deadline", "notices"]


class ContractFailure(AssertionError):
    pass


@dataclass(frozen=True)
class CliArgs:
    scenario: Scenario
    invalid_title: bool
    missing: MissingField | None
    duplicate: bool


def current_date_stamp() -> str:
    return date.today().strftime("%y%m%d")


def fail(message: str) -> NoReturn:
    raise ContractFailure(message)


def clean_verify_state() -> None:
    if VERIFY_ROOT.exists():
        shutil.rmtree(VERIFY_ROOT)


def load_fixture(name: str) -> dict[str, object]:
    runtime_load_json_payload = cast(
        Callable[[str | Path], dict[str, object]],
        getattr(RUNTIME_MODULE, "load_json_payload"),
    )
    payload = cast(object, runtime_load_json_payload(FIXTURES_DIR / name))
    if not isinstance(payload, dict):
        fail(f"Fixture payload must be a JSON object: {name}")
    payload_dict = cast(dict[object, object], payload)
    normalized_payload: dict[str, object] = {}
    for key, value in payload_dict.items():
        normalized_payload[str(key)] = value
    return normalized_payload


def require_str(value: object, field_name: str) -> str:
    if not isinstance(value, str):
        fail(f"Fixture field must be a string: {field_name}")
    return value


def require_bool(value: object, field_name: str) -> bool:
    if not isinstance(value, bool):
        fail(f"Argument must be a bool: {field_name}")
    return value


def require_str_list(value: object, field_name: str) -> list[str]:
    if not isinstance(value, list):
        fail(f"Fixture field must be a list of strings: {field_name}")
    items = cast(list[object], value)
    result: list[str] = []
    for item in items:
        if not isinstance(item, str):
            fail(f"Fixture field must be a list of strings: {field_name}")
        result.append(item)
    return result


def create_event_page(payload: Mapping[str, object], output_dir: Path) -> Path:
    runtime_write_event_page = cast(
        Callable[..., Path], getattr(RUNTIME_MODULE, "write_event_page")
    )
    return runtime_write_event_page(payload, output_dir=output_dir)


def read_skill_markdown() -> str:
    return SKILL_MD_PATH.read_text(encoding="utf-8")


def parse_trigger_aliases(skill_markdown: str) -> list[str]:
    aliases: list[str] = []
    in_section = False
    for line in skill_markdown.splitlines():
        if line.strip() == "## Trigger Aliases":
            in_section = True
            continue
        if in_section and line.startswith("## "):
            break
        stripped = line.strip()
        if in_section and stripped.startswith("- `") and stripped.endswith("`"):
            aliases.append(stripped[3:-1])
    return aliases


def parse_frontmatter_triggers(skill_markdown: str) -> list[str]:
    lines = skill_markdown.splitlines()
    if len(lines) < 3 or lines[0].strip() != "---":
        fail("SKILL.md is missing YAML frontmatter.")

    aliases: list[str] = []
    in_triggers = False
    for line in lines[1:]:
        stripped = line.strip()
        if stripped == "---":
            break
        if stripped == "triggers:":
            in_triggers = True
            continue
        if in_triggers:
            if line.startswith("  - "):
                aliases.append(line.removeprefix("  - ").strip())
            elif stripped:
                break
    return aliases


def parse_question_labels(skill_markdown: str) -> list[str]:
    labels: list[str] = []
    in_section = False
    for line in skill_markdown.splitlines():
        if line.strip() == "## Question Flow":
            in_section = True
            continue
        if in_section and line.startswith("## "):
            break
        stripped = line.strip()
        if in_section and stripped[:1].isdigit() and "**" in stripped:
            first = stripped.find("**")
            second = stripped.find("**", first + 2)
            if first != -1 and second != -1:
                labels.append(stripped[first + 2 : second])
    return labels


def parse_question_required_flags(skill_markdown: str) -> list[bool]:
    flags: list[bool] = []
    in_section = False
    for line in skill_markdown.splitlines():
        if line.strip() == "## Question Flow":
            in_section = True
            continue
        if in_section and line.startswith("## "):
            break
        stripped = line.strip().lower()
        if in_section and stripped[:1].isdigit() and "**" in stripped:
            if "— required" in stripped:
                flags.append(True)
            elif "— optional" in stripped:
                flags.append(False)
    return flags


def assert_skill_metadata_matches_documentation() -> None:
    skill_markdown = read_skill_markdown()
    frontmatter_aliases = parse_frontmatter_triggers(skill_markdown)
    documented_aliases = parse_trigger_aliases(skill_markdown)
    documented_questions = parse_question_labels(skill_markdown)
    documented_required_flags = parse_question_required_flags(skill_markdown)

    if frontmatter_aliases != list(TRIGGER_ALIASES):
        fail(
            f"Frontmatter triggers do not match runtime aliases: {frontmatter_aliases!r} != {list(TRIGGER_ALIASES)!r}"
        )
    if documented_aliases != list(TRIGGER_ALIASES):
        fail(
            f"Trigger aliases do not match SKILL.md: {documented_aliases!r} != {list(TRIGGER_ALIASES)!r}"
        )

    runtime_questions = [item.label for item in QUESTION_METADATA]
    if documented_questions != runtime_questions:
        fail(
            f"Question order does not match SKILL.md: {documented_questions!r} != {runtime_questions!r}"
        )
    runtime_required_flags = [item.required for item in QUESTION_METADATA]
    if documented_required_flags != runtime_required_flags:
        fail(
            f"Required flags do not match SKILL.md: {documented_required_flags!r} != {runtime_required_flags!r}"
        )


def read_created_html(path: Path) -> str:
    if not path.exists():
        fail(f"Expected output file to exist: {path}")
    return path.read_text(encoding="utf-8")


def assert_no_forbidden_tags(html: str) -> None:
    lowered = html.lower()
    for pattern in FORBIDDEN_PATTERNS:
        if pattern in lowered:
            fail(f"Generated HTML contains forbidden fragment: {pattern}")


def assert_contains(html: str, text: str) -> None:
    if text not in html:
        fail(f"Expected HTML to contain: {text}")


def assert_not_contains(html: str, text: str) -> None:
    if text in html:
        fail(f"Expected HTML not to contain: {text}")


def assert_contains_escaped(html: str, text: str) -> None:
    assert_contains(html, escape(text, quote=True))


def assert_section_order(html: str, sections: list[str]) -> None:
    last_index = -1
    for section in sections:
        current_index = html.find(section)
        if current_index == -1:
            fail(f"Expected section label not found: {section}")
        if current_index <= last_index:
            fail(f"Section order mismatch at: {section}")
        last_index = current_index


def assert_contract_order(
    html: str,
    *,
    title: str,
    description: str | None = None,
    cta_text: str | None = None,
) -> None:
    markers = [title]
    if description:
        markers.append(description)
    markers.extend(["이벤트 안내", "이벤트 기간", "포인트 사용기한", "유의사항"])
    if cta_text:
        markers.append(cta_text)
    assert_section_order(html, markers)


def expect_validation_error(payload: dict[str, object], expected_message: str) -> str:
    clean_verify_state()
    try:
        _ = create_event_page(payload, VERIFY_OUTPUT_DIR)
    except ValidationError as exc:
        if expected_message not in str(exc):
            fail(f"Validation error mismatch: {exc!r}")
        if VERIFY_ROOT.exists() and list(VERIFY_ROOT.rglob("*.html")):
            fail("Validation failure must not create HTML output files.")
        print(f"EXPECTED_VALIDATION: {exc}")
        return str(exc)
    else:
        fail("Expected ValidationError but file creation succeeded.")


def verify_filename_case(invalid_title: bool) -> list[Path]:
    payload = load_fixture("required-only.json")
    if invalid_title:
        payload["title"] = '  / \\ : * ? " < > |  '
        _ = expect_validation_error(payload, "title is invalid after sanitization")
        return []

    clean_verify_state()
    payload["title"] = "봄 이벤트"
    created_path = create_event_page(payload, VERIFY_OUTPUT_DIR)
    duplicate_path = create_event_page(payload, VERIFY_OUTPUT_DIR)
    expected_name = f"{current_date_stamp()}_봄_이벤트.html"
    if created_path.name != expected_name:
        fail(f"Filename mismatch: {created_path.name!r} != {expected_name!r}")
    if created_path.parent != VERIFY_OUTPUT_DIR:
        fail(f"Output directory mismatch: {created_path.parent} != {VERIFY_OUTPUT_DIR}")
    expected_duplicate_name = f"{current_date_stamp()}_봄_이벤트_2.html"
    if duplicate_path.name != expected_duplicate_name:
        fail(
            f"Duplicate filename mismatch: {duplicate_path.name!r} != {expected_duplicate_name!r}"
        )

    clean_verify_state()
    reserved_title_payload = load_fixture("required-only.json")
    reserved_title_payload["title"] = "봄/이벤트?"
    reserved_title_path = create_event_page(reserved_title_payload, VERIFY_OUTPUT_DIR)
    expected_reserved_name = f"{current_date_stamp()}_봄이벤트.html"
    if reserved_title_path.name != expected_reserved_name:
        fail(
            f"Reserved-char filename mismatch: {reserved_title_path.name!r} != {expected_reserved_name!r}"
        )
    return [created_path, duplicate_path, reserved_title_path]


def verify_required_only_case(missing_field: MissingField | None) -> list[Path]:
    payload = load_fixture("required-only.json")

    if missing_field is not None:
        if missing_field == "notices":
            payload[missing_field] = []
        else:
            payload[missing_field] = "   "
        _ = expect_validation_error(payload, f"{missing_field} is required")
        return []

    clean_verify_state()
    created_path = create_event_page(payload, VERIFY_OUTPUT_DIR)
    html = read_created_html(created_path)
    title = require_str(payload["title"], "title")
    period = require_str(payload["period"], "period")
    deadline = require_str(payload["point_usage_deadline"], "point_usage_deadline")
    notices = require_str_list(payload["notices"], "notices")

    assert_contains_escaped(html, title)
    assert_contains_escaped(html, period)
    assert_contains_escaped(html, deadline)
    for notice in notices:
        assert_contains_escaped(html, notice)
    if html.count("<li ") != len(notices):
        fail("Notices must render as one <li> per notice.")
    assert_contract_order(html, title=title)
    assert_not_contains(html, '<p style="margin: 0; font-size: 16px; color: #666666;">')
    if "<a " in html:
        fail("Required-only output must omit CTA block.")
    assert_no_forbidden_tags(html)
    return [created_path]


def verify_full_input_case(duplicate: bool) -> list[Path]:
    payload = load_fixture("full-input.json")
    title = require_str(payload["title"], "title")
    description = require_str(payload["description"], "description")
    period = require_str(payload["period"], "period")
    deadline = require_str(payload["point_usage_deadline"], "point_usage_deadline")
    notices = require_str_list(payload["notices"], "notices")
    cta_link = require_str(payload["cta_link"], "cta_link")

    invalid_link_payload = dict(payload)
    invalid_link_payload["cta_link"] = "ftp://example.com/not-allowed"
    _ = expect_validation_error(invalid_link_payload, "cta_link must use http or https")

    text_only_payload = dict(payload)
    text_only_payload["cta_link"] = ""
    clean_verify_state()
    text_only_path = create_event_page(text_only_payload, VERIFY_OUTPUT_DIR)
    text_only_html = read_created_html(text_only_path)
    if "<a " in text_only_html:
        fail("CTA must be omitted when link is blank.")

    link_only_payload = dict(payload)
    link_only_payload["cta_text"] = ""
    clean_verify_state()
    link_only_path = create_event_page(link_only_payload, VERIFY_OUTPUT_DIR)
    link_only_html = read_created_html(link_only_path)
    if "<a " in link_only_html:
        fail("CTA must be omitted when text is blank.")

    escaped_payload = dict(payload)
    escaped_payload["title"] = "<봄 & 이벤트>"
    escaped_payload["description"] = "혜택 <strong>설명</strong> & 안내"
    escaped_payload["notices"] = ["첫째 <공지>", '둘째 "주의" & 확인']
    clean_verify_state()
    escaped_path = create_event_page(escaped_payload, VERIFY_OUTPUT_DIR)
    escaped_html = read_created_html(escaped_path)
    assert_contains_escaped(escaped_html, "<봄 & 이벤트>")
    assert_contains_escaped(escaped_html, "혜택 <strong>설명</strong> & 안내")
    assert_contains_escaped(escaped_html, "첫째 <공지>")
    assert_contains_escaped(escaped_html, '둘째 "주의" & 확인')
    assert_not_contains(escaped_html, "<strong>설명</strong>")

    clean_verify_state()

    first_path = create_event_page(payload, VERIFY_OUTPUT_DIR)
    first_html = read_created_html(first_path)
    assert_contains_escaped(first_html, title)
    assert_contains_escaped(first_html, description)
    assert_contains_escaped(first_html, period)
    assert_contains_escaped(first_html, deadline)
    for notice in notices:
        assert_contains_escaped(first_html, notice)
    if first_html.count("<li ") != len(notices):
        fail("Full-input notices must render as one <li> per notice.")
    if first_html.count("<a ") != 1:
        fail("Full-input output must contain exactly one CTA anchor.")
    assert_contains(first_html, f'href="{cta_link}"')
    assert_contains(first_html, 'target="_blank"')
    assert_contains(first_html, 'rel="noopener noreferrer"')
    cta_text = require_str(payload["cta_text"], "cta_text")
    assert_contract_order(
        first_html, title=title, description=description, cta_text=cta_text
    )
    assert_not_contains(first_html.lower(), "<button")
    assert_no_forbidden_tags(first_html)

    if not duplicate:
        return [first_path]

    clean_verify_state()
    first_path = create_event_page(payload, VERIFY_OUTPUT_DIR)
    second_path = create_event_page(payload, VERIFY_OUTPUT_DIR)
    base_name = f"{current_date_stamp()}_봄_혜택_이벤트"
    if first_path.name != f"{base_name}.html":
        fail(f"Unexpected first duplicate filename: {first_path.name}")
    if second_path.name != f"{base_name}_2.html":
        fail(f"Unexpected duplicate filename: {second_path.name}")
    if read_created_html(first_path) != read_created_html(second_path):
        fail("Duplicate file contents should remain identical.")
    return [first_path, second_path]


def parse_args(argv: list[str]) -> CliArgs:
    parser = argparse.ArgumentParser(
        description="Verify Jodit event-page runtime contract."
    )
    _ = parser.add_argument(
        "scenario", choices=("filename", "required-only", "full-input")
    )
    _ = parser.add_argument("--invalid-title", action="store_true")
    _ = parser.add_argument(
        "--missing", choices=("title", "period", "point_usage_deadline", "notices")
    )
    _ = parser.add_argument("--duplicate", action="store_true")
    parsed = cast(dict[str, object], vars(parser.parse_args(argv)))

    scenario_value = require_str(parsed.get("scenario"), "scenario")
    if scenario_value not in ("filename", "required-only", "full-input"):
        fail(f"Unsupported scenario: {scenario_value}")

    missing_value = None
    missing_raw = parsed.get("missing")
    if missing_raw is not None:
        missing_text = require_str(missing_raw, "missing")
        if missing_text not in ("title", "period", "point_usage_deadline", "notices"):
            fail(f"Unsupported missing field: {missing_text}")
        missing_value = missing_text

    return CliArgs(
        scenario=scenario_value,
        invalid_title=require_bool(parsed.get("invalid_title"), "invalid_title"),
        missing=missing_value,
        duplicate=require_bool(parsed.get("duplicate"), "duplicate"),
    )


def run(args: CliArgs) -> list[Path]:
    assert_skill_metadata_matches_documentation()

    if args.scenario != "filename" and args.invalid_title:
        fail("--invalid-title is only supported with the filename scenario.")
    if args.scenario != "required-only" and args.missing is not None:
        fail("--missing is only supported with the required-only scenario.")
    if args.scenario != "full-input" and args.duplicate:
        fail("--duplicate is only supported with the full-input scenario.")

    if args.scenario == "filename":
        return verify_filename_case(invalid_title=args.invalid_title)
    if args.scenario == "required-only":
        return verify_required_only_case(missing_field=args.missing)
    return verify_full_input_case(duplicate=args.duplicate)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    try:
        created_paths = run(args)
    except (ContractFailure, ValidationError) as exc:
        print(f"FAIL: {exc}", file=sys.stderr)
        return 1

    for path in created_paths:
        print(path)
    print("OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
