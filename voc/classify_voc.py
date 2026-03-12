"""VOC classification script using GitHub Models API."""
import argparse
import base64
import json
import os
import sys
import time
import urllib.request
import urllib.error
from io import BytesIO
from pathlib import Path
from typing import Any

import openpyxl


VALID_CATEGORIES = {
    "반응_문의",
    "반응_교환",
    "반응_구해요",
    "반응_정보",
    "반응_기타",
}
TARGET_SHEET = "2026 list"
TARGET_MARKER = "반응_제품반응"
TARGET_COLUMN = 7
TITLE_COLUMN = 10
LINK_COLUMN = 11
START_ROW = 5
END_ROW = 220

EXCHANGE_KEYWORDS = ["[교환]", "교환원해", "교환 구해", "교환하실", "교환해주", "교환 원해"]
SEEKING_KEYWORDS = [
    "[구해요]", "구해요", "구해용", "구합니다", "구입하실분",
    "구해봐", "구해보아요", "구매원해", "구매합니다", "구함",
]

REACTION_SHEETS = ["BP_반응", "PK_반응"]


def preprocess_title(text) -> str | None:
    if text is None:
        return None
    cleaned = str(text).replace("\n", " ").strip()
    if not cleaned:
        return None
    return " ".join(cleaned.split())


def build_classification_prompt(title) -> str:
    categories = "\n".join(f"- {item}" for item in sorted(VALID_CATEGORIES))
    return (
        "당신은 VOC 제목을 아래 5개 카테고리 중 하나로만 분류하는 분류기입니다.\n"
        "카테고리 목록:\n"
        f"{categories}\n\n"
        "카테고리 분류 기준:\n"
        "- 반응_교환: [교환] 태그, 교환 요청, 사이즈 교환, 교환 구해요\n"
        "- 반응_구해요: [구해요] 태그, 구해요, 구합니다, 구해용, 구입하실분, 판매처 문의\n"
        "- 반응_문의: 궁금, 질문, 어디서 사나요, 사이즈 팁, 착샷 보여주세요, ~인가요?\n"
        "- 반응_정보: 가격, 성분, 할인, 세일, 제품 스펙 공유\n"
        "- 반응_기타: 위 4개에 해당하지 않는 단순 감상, 후기, 추천, 비추천\n\n"
        "few-shot 예시:\n"
        '입력: "궁금합니다" -> 출력: 반응_문의\n'
        '입력: "어디서 사나요" -> 출력: 반응_문의\n'
        '입력: "판매처 아시는분" -> 출력: 반응_구해요\n'
        '입력: "사이즈 팁이요" -> 출력: 반응_문의\n'
        '입력: "착샷 좀 보여주세요" -> 출력: 반응_문의\n'
        '입력: "[교환] 교환해주세요" -> 출력: 반응_교환\n'
        '입력: "스피드캣 교환원해요" -> 출력: 반응_교환\n'
        '입력: "[교환] 블랙150(저) > 140(맘님)" -> 출력: 반응_교환\n'
        '입력: "구합니다" -> 출력: 반응_구해요\n'
        '입력: "구해요" -> 출력: 반응_구해요\n'
        '입력: "[구해요] 실버 130 구해용" -> 출력: 반응_구해요\n'
        '입력: "구입하실분들" -> 출력: 반응_구해요\n'
        '입력: "가격이 얼마인가요" -> 출력: 반응_정보\n'
        '입력: "성분 정보" -> 출력: 반응_정보\n'
        '입력: "2만원대부터! 할인도 있네요" -> 출력: 반응_정보\n'
        '입력: "왕비추천" -> 출력: 반응_기타\n'
        '입력: "베베드피노 몬치치가방" -> 출력: 반응_기타\n\n'
        "중요: [교환] 또는 [구해요] 태그가 있으면 해당 카테고리로 분류하세요.\n"
        "반드시 카테고리 텍스트 하나만 출력하세요.\n"
        f"분류 대상 제목: {title}"
    )


def validate_result(text) -> str:
    value = (text or "").strip()
    if value in VALID_CATEGORIES:
        return value
    for cat in VALID_CATEGORIES:
        if cat in value:
            return cat
    return "반응_기타"


def _request_json(url: str, body: dict, api_key: str) -> dict:
    data = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=data,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
            "User-Agent": "voc-classifier/1.0",
        },
        method="POST",
    )
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read().decode("utf-8"))


class GithubModelsAdapter:
    endpoint = "https://models.github.ai/inference/chat/completions"

    def __init__(self, api_key: str, model: str = "openai/gpt-4.1"):
        self.api_key = api_key
        self.model = model

    def build_request_body(self, title):
        return {
            "model": self.model,
            "messages": [
                {
                    "role": "system",
                    "content": build_classification_prompt(title),
                },
                {
                    "role": "user",
                    "content": str(title),
                },
            ],
        }

    def parse_response(self, resp):
        return resp["choices"][0]["message"]["content"]


def _classify_with_adapter(adapter, title) -> str | None:
    """Returns None on API failure (HTTP/network), triggering fallback."""
    backoff = 1
    for attempt in range(3):
        try:
            body = adapter.build_request_body(title)
            response = _request_json(adapter.endpoint, body, adapter.api_key)
            return validate_result(adapter.parse_response(response))
        except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError):
            if attempt == 2:
                return None
            time.sleep(backoff)
            backoff *= 2
        except (KeyError, IndexError, TypeError, ValueError):
            return "반응_기타"
        finally:
            time.sleep(0.5)
    return None


def rule_classify(title: str) -> str | None:
    lower = title.lower()
    for kw in EXCHANGE_KEYWORDS:
        if kw.lower() in lower:
            return "반응_교환"
    for kw in SEEKING_KEYWORDS:
        if kw.lower() in lower:
            return "반응_구해요"
    return None


def _extract_link_text(cell) -> str:
    if cell.hyperlink and getattr(cell.hyperlink, "target", None):
        return str(cell.hyperlink.target)
    return str(cell.value or "")


def _image_anchor_col_row(img) -> tuple[int | None, int | None]:
    anchor = getattr(img, "anchor", None)
    marker = getattr(anchor, "_from", None)
    if marker is None:
        return None, None
    col = getattr(marker, "col", None)
    row = getattr(marker, "row", None)
    if col is None or row is None:
        return None, None
    return int(col) + 1, int(row) + 1


def _extract_image_bytes(img) -> bytes | None:
    if hasattr(img, "_data"):
        try:
            data = img._data()
            if isinstance(data, bytes):
                return data
        except Exception:
            pass
    ref = getattr(img, "ref", None)
    if isinstance(ref, bytes):
        return ref
    ref_reader = getattr(ref, "read", None)
    if callable(ref_reader):
        try:
            data = ref_reader()
            if isinstance(data, bytes):
                return data
            return None
        except Exception:
            return None
    return None


def build_image_index(wb) -> dict[str, list[bytes]]:
    index: dict[str, list[bytes]] = {}
    for sheet_name in REACTION_SHEETS:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]

        card_starts: list[int] = []
        for col in range(1, ws.max_column + 1):
            value = ws.cell(2, col).value
            if value is not None and "Voice Of Customer" in str(value):
                card_starts.append(col)
        if not card_starts:
            continue

        images = list(getattr(ws, "_images", []))
        for idx, card_start in enumerate(card_starts):
            link_col = card_start + 4
            link = _extract_link_text(ws.cell(4, link_col))
            if not link:
                continue

            next_card_start = card_starts[idx + 1] if idx + 1 < len(card_starts) else None
            body_images: list[tuple[int, int, bytes]] = []
            for img in images:
                col, row = _image_anchor_col_row(img)
                if col is None or row is None:
                    continue
                if col < card_start:
                    continue
                if next_card_start is not None and col >= next_card_start:
                    continue
                if row <= 6:
                    continue
                data = _extract_image_bytes(img)
                if data:
                    body_images.append((row, col, data))

            if body_images:
                body_images.sort(key=lambda item: (item[0], item[1]))
                index[link] = [item[2] for item in body_images]
    return index


def _build_vision_request_body(model: str, image_bytes_list: list[bytes]) -> dict[str, Any]:
    images = [
        {
            "type": "image_url",
            "image_url": {
                "url": f"data:image/jpeg;base64,{base64.b64encode(img).decode()}"
            },
        }
        for img in image_bytes_list
    ]
    return {
        "model": model,
        "messages": [
            {
                "role": "system",
                "content": (
                    "이미지에서 텍스트를 읽고 VOC를 분류하세요. "
                    "카테고리: 반응_문의, 반응_교환, 반응_구해요, 반응_정보, 반응_기타. "
                    "카테고리만 출력."
                ),
            },
            {
                "role": "user",
                "content": [
                    *images,
                    {
                        "type": "text",
                        "text": "이 게시글 본문 이미지를 읽고 VOC 카테고리를 분류하세요.",
                    },
                ],
            },
        ],
    }


def _classify_with_image_model(image_bytes_list: list[bytes], api_key: str, model: str) -> str | None:
    adapter = GithubModelsAdapter(api_key=api_key, model=model)
    backoff = 1
    for attempt in range(3):
        try:
            body = _build_vision_request_body(model, image_bytes_list)
            response = _request_json(adapter.endpoint, body, adapter.api_key)
            content = response["choices"][0]["message"]["content"]
            return validate_result(content)
        except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError):
            if attempt == 2:
                return None
            time.sleep(backoff)
            backoff *= 2
        except (KeyError, IndexError, TypeError, ValueError):
            return "반응_기타"
        finally:
            time.sleep(0.5)
    return None


def classify_with_image(image_bytes_list: list[bytes], api_key: str) -> str | None:
    if not image_bytes_list:
        return None
    primary = _classify_with_image_model(image_bytes_list, api_key, "openai/gpt-4.1")
    if primary is not None:
        return primary
    return _classify_with_image_model(image_bytes_list, api_key, "openai/gpt-4o")


def _classify_item_with_method(
    title,
    link: str | None,
    image_index: dict[str, list[bytes]] | None,
    api_key,
) -> tuple[str, str]:
    cleaned = preprocess_title(title)
    if cleaned is None:
        return "반응_기타", "empty"

    rule_result = rule_classify(cleaned)
    if rule_result is not None:
        return rule_result, "rule"

    primary_result = _classify_with_adapter(
        GithubModelsAdapter(api_key=api_key, model="openai/gpt-4.1"), cleaned
    )
    if primary_result is not None and primary_result != "반응_기타":
        return primary_result, "llm"

    if primary_result == "반응_기타" and link and image_index and image_index.get(link):
        image_result = classify_with_image(image_index[link], api_key)
        if image_result is not None:
            return image_result, "image"

    if primary_result == "반응_기타":
        return "반응_기타", "llm"

    fallback_result = _classify_with_adapter(
        GithubModelsAdapter(api_key=api_key, model="openai/gpt-4o"), cleaned
    )
    if fallback_result is not None:
        return fallback_result, "fallback"
    return "반응_기타", "fallback"


def classify_item(
    title,
    link: str | None,
    image_index: dict[str, list[bytes]] | None,
    api_key,
) -> str:
    category, _ = _classify_item_with_method(title, link, image_index, api_key)
    return category


def classify_title(title, api_key) -> str:
    return classify_item(title, link=None, image_index=None, api_key=api_key)


def read_target_rows(wb) -> list[dict]:
    ws = wb[TARGET_SHEET]
    targets = []
    for row in range(START_ROW, END_ROW + 1):
        if ws.cell(row, TARGET_COLUMN).value != TARGET_MARKER:
            continue
        title = preprocess_title(ws.cell(row, TITLE_COLUMN).value)
        if title is None:
            continue
        cell = ws.cell(row, LINK_COLUMN)
        link_val = _extract_link_text(cell)
        targets.append({"row": row, "title": title, "link": link_val})
    return targets


def write_classification_result(wb, row, category):
    ws = wb[TARGET_SHEET]
    ws.cell(row, TARGET_COLUMN).value = validate_result(category)


def _prebuffer_workbook_images(wb):
    for ws in wb.worksheets:
        for img in getattr(ws, "_images", []):
            try:
                ref = img.ref
                if hasattr(ref, "read"):
                    ref.seek(0)
                    data = ref.read()
                    img.ref = BytesIO(data)
                    img._cached_data = data
            except Exception:
                pass


def _refresh_image_refs(wb):
    for ws in wb.worksheets:
        for img in getattr(ws, "_images", []):
            cached = getattr(img, "_cached_data", None)
            if cached:
                img.ref = BytesIO(cached)


def run_classification(wb, api_key, output_path):
    _prebuffer_workbook_images(wb)
    image_index = build_image_index(wb)
    targets = read_target_rows(wb)
    summary = {category: 0 for category in sorted(VALID_CATEGORIES)}

    total = len(targets)
    for idx, item in enumerate(targets, start=1):
        category, method = _classify_item_with_method(
            item["title"],
            item.get("link"),
            image_index,
            api_key,
        )
        write_classification_result(wb, item["row"], category)
        summary[category] += 1
        print(f"[{idx}/{total}] row={item['row']} ({method}) -> {category}", file=sys.stderr)

    if output_path:
        _refresh_image_refs(wb)
        wb.save(output_path)

    print("분류 완료 요약")
    for key in sorted(summary):
        print(f"{key}: {summary[key]}")
    return summary


def classify_workbook(wb, api_key, output_path):
    return run_classification(wb, api_key=api_key, output_path=output_path)


INPUT_DIR = Path("input")


def _find_input_xlsx() -> Path | None:
    """Auto-detect a single xlsx file from the default input directory."""
    if not INPUT_DIR.is_dir():
        return None
    xlsx_files = sorted(INPUT_DIR.glob("*.xlsx"))
    if len(xlsx_files) == 1:
        return xlsx_files[0]
    return None


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="VOC workbook classifier")
    parser.add_argument("input", nargs="?", help="Input xlsx path (default: auto-detect from input/)")
    parser.add_argument("--output-dir", default="output/", help="Output directory")
    parser.add_argument("--dry-run", action="store_true", help="Skip API call")
    return parser


def main(argv=None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)

    if not args.input:
        auto = _find_input_xlsx()
        if auto:
            args.input = str(auto)
            print(f"Auto-detected input: {auto}", file=sys.stderr)
        else:
            parser.print_help()
            return 0

    api_key = os.getenv("GITHUB_TOKEN")
    if not args.dry_run and not api_key:
        print("Error: GITHUB_TOKEN is required", file=sys.stderr)
        return 1

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Error: 파일을 찾을 수 없습니다: {input_path}", file=sys.stderr)
        return 1

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / input_path.name

    wb = openpyxl.load_workbook(input_path)
    if args.dry_run:
        targets = read_target_rows(wb)
        print(f"dry-run: {len(targets)} target rows")
        return 0

    run_classification(wb, api_key=api_key, output_path=output_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
