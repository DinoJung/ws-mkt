"""RED-phase tests for VOC classification script."""

# pyright: reportMissingImports=false, reportAttributeAccessIssue=false, reportOptionalMemberAccess=false
import json
from pathlib import Path
import re
import tempfile
import unittest
from typing import Any
from unittest.mock import patch, MagicMock
from io import BytesIO
from zipfile import ZipFile

import openpyxl
import pytest
from openpyxl.styles import Border, Side


VALID_CATEGORIES = {
    "반응_문의",
    "반응_교환",
    "반응_구해요",
    "반응_정보",
    "반응_기타",
}


def _build_base_workbook():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "2026 list"
    ws.cell(4, 7).value = "주요 내용"
    ws.cell(4, 10).value = "제목"
    return wb, ws


def _add_card(ws, start_col, link, category="반응_제품반응"):
    ws.cell(2, start_col).value = "Voice Of Customer"
    ws.cell(4, start_col).value = "사이트"
    ws.cell(4, start_col + 3).value = "링크"
    ws.cell(4, start_col + 4).value = link
    ws.cell(5, start_col).value = "주요 이슈"
    ws.cell(5, start_col + 1).value = category


def _mock_http_response(payload):
    response = MagicMock()
    response.__enter__.return_value.read.return_value = json.dumps(payload).encode(
        "utf-8"
    )
    return response


def _http_error(url):
    http_error_cls = __import__("urllib.error", fromlist=["HTTPError"]).HTTPError
    return http_error_cls(url, 500, "mock error", hdrs=None, fp=None)


def _pipeline_callable(module):
    return getattr(module, "run_classification", None) or getattr(
        module, "classify_workbook", None
    )


def _real_input_workbook() -> Path:
    path = Path("input/2026더캐리VOC(리포트파일)_260309~260315.xlsx")
    if not path.exists():
        pytest.skip("real input workbook not available")
    return path


def _xlsx_part_counts(path: Path) -> tuple[int, int]:
    with ZipFile(path) as zf:
        media = [name for name in zf.namelist() if name.startswith("xl/media/")]
        drawings = [name for name in zf.namelist() if name.startswith("xl/drawings/")]
    return len(media), len(drawings)


def _apply_card_outline(ws, start_col: int):
    medium = Side(style="medium", color="000000")
    for row_idx in range(2, 36):
        for col_idx in range(start_col, start_col + 6):
            cell = ws.cell(row_idx, col_idx)
            if row_idx == 2:
                cell.border = Border(
                    top=medium,
                    left=cell.border.left,
                    right=cell.border.right,
                    bottom=cell.border.bottom,
                )
            if row_idx == 35:
                cell.border = Border(
                    bottom=medium,
                    left=cell.border.left,
                    right=cell.border.right,
                    top=cell.border.top,
                )
            if col_idx == start_col:
                cell.border = Border(
                    left=medium,
                    right=cell.border.right,
                    top=cell.border.top,
                    bottom=cell.border.bottom,
                )
            if col_idx == start_col + 5:
                cell.border = Border(
                    right=medium,
                    left=cell.border.left,
                    top=cell.border.top,
                    bottom=cell.border.bottom,
                )


class TestVocClassifier(unittest.TestCase):
    def test_rule_classify_exchange(self):
        import classify_voc

        assert classify_voc.rule_classify("[교환] 사이즈 바꿔요") == "반응_교환"

    def test_rule_classify_seeking(self):
        import classify_voc

        assert classify_voc.rule_classify("[구해요] 130 구해요") == "반응_구해요"

    def test_rule_classify_no_match(self):
        import classify_voc

        assert classify_voc.rule_classify("오늘 배송 빠르네요") is None

    def test_rule_classify_info_keywords(self):
        import classify_voc

        assert (
            classify_voc.rule_classify(
                "무신사) 푸마 키즈 스피드캣 발렛 키즈 79000원 무배"
            )
            == "반응_정보"
        )
        assert classify_voc.rule_classify("캐리마켓 최대 80% 할인 핫딜") == "반응_정보"

    def test_rule_classify_info_price_pattern_only(self):
        import classify_voc

        assert classify_voc.rule_classify("캐리마켓 14,310원 핫딜") == "반응_정보"
        assert classify_voc.rule_classify("원픽이긴 한데 구매처 아시는분") is None

    def test_rule_priority_keeps_seeking_before_info(self):
        import classify_voc

        assert (
            classify_voc.rule_classify("[구해요] 150 구해요 79,000원") == "반응_구해요"
        )

    def test_rule_classify_does_not_turn_purchase_question_into_info(self):
        import classify_voc

        assert classify_voc.rule_classify("푸마 스피드캣 고 구매처 아시는분 ㅜ") is None

    def test_build_classification_prompt(self):
        import classify_voc

        prompt = classify_voc.build_classification_prompt("이 제품 어디서 구매하나요")
        for category in VALID_CATEGORIES:
            assert category in prompt
        assert any(
            token in prompt for token in ["궁금", "교환", "구해", "가격", "성분"]
        )
        assert any(token in prompt for token in ["무배", "리셀", "핫딜", "신상"])

    def test_github_gpt41_adapter_request_format(self):
        import classify_voc

        adapter = classify_voc.GithubModelsAdapter(
            api_key="test-key", model="openai/gpt-4.1"
        )
        body = adapter.build_request_body("테스트 제목")

        assert body["model"] == "openai/gpt-4.1"
        assert "messages" in body
        assert isinstance(body["messages"], list)
        assert len(body["messages"]) >= 1

    def test_github_gpt41_adapter_parse_response(self):
        import classify_voc

        adapter = classify_voc.GithubModelsAdapter(
            api_key="test-key", model="openai/gpt-4.1"
        )
        mock_response = {"choices": [{"message": {"content": "반응_문의"}}]}

        result = adapter.parse_response(mock_response)
        assert result == "반응_문의"

    def test_github_gpt4o_adapter_request_format(self):
        import classify_voc

        adapter = classify_voc.GithubModelsAdapter(
            api_key="test-key", model="openai/gpt-4o"
        )
        body = adapter.build_request_body("테스트 제목")

        assert body["model"] == "openai/gpt-4o"
        assert "messages" in body

    def test_github_gpt4o_adapter_parse_response(self):
        import classify_voc

        adapter = classify_voc.GithubModelsAdapter(
            api_key="test-key", model="openai/gpt-4o"
        )
        mock_response = {"choices": [{"message": {"content": "반응_교환"}}]}

        result = adapter.parse_response(mock_response)
        assert result == "반응_교환"

    def test_validate_classification_result(self):
        import classify_voc

        assert classify_voc.validate_result("반응_문의") == "반응_문의"
        assert classify_voc.validate_result("반응_교환") == "반응_교환"
        assert classify_voc.validate_result("반응_구해요") == "반응_구해요"
        assert classify_voc.validate_result("반응_정보") == "반응_정보"
        assert classify_voc.validate_result("반응_기타") == "반응_기타"
        assert classify_voc.validate_result("알 수 없음") == "반응_기타"
        assert classify_voc.validate_result("") == "반응_기타"
        assert classify_voc.validate_result("긍정_제품후기") == "반응_기타"

    def test_preprocess_title(self):
        import classify_voc

        newline_case = classify_voc.preprocess_title("제목\n부제목")
        assert newline_case in ("제목 부제목", "제목")

        assert classify_voc.preprocess_title("  공백 포함  ") == "공백 포함"

        try:
            empty_result = classify_voc.preprocess_title("")
            assert empty_result is None
        except ValueError:
            pass

        try:
            none_result = classify_voc.preprocess_title(None)
            assert none_result is None
        except TypeError:
            pass

    def test_build_image_index_empty(self):
        import classify_voc

        wb, _ = _build_base_workbook()
        assert classify_voc.build_image_index(wb) == {}

    def test_build_card_index_includes_duplicate_links(self):
        import classify_voc

        wb, _ = _build_base_workbook()
        ws1 = wb.create_sheet("IB")
        ws2 = wb.create_sheet("KR")
        link = "https://example.com/post/1"
        _add_card(ws1, 2, link)
        _add_card(ws2, 14, link)

        card_index = classify_voc.build_card_index(wb)

        assert link in card_index
        assert len(card_index[link]) == 2

    def test_extract_article_text_removes_notice(self):
        import classify_voc

        html = """
        <div class="article viewer">
          <p>회원간의 거래 분쟁에 대한 공론화 금지</p>
          <p>가격이 얼마인가요?</p>
        </div>
        """

        text = classify_voc.extract_article_text(html)

        assert text is not None
        assert "회원간의 거래 분쟁에 대한 공론화 금지" not in text
        assert "가격이 얼마인가요?" in text

    def test_extract_article_text_supports_reversed_class_order(self):
        import classify_voc

        html = """
        <div class="viewer article">
          <p>사이즈가 궁금해요</p>
        </div>
        """

        text = classify_voc.extract_article_text(html)

        assert text == "사이즈가 궁금해요"

    def test_patch_sheet_xml_creates_missing_cell(self):
        import classify_voc

        xml = b"""<?xml version="1.0" encoding="UTF-8"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData>
            <row r="5"></row>
          </sheetData>
        </worksheet>
        """

        patched = classify_voc._patch_sheet_xml(xml, {"G5": "반응_문의"})

        assert b'c r="G5"' in patched
        assert b't="inlineStr"' in patched

    def test_patch_sheet_xml_preserves_ignorable_namespace_prefixes(self):
        import classify_voc

        xml = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
          mc:Ignorable="x14ac"
          xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
          <sheetFormatPr defaultRowHeight="16.5" x14ac:dyDescent="0.3"/>
          <sheetData>
            <row r="5" x14ac:dyDescent="0.3"><c r="G5" s="35" t="s"><v>25</v></c></row>
          </sheetData>
        </worksheet>
        """

        patched = classify_voc._patch_sheet_xml(xml, {"G5": "반응_문의"})
        patched_text = patched.decode("utf-8")

        assert (
            'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
            in patched_text
        )
        assert (
            'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"'
            in patched_text
        )
        assert 'mc:Ignorable="x14ac"' in patched_text
        assert 'x14ac:dyDescent="0.3"' in patched_text

    @patch("classify_voc.urllib.request.urlopen")
    def test_classify_item_rule_bypasses_api(self, mock_urlopen):
        import classify_voc

        result = classify_voc.classify_item(
            "[교환] 사이즈 교환해요",
            link="https://example.com/post/1",
            image_index={"https://example.com/post/1": [b"img"]},
            api_key="test-key",
        )

        assert result == "반응_교환"
        mock_urlopen.assert_not_called()

    @patch("classify_voc.urllib.request.urlopen")
    def test_classify_item_llm_then_image_ocr(self, mock_urlopen):
        import classify_voc

        image_index = {"https://example.com/post/1": [b"image-a", b"image-b"]}

        def side_effect(request, *args, **kwargs):
            if request.data is None:
                response = MagicMock()
                response.__enter__.return_value.read.return_value = b"<html></html>"
                return response
            body = json.loads(request.data.decode())
            user_content = body["messages"][1]["content"]
            if isinstance(user_content, str):
                return _mock_http_response(
                    {"choices": [{"message": {"content": "반응_기타"}}]}
                )
            return _mock_http_response(
                {"choices": [{"message": {"content": "반응_구해요"}}]}
            )

        mock_urlopen.side_effect = side_effect
        result = classify_voc.classify_item(
            "애매한 제목",
            link="https://example.com/post/1",
            image_index=image_index,
            api_key="test-key",
        )
        assert result == "반응_구해요"

    @patch("classify_voc.urllib.request.urlopen")
    def test_classify_item_llm_then_article_body(self, mock_urlopen):
        import classify_voc

        html = """
        <div class="article viewer">
          <p>회원간의 거래 분쟁에 대한 공론화 금지</p>
          <p>가격이 얼마인가요?</p>
        </div>
        """

        def side_effect(request, *args, **kwargs):
            if isinstance(request, str):
                response = MagicMock()
                response.__enter__.return_value.read.return_value = html.encode("utf-8")
                return response

            if request.data is None:
                response = MagicMock()
                response.__enter__.return_value.read.return_value = html.encode("utf-8")
                return response

            body = json.loads(request.data.decode())
            prompt = body["messages"][0]["content"]
            if "분류 대상 본문" in prompt:
                return _mock_http_response(
                    {"choices": [{"message": {"content": "반응_정보"}}]}
                )
            return _mock_http_response(
                {"choices": [{"message": {"content": "반응_기타"}}]}
            )

        mock_urlopen.side_effect = side_effect

        result = classify_voc.classify_item(
            "애매한 제목",
            link="https://example.com/post/1",
            image_index={},
            api_key="test-key",
        )

        assert result == "반응_정보"

    @patch("classify_voc.urllib.request.urlopen")
    def test_classify_with_image_format(self, mock_urlopen):
        import classify_voc

        captured: dict[str, Any] = {}

        def side_effect(request, *args, **kwargs):
            captured.update(json.loads(request.data.decode()))
            return _mock_http_response(
                {"choices": [{"message": {"content": "반응_문의"}}]}
            )

        mock_urlopen.side_effect = side_effect
        result = classify_voc.classify_with_image([b"abc", b"def"], api_key="test-key")

        assert result == "반응_문의"
        assert captured["model"] == "openai/gpt-4.1"
        user_message = captured["messages"][1]["content"]
        assert user_message[-1]["type"] == "text"
        assert user_message[0]["type"] == "image_url"
        assert user_message[1]["type"] == "image_url"
        assert user_message[0]["image_url"]["url"].startswith("data:image/jpeg;base64,")

    def test_read_target_rows_includes_link(self):
        import classify_voc

        wb, ws = _build_base_workbook()
        ws.cell(5, 7).value = "반응_제품반응"
        ws.cell(5, 10).value = "이 제품 어디서 사나요"
        ws.cell(5, 11).value = "원문"
        ws.cell(5, 11).hyperlink = "https://example.com/post/5"

        rows = classify_voc.read_target_rows(wb)

        assert len(rows) == 1
        assert rows[0]["link"] == "https://example.com/post/5"

    @patch("classify_voc.urllib.request.urlopen")
    def test_fallback_to_gpt4o(self, mock_urlopen):
        import classify_voc

        call_count = 0

        def side_effect(request, *args, **kwargs):
            nonlocal call_count
            call_count += 1
            body = json.loads(request.data.decode())
            if body.get("model") == "openai/gpt-4.1":
                raise _http_error(request.full_url)
            if body.get("model") == "openai/gpt-4o":
                return _mock_http_response(
                    {"choices": [{"message": {"content": "반응_정보"}}]}
                )
            raise AssertionError(f"unexpected model: {body.get('model')}")

        mock_urlopen.side_effect = side_effect
        result = classify_voc.classify_title("성분이 뭔가요", api_key="test-key")
        assert result == "반응_정보"

    def test_read_target_rows(self):
        import classify_voc

        wb, ws = _build_base_workbook()
        ws.cell(5, 7).value = "반응_제품반응"
        ws.cell(5, 10).value = "이 제품 어디서 사나요"
        ws.cell(6, 7).value = "긍정_제품후기"
        ws.cell(6, 10).value = "좋아요"

        rows = classify_voc.read_target_rows(wb)

        assert len(rows) == 1
        assert rows[0]["row"] == 5
        assert rows[0]["title"] == "이 제품 어디서 사나요"
        assert all(item.get("row") != 6 for item in rows)

    def test_write_classification_result(self):
        import classify_voc

        wb, ws = _build_base_workbook()
        ws.cell(5, 7).value = "반응_제품반응"
        ws.cell(5, 5).value = "원본값"

        classify_voc.write_classification_result(wb, row=5, category="반응_문의")

        assert ws.cell(5, 7).value == "반응_문의"
        assert ws.cell(5, 5).value == "원본값"

    def test_preserve_non_target_rows(self):
        import classify_voc

        wb, ws = _build_base_workbook()
        ws.cell(5, 7).value = "반응_제품반응"
        ws.cell(5, 5).value = "A"
        ws.cell(6, 7).value = "긍정_제품후기"
        ws.cell(6, 5).value = "B"
        ws.cell(7, 7).value = "부정_제품반응"
        ws.cell(7, 5).value = "C"

        classify_voc.write_classification_result(wb, row=5, category="반응_문의")

        assert ws.cell(6, 7).value == "긍정_제품후기"
        assert ws.cell(7, 7).value == "부정_제품반응"
        assert ws.cell(6, 5).value == "B"


@pytest.mark.integration
@patch("classify_voc.urllib.request.urlopen")
def test_classify_pipeline_with_mock_api(mock_urlopen):
    import classify_voc

    wb, ws = _build_base_workbook()
    ws.cell(5, 7).value = "반응_제품반응"
    ws.cell(5, 10).value = "재고 있나요"
    ws.cell(6, 7).value = "반응_제품반응"
    ws.cell(6, 10).value = "교환 가능한가요"
    ws.cell(7, 7).value = "반응_제품반응"
    ws.cell(7, 10).value = "가격 궁금해요"

    categories = ["반응_문의", "반응_교환", "반응_정보"]

    def side_effect(request, *args, **kwargs):
        _ = BytesIO(b"request")
        if "chat/completions" in str(getattr(request, "full_url", request)):
            payload = {"choices": [{"message": {"content": categories.pop(0)}}]}
            return _mock_http_response(payload)
        raise AssertionError(f"unexpected url: {getattr(request, 'full_url', request)}")

    mock_urlopen.side_effect = side_effect
    pipeline = _pipeline_callable(classify_voc)
    assert pipeline is not None
    pipeline(wb, api_key="test-key", output_path=None)

    assert ws.cell(5, 7).value in VALID_CATEGORIES
    assert ws.cell(6, 7).value in VALID_CATEGORIES
    assert ws.cell(7, 7).value in VALID_CATEGORIES
    assert ws.cell(5, 7).value != "반응_제품반응"
    assert ws.cell(6, 7).value != "반응_제품반응"
    assert ws.cell(7, 7).value != "반응_제품반응"


@pytest.mark.integration
@patch("classify_voc.urllib.request.urlopen")
def test_output_preserves_all_sheets(mock_urlopen):
    import classify_voc

    wb, ws = _build_base_workbook()
    ws.cell(5, 7).value = "반응_제품반응"
    ws.cell(5, 10).value = "어디서 구매"
    ws2 = wb.create_sheet("Sheet2")
    ws2.cell(1, 1).value = "keep"
    ws3 = wb.create_sheet("Sheet3")
    ws3.cell(2, 2).value = "keep-too"

    mock_urlopen.return_value = _mock_http_response(
        {"choices": [{"message": {"content": "반응_문의"}}]}
    )
    pipeline = _pipeline_callable(classify_voc)
    assert pipeline is not None
    pipeline(wb, api_key="test-key", output_path=None)

    assert wb.sheetnames == ["2026 list", "Sheet2", "Sheet3"]
    assert wb["Sheet2"].cell(1, 1).value == "keep"
    assert wb["Sheet3"].cell(2, 2).value == "keep-too"


@pytest.mark.integration
@patch("classify_voc.urllib.request.urlopen")
def test_pipeline_updates_linked_cards(mock_urlopen):
    import classify_voc

    wb, ws = _build_base_workbook()
    ws.cell(5, 7).value = "반응_제품반응"
    ws.cell(5, 10).value = "어디서 구매"
    ws.cell(5, 11).value = "원문"
    ws.cell(5, 11).hyperlink = "https://example.com/post/5"

    card_sheet = wb.create_sheet("IB")
    _add_card(card_sheet, 2, "https://example.com/post/5")
    _add_card(card_sheet, 14, "https://example.com/post/5")

    mock_urlopen.return_value = _mock_http_response(
        {"choices": [{"message": {"content": "반응_문의"}}]}
    )

    pipeline = _pipeline_callable(classify_voc)
    assert pipeline is not None
    pipeline(wb, api_key="test-key", output_path=None)

    assert ws.cell(5, 7).value == "반응_문의"
    assert card_sheet.cell(5, 3).value == "반응_문의"
    assert card_sheet.cell(5, 15).value == "반응_문의"


@pytest.mark.integration
@patch("classify_voc.urllib.request.urlopen")
def test_classification_summary(mock_urlopen, capsys):
    import classify_voc

    wb, ws = _build_base_workbook()
    ws.cell(5, 7).value = "반응_제품반응"
    ws.cell(5, 10).value = "제목1"
    ws.cell(6, 7).value = "반응_제품반응"
    ws.cell(6, 10).value = "제목2"
    ws.cell(7, 7).value = "반응_제품반응"
    ws.cell(7, 10).value = "제목3"

    mock_urlopen.return_value = _mock_http_response(
        {"choices": [{"message": {"content": "반응_문의"}}]}
    )
    pipeline = _pipeline_callable(classify_voc)
    assert pipeline is not None
    pipeline(wb, api_key="test-key", output_path=None)

    captured = capsys.readouterr()
    assert any(token in captured.err for token in ["[1/", "[2/", "[3/"])


@pytest.mark.integration
@patch("classify_voc.urllib.request.urlopen")
def test_api_error_fallback_full_pipeline(mock_urlopen):
    import classify_voc

    wb, ws = _build_base_workbook()
    ws.cell(5, 7).value = "반응_제품반응"
    ws.cell(5, 10).value = "성분 알려주세요"
    ws.cell(6, 7).value = "반응_제품반응"
    ws.cell(6, 10).value = "교환 정책은?"

    def side_effect(request, *args, **kwargs):
        body = json.loads(request.data.decode())
        if body.get("model") == "openai/gpt-4.1":
            raise _http_error(request.full_url)
        if body.get("model") == "openai/gpt-4o":
            return _mock_http_response(
                {"choices": [{"message": {"content": "반응_정보"}}]}
            )
        raise AssertionError(f"unexpected model: {body.get('model')}")

    mock_urlopen.side_effect = side_effect
    pipeline = _pipeline_callable(classify_voc)
    assert pipeline is not None
    pipeline(wb, api_key="test-key", output_path=None)

    assert ws.cell(5, 7).value in VALID_CATEGORIES
    assert ws.cell(6, 7).value in VALID_CATEGORIES
    assert ws.cell(5, 7).value != "반응_제품반응"
    assert ws.cell(6, 7).value != "반응_제품반응"


@pytest.mark.integration
@patch("classify_voc.urllib.request.urlopen")
def test_hybrid_pipeline_integration(mock_urlopen):
    import classify_voc

    wb, ws = _build_base_workbook()
    ws.cell(5, 7).value = "반응_제품반응"
    ws.cell(5, 10).value = "[교환] 140으로 바꿔요"
    ws.cell(6, 7).value = "반응_제품반응"
    ws.cell(6, 10).value = "가격 궁금합니다"

    mock_urlopen.return_value = _mock_http_response(
        {"choices": [{"message": {"content": "반응_정보"}}]}
    )
    pipeline = _pipeline_callable(classify_voc)
    assert pipeline is not None
    pipeline(wb, api_key="test-key", output_path=None)

    assert ws.cell(5, 7).value == "반응_교환"
    assert ws.cell(6, 7).value == "반응_정보"
    assert mock_urlopen.call_count == 0


@pytest.mark.integration
def test_openpyxl_roundtrip_drops_media_on_real_workbook():
    src = _real_input_workbook()
    before_media, before_drawings = _xlsx_part_counts(src)

    with tempfile.TemporaryDirectory() as tmpdir:
        out = Path(tmpdir) / "roundtrip.xlsx"
        wb = openpyxl.load_workbook(src)
        wb.save(out)

        after_media, after_drawings = _xlsx_part_counts(out)

    assert after_media < before_media
    assert after_drawings < before_drawings


@pytest.mark.integration
@patch("classify_voc.urllib.request.urlopen")
def test_saved_output_preserves_media_on_real_workbook(mock_urlopen):
    import classify_voc

    src = _real_input_workbook()
    before_media, before_drawings = _xlsx_part_counts(src)
    out = Path(tempfile.mkdtemp()) / src.name

    mock_urlopen.return_value = _mock_http_response(
        {"choices": [{"message": {"content": "반응_문의"}}]}
    )

    wb = openpyxl.load_workbook(src)
    classify_voc.run_classification(
        wb, api_key="test-key", output_path=out, source_path=src
    )

    after_media, after_drawings = _xlsx_part_counts(out)

    assert after_media == before_media
    assert after_drawings == before_drawings


@pytest.mark.integration
@patch("classify_voc.urllib.request.urlopen")
def test_saved_output_reloads_linked_card_updates(mock_urlopen):
    import classify_voc

    with tempfile.TemporaryDirectory() as tmpdir:
        src = Path(tmpdir) / "source.xlsx"
        out = Path(tmpdir) / "result.xlsx"

        wb, ws = _build_base_workbook()
        ws.cell(5, 7).value = "반응_제품반응"
        ws.cell(5, 10).value = "어디서 구매"
        ws.cell(5, 11).value = "원문"
        ws.cell(5, 11).hyperlink = "https://example.com/post/5"
        card_sheet = wb.create_sheet("IB")
        _add_card(card_sheet, 2, "https://example.com/post/5")
        _add_card(card_sheet, 14, "https://example.com/post/5")
        wb.save(src)

        mock_urlopen.return_value = _mock_http_response(
            {"choices": [{"message": {"content": "반응_문의"}}]}
        )

        loaded = openpyxl.load_workbook(src)
        classify_voc.run_classification(
            loaded, api_key="test-key", output_path=out, source_path=src
        )

        saved = openpyxl.load_workbook(out)

    assert saved["2026 list"].cell(5, 7).value == "반응_문의"
    assert saved["IB"].cell(5, 3).value == "반응_문의"
    assert saved["IB"].cell(5, 17).value == "반응_문의"


@pytest.mark.integration
@patch("classify_voc.urllib.request.urlopen")
def test_saved_output_without_linked_card_mapping_is_readable(mock_urlopen):
    import classify_voc

    with tempfile.TemporaryDirectory() as tmpdir:
        src = Path(tmpdir) / "source.xlsx"
        out = Path(tmpdir) / "result.xlsx"

        wb, ws = _build_base_workbook()
        ws.cell(5, 7).value = "반응_제품반응"
        ws.cell(5, 10).value = "가격 궁금해요"
        wb.save(src)

        mock_urlopen.return_value = _mock_http_response(
            {"choices": [{"message": {"content": "반응_정보"}}]}
        )

        loaded = openpyxl.load_workbook(src)
        classify_voc.run_classification(
            loaded, api_key="test-key", output_path=out, source_path=src
        )

        saved = openpyxl.load_workbook(out)

    assert saved["2026 list"].cell(5, 7).value == "반응_정보"


@pytest.mark.integration
@patch("classify_voc.urllib.request.urlopen")
def test_saved_output_preserves_worksheet_namespace_prefixes(mock_urlopen):
    import classify_voc

    src = _real_input_workbook()
    out = Path(tempfile.mkdtemp()) / src.name

    mock_urlopen.return_value = _mock_http_response(
        {"choices": [{"message": {"content": "반응_문의"}}]}
    )

    wb = openpyxl.load_workbook(src)
    classify_voc.run_classification(
        wb, api_key="test-key", output_path=out, source_path=src
    )

    with ZipFile(out) as zf:
        sheet_xml = zf.read("xl/worksheets/sheet2.xml").decode("utf-8")

    xmlns = dict(re.findall(r'xmlns(?::([^=]+))?="([^"]+)"', sheet_xml[:600]))
    ignorable_match = re.search(r'([A-Za-z0-9_]+):Ignorable="([^"]+)"', sheet_xml[:600])

    assert xmlns["mc"] == "http://schemas.openxmlformats.org/markup-compatibility/2006"
    assert (
        xmlns["x14ac"] == "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
    )
    assert ignorable_match is not None
    assert ignorable_match.group(1) == "mc"
    assert ignorable_match.group(2) == "x14ac"


def test_apply_layout_to_output_workbook_formats_pages():
    import classify_voc

    with tempfile.TemporaryDirectory() as tmpdir:
        path = Path(tmpdir) / "layout.xlsx"
        wb, _ = _build_base_workbook()
        wb.create_sheet("주현황보고", 0)
        ws = wb.create_sheet("IB")
        ws.cell(2, 2).value = "Voice Of Customer"
        ws.cell(2, 8).value = "Voice Of Customer"
        ws.cell(5, 3).value = "반응_제품반응"
        ws.cell(5, 9).value = "반응_제품반응"
        _apply_card_outline(ws, 2)
        _apply_card_outline(ws, 8)
        original_left_margin = ws.page_margins.left
        ws.print_area = "'IB'!$B$2:$M$35"
        wb.save(path)

        classify_voc._apply_layout_to_output_workbook(path)

        saved = openpyxl.load_workbook(path)

    assert "주현황보고" not in saved.sheetnames
    assert saved["IB"]["B2"].value == "Voice Of Customer"
    assert saved["IB"]["I2"].value == "Voice Of Customer"
    assert saved["IB"]["H2"].value is None
    assert saved["IB"]["O2"].value is None
    assert saved["IB"].print_area == "'IB'!$B$2:$O$35"
    assert saved["IB"].column_dimensions["H"].width == pytest.approx(0.875)
    assert saved["IB"].column_dimensions["O"].width == pytest.approx(0.875)
    assert saved["IB"]["G2"].border.right.style == "medium"
    assert saved["IB"]["N2"].border.right.style == "medium"
    assert saved["IB"].print_options.horizontalCentered is True
    assert saved["IB"].print_options.verticalCentered is True
    assert saved["IB"].page_setup.pageOrder == "overThenDown"
    assert saved["IB"].page_margins.left == pytest.approx(original_left_margin + 0.2)


@pytest.mark.integration
@patch("classify_voc.urllib.request.urlopen")
def test_saved_output_applies_layout_rules(mock_urlopen):
    import classify_voc

    src = _real_input_workbook()
    out = Path(tempfile.mkdtemp()) / src.name

    mock_urlopen.return_value = _mock_http_response(
        {"choices": [{"message": {"content": "반응_문의"}}]}
    )

    wb = openpyxl.load_workbook(src)
    classify_voc.run_classification(
        wb, api_key="test-key", output_path=out, source_path=src
    )

    saved = openpyxl.load_workbook(out)

    assert "주현황보고" not in saved.sheetnames
    assert saved["IB"].print_area == "'IB'!$B$2:$AQ$35"
    assert saved["IB"].column_dimensions["H"].width == pytest.approx(0.875)
    assert saved["IB"].column_dimensions["O"].width == pytest.approx(0.875)
    assert saved["IB"].page_setup.pageOrder == "overThenDown"
    assert saved["IB"].print_options.horizontalCentered is True
    assert saved["IB"].print_options.verticalCentered is True
    assert saved["IB"].page_margins.left == pytest.approx(0.45)
