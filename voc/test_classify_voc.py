"""RED-phase tests for VOC classification script."""
# pyright: reportMissingImports=false, reportAttributeAccessIssue=false, reportOptionalMemberAccess=false
import json
import unittest
from unittest.mock import patch, MagicMock
from io import BytesIO

import openpyxl
import pytest


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


def _mock_http_response(payload):
    response = MagicMock()
    response.__enter__.return_value.read.return_value = json.dumps(payload).encode("utf-8")
    return response


def _http_error(url):
    http_error_cls = __import__("urllib.error", fromlist=["HTTPError"]).HTTPError
    return http_error_cls(url, 500, "mock error", hdrs=None, fp=None)


def _pipeline_callable(module):
    return getattr(module, "run_classification", None) or getattr(module, "classify_workbook", None)


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

    def test_build_classification_prompt(self):
        import classify_voc

        prompt = classify_voc.build_classification_prompt("이 제품 어디서 구매하나요")
        for category in VALID_CATEGORIES:
            assert category in prompt
        assert any(token in prompt for token in ["궁금", "교환", "구해", "가격", "성분"])

    def test_github_gpt41_adapter_request_format(self):
        import classify_voc

        adapter = classify_voc.GithubModelsAdapter(api_key="test-key", model="openai/gpt-4.1")
        body = adapter.build_request_body("테스트 제목")

        assert body["model"] == "openai/gpt-4.1"
        assert "messages" in body
        assert isinstance(body["messages"], list)
        assert len(body["messages"]) >= 1

    def test_github_gpt41_adapter_parse_response(self):
        import classify_voc

        adapter = classify_voc.GithubModelsAdapter(api_key="test-key", model="openai/gpt-4.1")
        mock_response = {"choices": [{"message": {"content": "반응_문의"}}]}

        result = adapter.parse_response(mock_response)
        assert result == "반응_문의"

    def test_github_gpt4o_adapter_request_format(self):
        import classify_voc

        adapter = classify_voc.GithubModelsAdapter(api_key="test-key", model="openai/gpt-4o")
        body = adapter.build_request_body("테스트 제목")

        assert body["model"] == "openai/gpt-4o"
        assert "messages" in body

    def test_github_gpt4o_adapter_parse_response(self):
        import classify_voc

        adapter = classify_voc.GithubModelsAdapter(api_key="test-key", model="openai/gpt-4o")
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
    def test_classify_with_image_format(self, mock_urlopen):
        import classify_voc

        captured: dict = {}

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
                    {
                        "choices": [
                            {"message": {"content": "반응_정보"}}
                        ]
                    }
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

    mock_urlopen.return_value = _mock_http_response({"choices": [{"message": {"content": "반응_문의"}}]})
    pipeline = _pipeline_callable(classify_voc)
    assert pipeline is not None
    pipeline(wb, api_key="test-key", output_path=None)

    assert wb.sheetnames == ["2026 list", "Sheet2", "Sheet3"]
    assert wb["Sheet2"].cell(1, 1).value == "keep"
    assert wb["Sheet3"].cell(2, 2).value == "keep-too"


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

    mock_urlopen.return_value = _mock_http_response({"choices": [{"message": {"content": "반응_문의"}}]})
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
    assert mock_urlopen.call_count == 1
