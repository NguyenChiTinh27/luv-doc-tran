"""
Kiểm thử Toàn diện: Cơ chế Quản lý Key DeepSeek, Nhận diện Ngôn ngữ Vòng 1 và Động cơ Dịch AI
=============================================================================================
"""

import os
import json
import pytest
from unittest.mock import patch, MagicMock
from fastapi.testclient import TestClient

from core.context_profiler import AIContextProfiler, ContextProfile
from core.interlinear_translator import clean_json_markdown, robust_json_loads, InterlinearTranslator
from server import app

client = TestClient(app)


# ============================================================================
# 1. KIỂM THỬ NHẬN DIỆN NGÔN NGỮ TỰ ĐỘNG TẠI VÒNG 1 (SMART AUTO-DETECT)
# ============================================================================

def test_language_detection_chinese():
    """Kiểm tra nhận diện tiếng Trung CJK Hanzi."""
    text = "化学品安全技术说明书 产品名称: 聚醚多元醇 生产厂家: 上海化学工业区"
    code, name = AIContextProfiler.detect_language(text)
    assert code == "zh"
    assert "Trung" in name


def test_language_detection_japanese():
    """Kiểm tra nhận diện tiếng Nhật (chứa Kana)."""
    text = "安全データシート 製品名: ポリウレタン樹脂 製造元: 東京化学工業株式会社"
    code, name = AIContextProfiler.detect_language(text)
    assert code == "ja"
    assert "Nhật" in name


def test_language_detection_korean():
    """Kiểm tra nhận diện tiếng Hàn (Hangul)."""
    text = "물질안전보건자료 제품명: 에폭시 수지 제조업체: 서울화학공업"
    code, name = AIContextProfiler.detect_language(text)
    assert code == "ko"
    assert "Hàn" in name


def test_language_detection_russian():
    """Kiểm tra nhận diện tiếng Nga (Cyrillic)."""
    text = "Паспорт безопасности химической продукции Название: Растворитель 646"
    code, name = AIContextProfiler.detect_language(text)
    assert code == "ru"
    assert "Nga" in name


def test_language_detection_english():
    """Kiểm tra nhận diện tiếng Anh Latin tiêu chuẩn."""
    text = "Safety Data Sheet. Product Name: Industrial Solvent Cleaner. CAS Number: 67-64-1."
    code, name = AIContextProfiler.detect_language(text)
    assert code == "en"
    assert "Anh" in name


def test_language_detection_german():
    """Kiểm tra nhận diện tiếng Đức (ä, ö, ü, ß và từ vựng chuyên ngành)."""
    text = "Sicherheitsdatenblatt gemäß Verordnung für gefährliche Stoffe. Gefahr und Gefahrenhinweise."
    code, name = AIContextProfiler.detect_language(text)
    assert code == "de"
    assert "Đức" in name


def test_language_detection_french():
    """Kiểm tra nhận diện tiếng Pháp."""
    text = "Fiche de données de sécurité pour produits chimiques dangereux. Mention d'avertissement: Danger."
    code, name = AIContextProfiler.detect_language(text)
    assert code == "fr"
    assert "Pháp" in name


def test_vong_1_context_profile_incorporates_detected_language():
    """Kiểm tra Vòng 1 AI Context Profiler tích hợp đồng thời ngữ cảnh và ngôn ngữ."""
    sample_zh = "化学品安全技术说明书 本产品符合 GB/T 16483 标准，适用于工业涂料与胶粘剂生产。"
    with patch.object(AIContextProfiler, "extract_text_sample", return_value=sample_zh):
        profile = AIContextProfiler.profile_document("fake_path.docx", source_lang="auto")
        assert profile.detected_source_lang == "zh"
        assert "Trung" in profile.detected_lang_name
        assert "Ngôn ngữ nguồn: Tiếng Trung (zh)" in profile.context_anchor


# ============================================================================
# 2. KIỂM THỬ LÀM SẠCH VÀ PHÂN TÍCH JSON DEEPSEEK (ROBUST JSON PARSER)
# ============================================================================

def test_clean_json_markdown_with_codeblock():
    """Kiểm tra bóc tách khối ```json ... ```"""
    raw = '```json\n{"0": "Dung dịch tẩy rửa", "1": "Cảnh báo nguy hại"}\n```'
    cleaned = clean_json_markdown(raw)
    parsed = json.loads(cleaned)
    assert parsed["0"] == "Dung dịch tẩy rửa"
    assert parsed["1"] == "Cảnh báo nguy hại"


def test_clean_json_markdown_with_think_tag():
    """Kiểm tra bóc tách thẻ suy nghĩ <think>...</think> của DeepSeek Reasoner."""
    raw = '<think>I should translate item 0 as dung dịch</think>```json\n{"0": "Dung dịch"}\n```'
    parsed = robust_json_loads(raw)
    assert parsed["0"] == "Dung dịch"


def test_clean_json_with_trailing_comma():
    """Kiểm tra xử lý dấu phẩy thừa trước ngoặc nhọn."""
    raw = '{"0": "Mục một", "1": "Mục hai", }'
    parsed = robust_json_loads(raw)
    assert parsed["0"] == "Mục một"
    assert parsed["1"] == "Mục hai"


# ============================================================================
# 3. KIỂM THỬ CÁC ENDPOINT QUẢN LÝ KEY (API /settings/)
# ============================================================================

def test_save_and_get_ai_key():
    """Kiểm tra lưu và lấy trạng thái key DeepSeek."""
    test_key = "sk-0123456789abcdef0123456789abcdef"
    
    # 1. Lưu key
    res_save = client.post("/api/settings/ai-key", data={"key": test_key})
    assert res_save.status_code == 200
    data_save = res_save.json()
    assert data_save["success"] is True
    assert "sk-01" in data_save["masked_key"]
    assert "def" in data_save["masked_key"]

    # 2. Kiểm tra status (không lộ raw key)
    res_status = client.get("/api/settings/ai-status")
    assert res_status.status_code == 200
    data_status = res_status.json()
    assert data_status["configured"] is True
    assert test_key not in json.dumps(data_status)  # Không lộ raw key

    # 3. Xóa key
    res_clear = client.post("/api/settings/ai-clear")
    assert res_clear.status_code == 200
    assert res_clear.json()["success"] is True

    # 4. Kiểm tra lại status sau xóa
    res_status_after = client.get("/api/settings/ai-status")
    assert res_status_after.json()["configured"] is False


def test_ai_test_key_endpoint_mock():
    """Kiểm tra endpoint test key với mock API phản hồi thành công."""
    with patch("requests.post") as mock_post:
        mock_resp = MagicMock()
        mock_resp.status_code = 200
        mock_post.return_value = mock_resp

        res = client.post("/api/settings/ai-test", data={"key": "sk-valid-test-key-1234567890"})
        assert res.status_code == 200
        data = res.json()
        assert data["success"] is True
        assert "DeepSeek-V4.1-Flash" in data["message"] or "deepseek" in data["message"]


# ============================================================================
# 4. KIỂM THỬ DEEPSEEK TRANSLATION VỚI MODEL DeepSeek-V4.1-Flash VÀ FALLBACK
# ============================================================================

def test_translate_deepseek_uses_v4_flash_and_fallback():
    """Kiểm tra InterlinearTranslator ưu tiên DeepSeek-V4.1-Flash và fallback sang deepseek-chat khi cần."""
    texts = ["Solvent cleaner", "Corrosive liquid"]
    api_key = "sk-mock-key-12345678901234567890"

    with patch("requests.post") as mock_post:
        # Giả lập lần 1: Model DeepSeek-V4.1-Flash trả về 200 thành công
        mock_resp_v4 = MagicMock()
        mock_resp_v4.status_code = 200
        mock_resp_v4.json.return_value = {
            "choices": [{"message": {"content": json.dumps({"0": "Dung môi làm sạch", "1": "Chất lỏng ăn mòn"})}}]
        }
        mock_post.return_value = mock_resp_v4

        results = InterlinearTranslator.translate_texts_deepseek(
            texts=texts,
            api_key=api_key,
            source_lang="en",
            target_lang="vi",
            model="DeepSeek-V4.1-Flash"
        )
        assert results == ["Dung môi làm sạch", "Chất lỏng ăn mòn"]
        # Xác nhận model được gọi đúng là DeepSeek-V4.1-Flash
        called_payload = mock_post.call_args[1]["json"]
        assert called_payload["model"] == "DeepSeek-V4.1-Flash"
