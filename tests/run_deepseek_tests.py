"""
Bộ kiểm thử chạy bằng unittest chuẩn: Kiểm tra tính năng Key DeepSeek, Auto-detect Vòng 1 và Model DeepSeek-V4.1-Flash
==================================================================================================================
"""

import os
import sys
import json
import unittest
from unittest.mock import patch, MagicMock

# Đảm bảo import được module dự án
proj_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if proj_dir not in sys.path:
    sys.path.insert(0, proj_dir)

from core.context_profiler import AIContextProfiler, ContextProfile
from core.interlinear_translator import clean_json_markdown, robust_json_loads, InterlinearTranslator
from server import app
from fastapi.testclient import TestClient

client = TestClient(app)


class TestDeepSeekAndAutoDetect(unittest.TestCase):

    def test_01_language_detection_chinese(self):
        """1. Nhận diện tiếng Trung CJK Hanzi."""
        text = "化学品安全技术说明书 产品名称: 聚醚多元醇 生产厂家: 上海化学工业区"
        code, name = AIContextProfiler.detect_language(text)
        self.assertEqual(code, "zh")
        self.assertIn("Trung", name)

    def test_02_language_detection_japanese(self):
        """2. Nhận diện tiếng Nhật (chứa Kana)."""
        text = "安全データシート 製品名: ポリウレタン樹脂 製造元: 東京化学工業株式会社"
        code, name = AIContextProfiler.detect_language(text)
        self.assertEqual(code, "ja")
        self.assertIn("Nhật", name)

    def test_03_language_detection_korean(self):
        """3. Nhận diện tiếng Hàn (Hangul)."""
        text = "물질안전보건자료 제품명: 에폭시 수지 제조업체: 서울화학공업"
        code, name = AIContextProfiler.detect_language(text)
        self.assertEqual(code, "ko")
        self.assertIn("Hàn", name)

    def test_04_language_detection_russian(self):
        """4. Nhận diện tiếng Nga (Cyrillic)."""
        text = "Паспорт безопасности химической продукции Название: Растворитель 646"
        code, name = AIContextProfiler.detect_language(text)
        self.assertEqual(code, "ru")
        self.assertIn("Nga", name)

    def test_05_language_detection_english(self):
        """5. Nhận diện tiếng Anh Latin tiêu chuẩn."""
        text = "Safety Data Sheet. Product Name: Industrial Solvent Cleaner. CAS Number: 67-64-1."
        code, name = AIContextProfiler.detect_language(text)
        self.assertEqual(code, "en")
        self.assertIn("Anh", name)

    def test_06_vong_1_context_profile_incorporates_detected_language(self):
        """6. Vòng 1 AI Context Profiler tích hợp đồng thời ngữ cảnh và ngôn ngữ."""
        sample_zh = "化学品安全技术说明书 本产品符合 GB/T 16483 标准，适用于工业涂料与胶粘剂生产。"
        profile = AIContextProfiler.profile_document(sample_zh, source_lang="auto")
        self.assertEqual(profile.detected_source_lang, "zh")
        self.assertIn("Trung", profile.detected_lang_name)
        self.assertIn("Ngôn ngữ nguồn: Tiếng Trung (zh)", profile.context_anchor)

    def test_07_clean_json_markdown_with_codeblock(self):
        """7. Bóc tách code block ```json ... ```"""
        raw = '```json\n{"0": "Dung dịch tẩy rửa", "1": "Cảnh báo nguy hại"}\n```'
        cleaned = clean_json_markdown(raw)
        parsed = json.loads(cleaned)
        self.assertEqual(parsed["0"], "Dung dịch tẩy rửa")
        self.assertEqual(parsed["1"], "Cảnh báo nguy hại")

    def test_08_clean_json_markdown_with_think_tag(self):
        """8. Bóc tách thẻ suy nghĩ <think>...</think> của DeepSeek Reasoner."""
        raw = '<think>I should translate item 0 as dung dịch</think>```json\n{"0": "Dung dịch"}\n```'
        parsed = robust_json_loads(raw)
        self.assertEqual(parsed["0"], "Dung dịch")

    def test_09_clean_json_with_trailing_comma(self):
        """9. Khử dấu phẩy thừa trước ngoặc nhọn."""
        raw = '{"0": "Mục một", "1": "Mục hai", }'
        parsed = robust_json_loads(raw)
        self.assertEqual(parsed["0"], "Mục một")
        self.assertEqual(parsed["1"], "Mục hai")

    def test_10_save_and_get_ai_key(self):
        """10. Lưu, kiểm tra status và xóa DeepSeek key qua API."""
        test_key = "sk-0123456789abcdef0123456789abcdef"
        
        # 1. Lưu key
        res_save = client.post("/api/settings/ai-key", data={"key": test_key})
        self.assertEqual(res_save.status_code, 200)
        data_save = res_save.json()
        self.assertTrue(data_save["success"])
        self.assertIn("sk-01", data_save["masked_key"])
        self.assertIn("def", data_save["masked_key"])

        # 2. Kiểm tra status (không lộ raw key)
        res_status = client.get("/api/settings/ai-status")
        self.assertEqual(res_status.status_code, 200)
        data_status = res_status.json()
        self.assertTrue(data_status["configured"])
        self.assertNotIn(test_key, json.dumps(data_status))

        # 3. Test key endpoint
        with patch("requests.post") as mock_post:
            mock_resp = MagicMock()
            mock_resp.status_code = 200
            mock_post.return_value = mock_resp

            res_test = client.post("/api/settings/ai-test", data={"key": test_key})
            self.assertEqual(res_test.status_code, 200)
            self.assertTrue(res_test.json()["success"])

        # 4. Xóa key
        res_clear = client.post("/api/settings/ai-clear")
        self.assertEqual(res_clear.status_code, 200)
        self.assertTrue(res_clear.json()["success"])

        # 5. Status sau xóa
        res_status_after = client.get("/api/settings/ai-status")
        self.assertFalse(res_status_after.json()["configured"])

    def test_11_translate_deepseek_uses_v4_flash(self):
        """11. Ưu tiên gọi model DeepSeek-V4.1-Flash."""
        texts = ["Solvent cleaner", "Corrosive liquid"]
        api_key = "sk-mock-key-12345678901234567890"

        with patch("requests.post") as mock_post:
            mock_resp = MagicMock()
            mock_resp.status_code = 200
            mock_resp.json.return_value = {
                "choices": [{"message": {"content": json.dumps({"0": "Dung môi làm sạch", "1": "Chất lỏng ăn mòn"})}}]
            }
            mock_post.return_value = mock_resp

            results = InterlinearTranslator.translate_texts_deepseek(
                texts=texts,
                api_key=api_key,
                source_lang="en",
                target_lang="vi",
                model="DeepSeek-V4.1-Flash"
            )
            self.assertEqual(results, ["Dung môi làm sạch", "Chất lỏng ăn mòn"])
            called_payload = mock_post.call_args[1]["json"]
            self.assertEqual(called_payload["model"], "DeepSeek-V4.1-Flash")


if __name__ == "__main__":
    unittest.main()
