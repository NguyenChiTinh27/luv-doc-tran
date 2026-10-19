"""
tests/test_m1_adversarial.py
────────────────────────────
Adversarial Empirical Stress Tests for Milestone 1:
Full-Sentence AI Translation & Cache Engine.

Challenger M1-1 Test Suite.
Evaluates:
  1. Tricky chemical sentences and edge-case abbreviations:
     - "UN No. 1234", "EMS No. F-A, S-F", "CAS No. 1330-20-7"
     - "- UN No.", "- EMS No.", "- CAS No."
     - "- UN No.: 1263", "- EMS No.: F-E, S-E", "- CAS No.: 1330-20-7"
     - "can form explosive mixture", "international air transport association"
  2. Prohibited artifact emission under any circumstance:
     - Zero "(Thông tin)"
     - Zero "UN không có"
     - Zero "EMS không có"
     - Zero "CAS không có"
     - Zero "can Dạng"
  3. Quality Assertion Enforcement:
     - TranslationQualityError is genuinely raised for mock/invalid outputs
     - TranslationQualityError is genuinely raised for excessive untranslated English
  4. Codebase Scan:
     - Zero occurrences of "translate.googleapis.com" or "client=gtx" in codebase
  5. Cache integrity and thread-safety under concurrent stress.
"""

import os
import re
import sys
import json
import time
import shutil
import tempfile
import unittest
import threading
from concurrent.futures import ThreadPoolExecutor

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from src.engine.phase3_bilingual.full_sentence_translator import (
    FullSentenceTranslator,
    TranslationQualityError,
)
from src.engine.phase3_bilingual.rule_synthesizer import RuleBasedSynthesizer
from src.engine.phase3_bilingual.cost_guard import ZeroCostGuard
from src.engine.phase3_bilingual.offline_glossary import OfflineGlossary


class TestMilestone1Adversarial(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
        cls.cache_file = os.path.join(cls.project_root, "translation_cache.json")
        FullSentenceTranslator.set_cache_file(cls.cache_file)
        FullSentenceTranslator._ensure_cache_loaded()

    # ─────────────────────────────────────────────────────────────────────────
    # 1. TRICKY CHEMICAL SENTENCES & EDGE-CASE ABBREVIATIONS
    # ─────────────────────────────────────────────────────────────────────────
    def test_01_edge_case_abbreviations(self):
        """
        Adversarial test for abbreviations:
        'UN No. 1234', 'EMS No. F-A, S-F', 'CAS No. 1330-20-7', '- UN No.', '- EMS No.'.
        Must translate to Vietnamese equivalent and NEVER emit negative particles ('không có').
        """
        abbreviation_cases = [
            ("UN No. 1234", ["Số UN", "1234"]),
            ("EMS No. F-A, S-F", ["Số EMS", "F-A, S-F"]),
            ("CAS No. 1330-20-7", ["Số CAS", "1330-20-7"]),
            ("- UN No.", ["- Số UN"]),
            ("- EMS No.", ["- Số EMS"]),
            ("- CAS No.", ["- Số CAS"]),
            ("- UN No.: 1263", ["- Số UN:", "1263"]),
            ("- EMS No.: F-E, S-E", ["- Số EMS:", "F-E, S-E"]),
            ("- CAS No.: 1330-20-7", ["- Số CAS:", "1330-20-7"]),
            ("UN No.: 1234", ["Số UN:", "1234"]),
            ("EMS No.: F-A, S-F", ["Số EMS:", "F-A, S-F"]),
            ("CAS No.: 1330-20-7", ["Số CAS:", "1330-20-7"]),
            ("UN No", ["Số UN"]),
            ("EMS No", ["Số EMS"]),
            ("CAS No", ["Số CAS"]),
            ("IMO Class", ["Phân loại IMO"]),
            ("Package group", ["Nhóm đóng gói"]),
            ("Marine pollutants", ["Chất gây ô nhiễm biển"]),
            ("Packing method", ["Phương pháp đóng gói"]),
        ]

        with ZeroCostGuard():
            for src_text, expected_tokens in abbreviation_cases:
                # Test with RuleBasedSynthesizer
                res_rule = RuleBasedSynthesizer.synthesize(src_text)
                self.assertIsNotNone(res_rule, f"RuleBasedSynthesizer returned None for: {src_text}")
                for token in expected_tokens:
                    self.assertIn(
                        token,
                        res_rule,
                        f"Expected token '{token}' missing from RuleBasedSynthesizer output '{res_rule}' (Source: '{src_text}')"
                    )

                # Test with FullSentenceTranslator
                res_fst = FullSentenceTranslator.translate(src_text)
                self.assertIsNotNone(res_fst, f"FullSentenceTranslator returned None for: {src_text}")
                for token in expected_tokens:
                    self.assertIn(
                        token,
                        res_fst,
                        f"Expected token '{token}' missing from FullSentenceTranslator output '{res_fst}' (Source: '{src_text}')"
                    )

    def test_02_chemical_phrases(self):
        """
        Adversarial test for phrases that previously triggered pseudo-Vietnamese:
        'can form explosive mixture', 'international air transport association',
        'Other Special Transportation Protective Measures'.
        """
        phrase_cases = [
            ("can form explosive mixture", "có thể tạo thành hỗn hợp nổ"),
            ("International Air Transport Association", "Hiệp hội Vận tải Hàng không Quốc tế"),
            ("international air transport association", "Hiệp hội Vận tải Hàng không Quốc tế"),
            ("Other Special Transportation Protective Measures", "Các biện pháp bảo vệ đặc biệt khác khi vận chuyển"),
            ("According to regulation GB/T16483, GB/T17519", "Theo quy định GB/T16483, GB/T17519"),
            ("No smoking, drinking and eating.", "Không hút thuốc, ăn uống."),
            ("Trade name: Polyurethane Resin", "Tên thương phẩm: Nhựa Polyuretan"),
        ]

        with ZeroCostGuard():
            for src_text, expected_translation in phrase_cases:
                res_rule = RuleBasedSynthesizer.synthesize(src_text)
                self.assertEqual(
                    res_rule.strip().lower(),
                    expected_translation.strip().lower(),
                    f"RuleBasedSynthesizer mismatch for '{src_text}': got '{res_rule}', expected '{expected_translation}'"
                )

                res_fst = FullSentenceTranslator.translate(src_text)
                self.assertEqual(
                    res_fst.strip().lower(),
                    expected_translation.strip().lower(),
                    f"FullSentenceTranslator mismatch for '{src_text}': got '{res_fst}', expected '{expected_translation}'"
                )

    def test_03_multi_clause_and_label_value_inputs(self):
        """
        Adversarial test for multi-clause structured inputs containing colons and values:
        e.g., 'Product Code: ZP76011', 'Emergency telephone: +86-532-83889090', 'Flash point: 25 °C'.
        """
        structured_cases = [
            ("Product Code: ZP76011", "Mã sản phẩm: ZP76011"),
            ("Revision date: 2026-03-15", "Ngày hiệu đính: 2026-03-15"),
            ("CAS no.: 7631-86-9", "Số CAS: 7631-86-9"),
            ("Chemical formula: CaCO3", "Công thức hóa học: CaCO3"),
            ("Boiling point: > 1700 °C", "Điểm sôi: > 1700 °C"),
            ("Flash point: 25 °C", "Điểm chớp cháy: 25 °C"),
            ("Relative density: 1.5 (±0.05)", "Tỷ trọng tương đối: 1.5 (±0.05)"),
            ("Emergency telephone: +86-532-83889090", "Điện thoại khẩn cấp: +86-532-83889090"),
            ("Company name: Zhanchen Coating Co., Ltd.", "Tên công ty: Zhanchen Coating Co., Ltd."),
            ("Telephone: 0757-27332709", "Điện thoại: 0757-27332709"),
        ]

        with ZeroCostGuard():
            for src_text, expected_trans in structured_cases:
                res = FullSentenceTranslator.translate(src_text)
                self.assertEqual(
                    res.strip().lower(),
                    expected_trans.strip().lower(),
                    f"Structured line deconstruction failed for '{src_text}': got '{res}', expected '{expected_trans}'"
                )

    # ─────────────────────────────────────────────────────────────────────────
    # 2. ZERO PROHIBITED ARTIFACT EMISSION UNDER ANY CIRCUMSTANCE
    # ─────────────────────────────────────────────────────────────────────────
    def test_04_zero_prohibited_artifacts_emitted(self):
        """
        Verifies that zero prohibited mock artifacts are emitted by any translator:
        - (Thông tin)
        - UN không có
        - EMS không có
        - CAS không có
        - can Dạng
        """
        adversarial_inputs = [
            "UN No.",
            "UN No. 1234",
            "- UN No.",
            "- UN No.: 1263",
            "EMS No.",
            "EMS No. F-A, S-F",
            "- EMS No.",
            "- EMS No.: F-E, S-E",
            "CAS No.",
            "CAS No. 1330-20-7",
            "- CAS No.",
            "- CAS No.: 1330-20-7",
            "can form explosive mixture",
            "International Air Transport Association",
            "international air transport association",
            "Other Special Transportation Protective Measures",
            "No smoking, drinking and eating.",
            "Trade name: Polyurethane Resin",
            "Product Code: ZP76011",
            "Revision date: 2026-03-15",
            "Flash point: 25 °C",
            "Relative density: 1.5 (±0.05)",
            "Emergency telephone: +86-532-83889090",
            "MATERIAL SAFETY DATA SHEET",
        ]

        prohibited_substrings = [
            "(Thông tin)",
            "UN không có",
            "EMS không có",
            "CAS không có",
            "can Dạng",
        ]

        with ZeroCostGuard():
            for inp in adversarial_inputs:
                out_rule = RuleBasedSynthesizer.synthesize(inp)
                out_fst = FullSentenceTranslator.translate(inp)

                for bad in prohibited_substrings:
                    self.assertNotIn(
                        bad,
                        out_rule,
                        f"Prohibited artifact '{bad}' emitted by RuleBasedSynthesizer for input '{inp}': '{out_rule}'"
                    )
                    self.assertNotIn(
                        bad,
                        out_fst,
                        f"Prohibited artifact '{bad}' emitted by FullSentenceTranslator for input '{inp}': '{out_fst}'"
                    )

    # ─────────────────────────────────────────────────────────────────────────
    # 3. QUALITY ASSERTION ENFORCEMENT (TranslationQualityError)
    # ─────────────────────────────────────────────────────────────────────────
    def test_05_assert_valid_translation_catches_all_prohibited_strings(self):
        """
        Empirically verifies that FullSentenceTranslator.assert_valid_translation
        strictly raises TranslationQualityError when any prohibited artifact is present.
        """
        violations = [
            ("(Thông tin) Hỗn hợp nguy hiểm", "(Thông tin)"),
            ("Hỗn hợp (Thông tin)", "(Thông tin)"),
            ("Số UN không có 1234", "UN không có"),
            ("- UN không có", "UN không có"),
            ("Số EMS không có F-A", "EMS không có"),
            ("Số CAS không có 1330-20-7", "CAS không có"),
            ("Hơi can Dạng nổ", "can Dạng"),
        ]

        for bad_vi, expected_artifact in violations:
            with self.assertRaises(
                TranslationQualityError,
                msg=f"Failed to raise TranslationQualityError for output containing '{expected_artifact}': '{bad_vi}'"
            ):
                FullSentenceTranslator.assert_valid_translation("Adversarial Source", bad_vi)

    def test_06_assert_valid_translation_catches_excessive_english(self):
        """
        Empirically verifies that TranslationQualityError is raised when > 2 common
        untranslated English words appear in Vietnamese output.
        """
        excessive_english_outputs = [
            "Hỗn hợp này là water and flammable liquid",  # 4 common english words: water, and, flammable, liquid
            "Contains hazardous mixture for storage",   # 3 words: mixture, for, storage
            "causes irritation to respiratory system",   # 4 words: causes, irritation, respiratory, system
        ]

        for bad_vi in excessive_english_outputs:
            with self.assertRaises(
                TranslationQualityError,
                msg=f"Failed to raise TranslationQualityError for excessive English leak: '{bad_vi}'"
            ):
                FullSentenceTranslator.assert_valid_translation("Adversarial English Source", bad_vi)

    def test_07_assert_valid_translation_accepts_clean_translations(self):
        """
        Verifies that legitimate Vietnamese translations, formulas, numbers, and CAS codes
        pass assert_valid_translation without raising an error.
        """
        clean_outputs = [
            ("UN No.: 1234", "Số UN: 1234"),
            ("EMS No.: F-A, S-F", "Số EMS: F-A, S-F"),
            ("CAS No.: 1330-20-7", "Số CAS: 1330-20-7"),
            ("can form explosive mixture", "có thể tạo thành hỗn hợp nổ"),
            ("International Air Transport Association", "Hiệp hội Vận tải Hàng không Quốc tế"),
            ("1330-20-7", "1330-20-7"),
            ("CaCO3", "CaCO3"),
            ("> 1700 °C", "> 1700 °C"),
            ("15 minutes", "15 phút"),
            ("No smoking, drinking and eating.", "Không hút thuốc, ăn uống."),
        ]

        for src, clean_vi in clean_outputs:
            try:
                FullSentenceTranslator.assert_valid_translation(src, clean_vi)
            except TranslationQualityError as e:
                self.fail(f"Legitimate translation '{clean_vi}' incorrectly rejected: {e}")

    # ─────────────────────────────────────────────────────────────────────────
    # 4. CODEBASE SCAN: OCCURRENCES OF GOOGLE GTX OR GOOGLEAPIS
    # ─────────────────────────────────────────────────────────────────────────
    def test_08_codebase_scan_google_gtx(self):
        """
        Empirically scans the codebase to verify:
        1. src/engine/ contains strictly 0 occurrences of 'translate.googleapis.com' or 'client=gtx'.
        2. Detects whether legacy files (such as core/interlinear_translator.py) contain occurrences.
        """
        src_engine_dir = os.path.join(self.project_root, "src", "engine")
        engine_violations = []

        for root, dirs, files in os.walk(src_engine_dir):
            for file in files:
                if file.endswith(".py"):
                    fpath = os.path.join(root, file)
                    with open(fpath, "r", encoding="utf-8", errors="ignore") as f:
                        for lnum, line in enumerate(f, 1):
                            if "translate.googleapis.com" in line or "client=gtx" in line:
                                engine_violations.append((fpath, lnum, line.strip()))

        self.assertEqual(
            len(engine_violations),
            0,
            f"Prohibited Google GTX endpoint found in src/engine: {engine_violations}"
        )

        # Scan Production source files in src/, core/, and server.py
        target_sources = []
        for d in ["src", "core"]:
            p = os.path.join(self.project_root, d)
            for root, _, files in os.walk(p):
                for f in files:
                    if f.endswith(".py"):
                        target_sources.append(os.path.join(root, f))
        server_py = os.path.join(self.project_root, "server.py")
        if os.path.exists(server_py):
            target_sources.append(server_py)

        codebase_violations = []
        for fpath in target_sources:
            rel_path = os.path.relpath(fpath, self.project_root)
            with open(fpath, "r", encoding="utf-8", errors="ignore") as f:
                for lnum, line in enumerate(f, 1):
                    if "translate.googleapis.com" in line or "client=gtx" in line:
                        codebase_violations.append((rel_path, lnum, line.strip()))

        # Note: reports any occurrences across production codebase
        self.codebase_gtx_violations = codebase_violations
        self.assertEqual(
            len(codebase_violations),
            0,
            f"Prohibited Google GTX endpoint found in production codebase: {codebase_violations}"
        )

    # ─────────────────────────────────────────────────────────────────────────
    # 5. CACHE INTEGRITY & CONCURRENT THREAD-SAFETY
    # ─────────────────────────────────────────────────────────────────────────
    def test_09_concurrent_cache_reads_and_writes(self):
        """
        Stress-tests FullSentenceTranslator under high concurrency (20 threads)
        verifying thread-safety, atomic persistence, and zero JSON corruption.
        """
        temp_dir = tempfile.mkdtemp(prefix="tcache_test_")
        temp_cache = os.path.join(temp_dir, "test_concurrent_cache.json")

        try:
            # Seed cache
            initial_data = {
                "UN No.": "Số UN",
                "EMS No.": "Số EMS",
                "CAS No.": "Số CAS",
                "can form explosive mixture": "có thể tạo thành hỗn hợp nổ",
            }
            with open(temp_cache, "w", encoding="utf-8") as f:
                json.dump(initial_data, f, ensure_ascii=False, indent=2)

            FullSentenceTranslator.set_cache_file(temp_cache)
            FullSentenceTranslator._ensure_cache_loaded()

            errors = []

            def worker_task(worker_id: int):
                try:
                    for i in range(15):
                        # Mixed reads and writes
                        key = f"Concurrent_Key_{worker_id}_{i}"
                        val = f"Bản dịch đồng thời {worker_id}_{i}"

                        # Read existing standard
                        r1 = FullSentenceTranslator.translate("UN No.")
                        if r1 != "Số UN":
                            errors.append(f"Worker {worker_id} read wrong value: {r1}")

                        # Write new key
                        with FullSentenceTranslator._lock:
                            FullSentenceTranslator._cache[key] = val
                        FullSentenceTranslator._save_cache()

                        # Read back
                        with FullSentenceTranslator._lock:
                            read_val = FullSentenceTranslator._cache.get(key)
                        if read_val != val:
                            errors.append(f"Worker {worker_id} write-read mismatch: {read_val} != {val}")
                except Exception as e:
                    errors.append(f"Worker {worker_id} exception: {str(e)}")

            threads = []
            for tid in range(20):
                t = threading.Thread(target=worker_task, args=(tid,))
                threads.append(t)
                t.start()

            for t in threads:
                t.join(timeout=10)

            self.assertEqual(len(errors), 0, f"Concurrent cache errors: {errors}")

            # Verify file on disk is valid JSON and contains all writes
            self.assertTrue(os.path.exists(temp_cache))
            with open(temp_cache, "r", encoding="utf-8") as f:
                disk_data = json.load(f)

            self.assertIn("UN No.", disk_data)
            self.assertEqual(disk_data["UN No."], "Số UN")
            # Verify written keys count: 20 workers * 15 items = 300 items + 4 initial
            self.assertGreaterEqual(len(disk_data), 300)

        finally:
            # Restore default cache
            FullSentenceTranslator.set_cache_file(self.cache_file)
            shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
