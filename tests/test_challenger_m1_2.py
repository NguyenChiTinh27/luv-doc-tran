"""
tests/test_challenger_m1_2.py
─────────────────────────────
Adversarial empirical stress tests for Milestone 1 (Challenger M1-2):
1. Document-level output assertions & run-level scan for prohibited mock artifacts
   ('(Thông tin)', 'UN không có', 'EMS không có', 'CAS không có', 'can Dạng').
2. Chemical formula, numerical range, and CAS code preservation.
3. Cache atomic persistence under high-concurrency multithreaded contention.
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
import docx

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from src.engine.phase3_bilingual.full_sentence_translator import (
    FullSentenceTranslator,
    TranslationQualityError,
)
from tests.audit_verifier import verify_bilingual_docx


class TestChallengerM1_2(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
        cls.default_cache = os.path.join(cls.project_root, "translation_cache.json")
        FullSentenceTranslator.set_cache_file(cls.default_cache)

    # ─────────────────────────────────────────────────────────────────────────
    # CHALLENGE 1: Document-Level Output Assertions & Run-Level Prohibited Strings
    # ─────────────────────────────────────────────────────────────────────────
    def test_challenge_1_document_run_scan_for_prohibited_strings(self):
        """
        Scan all text runs, paragraphs, table cells, and headers in generated DOCX
        to ensure zero occurrences of prohibited mock strings or pseudo-Vietnamese.
        """
        output_docx = os.path.join(self.project_root, "tests", "output_verified_1111111.docx")
        self.assertTrue(os.path.exists(output_docx), f"Target output docx does not exist: {output_docx}")

        doc = docx.Document(output_docx)
        prohibited_regexes = [
            (re.compile(r"\(Thông\s*tin\)", re.IGNORECASE), "(Thông tin)"),
            (re.compile(r"\bUN\s*không\s*có\b", re.IGNORECASE), "UN không có"),
            (re.compile(r"\bEMS\s*không\s*có\b", re.IGNORECASE), "EMS không có"),
            (re.compile(r"\bCAS\s*không\s*có\b", re.IGNORECASE), "CAS không có"),
            (re.compile(r"\bcan\s*Dạng\b", re.IGNORECASE), "can Dạng"),
        ]

        found_violations = []

        def check_text(location_desc: str, text: str):
            for regex, label in prohibited_regexes:
                if regex.search(text):
                    found_violations.append(f"Found '{label}' at {location_desc}: '{text[:100]}'")

        # 1. Check sections (headers & footers)
        for s_idx, sec in enumerate(doc.sections):
            for p_idx, p in enumerate(sec.header.paragraphs):
                check_text(f"Section {s_idx} Header P{p_idx}", p.text)
                for r_idx, r in enumerate(p.runs):
                    check_text(f"Section {s_idx} Header P{p_idx} R{r_idx}", r.text)
            for p_idx, p in enumerate(sec.footer.paragraphs):
                check_text(f"Section {s_idx} Footer P{p_idx}", p.text)
                for r_idx, r in enumerate(p.runs):
                    check_text(f"Section {s_idx} Footer P{p_idx} R{r_idx}", r.text)

        # 2. Check body paragraphs
        for p_idx, p in enumerate(doc.paragraphs):
            check_text(f"Body P{p_idx}", p.text)
            for r_idx, r in enumerate(p.runs):
                check_text(f"Body P{p_idx} R{r_idx}", r.text)

        # 3. Check tables
        for t_idx, tbl in enumerate(doc.tables):
            for r_idx, row in enumerate(tbl.rows):
                for c_idx, cell in enumerate(row.cells):
                    check_text(f"Table {t_idx} R{r_idx} C{c_idx}", cell.text)
                    for p_i, cp in enumerate(cell.paragraphs):
                        for cr_i, cr in enumerate(cp.runs):
                            check_text(f"Table {t_idx} R{r_idx} C{c_idx} P{p_i} R{cr_i}", cr.text)

        self.assertEqual(
            len(found_violations),
            0,
            f"Adversarial scan detected {len(found_violations)} prohibited mock string violations:\n"
            + "\n".join(found_violations[:10])
        )

        # Check body text does not contain recurring header table/p
        body_headers = [
            p.text for p in doc.paragraphs
            if "MATERIAL SAFETY DATA SHEET" in p.text.upper() and "GB/T16483" in p.text
        ]
        self.assertEqual(len(body_headers), 0, "Found recurring header block in body paragraphs!")

    def test_challenge_1_audit_verifier_check_zero_mock_artifacts(self):
        """Run full 8-point audit verifier directly on tests/output_verified_1111111.docx."""
        output_docx = os.path.join(self.project_root, "tests", "output_verified_1111111.docx")
        rep = verify_bilingual_docx(output_docx)
        self.assertEqual(
            rep.get("mock_artifact_violations", -1),
            0,
            f"Audit verifier Check 8 failed with {rep.get('mock_artifact_violations')} mock violations!"
        )

    # ─────────────────────────────────────────────────────────────────────────
    # CHALLENGE 2: Chemical Formula, Numerical Range, and CAS Preservation
    # ─────────────────────────────────────────────────────────────────────────
    def test_challenge_2_chemical_formula_preservation(self):
        """
        Verify that FullSentenceTranslator preserves standard chemical formulas
        (TiO2, H2SO4, CaCO3, SiO2, etc.) without altering subscripts, numbers, or letters.
        """
        test_formulas = [
            "TiO2",
            "H2SO4",
            "CaCO3",
            "SiO2",
            "Al2O3",
            "Fe2O3",
            "H2O",
            "NaOH",
            "NaCl",
            "KCl",
            "CaCl2",
            "MgSO4",
            "CuSO4",
            "HCl",
            "HNO3",
            "H3PO4",
            "BaSO4",
            "ZnO",
            "NH3",
            "CH3COOH",
        ]
        for formula in test_formulas:
            with self.subTest(formula=formula):
                res = FullSentenceTranslator.translate(formula)
                self.assertEqual(
                    res,
                    formula,
                    f"Formula '{formula}' was altered during translation: '{res}'"
                )

        # Compound formula model codes
        model_codes = ["TiO2-99", "ZP76011", "YTU0035", "FU43288"]
        for code in model_codes:
            with self.subTest(code=code):
                res = FullSentenceTranslator.translate(code)
                self.assertEqual(
                    res,
                    code,
                    f"Model code '{code}' was altered during translation: '{res}'"
                )

    def test_challenge_2_numerical_range_preservation(self):
        """
        Verify that numerical ranges, percentages, and units are preserved with 100% precision.
        """
        test_numeric_cases = [
            "1.0 - 2.5",
            "1.0-2.5",
            "1.0 - 2.5 %",
            "10 - 25%",
            "> 1700 °C",
            "< 0.1 %",
            "1.5 (±0.05)",
            ": 1.0 - 2.5",
            "0.1 - 1.0",
            "20 - 30 %",
            "~ 100 °C",
            ">= 99.5%",
            "<= 0.01%",
            "15 minutes",
            "2 hours",
        ]
        for num_case in test_numeric_cases:
            with self.subTest(num_case=num_case):
                res = FullSentenceTranslator.translate(num_case)
                # Verify all digits from input remain intact
                orig_digits = "".join(c for c in num_case if c.isdigit())
                res_digits = "".join(c for c in res if c.isdigit())
                self.assertEqual(
                    orig_digits,
                    res_digits,
                    f"Numeric content lost in '{num_case}': got '{res}'"
                )

    def test_challenge_2_cas_codes_preservation(self):
        """
        Verify that CAS registry numbers are preserved with exact hyphenation and zero number loss.
        """
        test_cas_cases = [
            ("13463-67-7", "13463-67-7"),
            ("7664-93-9", "7664-93-9"),
            ("471-34-1", "471-34-1"),
            ("7732-18-5", "7732-18-5"),
            ("14808-60-7", "14808-60-7"),
            ("CAS No.: 13463-67-7", "Số CAS: 13463-67-7"),
            ("- CAS No.: 7664-93-9", "- Số CAS: 7664-93-9"),
            ("CAS No: 471-34-1", "Số CAS: 471-34-1"),
            ("- CAS No: 7732-18-5", "- Số CAS: 7732-18-5"),
        ]
        for inp, expected in test_cas_cases:
            with self.subTest(inp=inp):
                res = FullSentenceTranslator.translate(inp)
                # In all cases the CAS number must appear verbatim in the output
                cas_number = re.search(r"\b\d{2,7}-\d{2}-\d\b", inp).group(0)
                self.assertIn(
                    cas_number,
                    res,
                    f"CAS number '{cas_number}' lost or corrupted in translation '{res}'"
                )
                if expected:
                    self.assertEqual(res, expected)

    def test_challenge_2_label_value_with_formulas_and_ranges(self):
        """
        Verify label-value structures containing formulas, CAS numbers, and ranges.
        """
        cases = [
            ("CAS No. : 13463-67-7", "Số CAS: 13463-67-7"),
            ("Chemical formula : CaCO3", "Công thức hóa học: CaCO3"),
            ("Chemical formula : TiO2", "Công thức hóa học: TiO2"),
            ("Melting point : > 1700 °C", "Điểm nóng chảy: > 1700 °C"),
            ("Boiling point : > 1700 °C", "Điểm sôi: > 1700 °C"),
        ]
        for inp, expected in cases:
            with self.subTest(inp=inp):
                res = FullSentenceTranslator.translate(inp)
                self.assertEqual(res, expected)

    # ─────────────────────────────────────────────────────────────────────────
    # CHALLENGE 3: Cache Atomic Persistence Under High-Concurrency Contention
    # ─────────────────────────────────────────────────────────────────────────
    def test_challenge_3_multithreaded_cache_concurrency(self):
        """
        Stress-test FullSentenceTranslator cache under 20 concurrent threads:
        - 10 threads reading known & newly added keys
        - 10 threads writing novel translations concurrently
        Verifies:
        - Zero JSONDecodeError
        - Zero file truncation (size > 0 bytes always)
        - Thread safety with 100% data integrity
        """
        temp_dir = tempfile.mkdtemp(prefix="cache_test_")
        temp_cache_file = os.path.join(temp_dir, "test_cache.json")
        try:
            # Seed 50 initial phrase entries (with spaces so they exercise cache lookup)
            seed_data = {f"Phrase Entry Number {i}": f"Bản dịch câu mẫu số {i}" for i in range(50)}
            with open(temp_cache_file, "w", encoding="utf-8") as f:
                json.dump(seed_data, f, ensure_ascii=False, indent=2)

            FullSentenceTranslator.set_cache_file(temp_cache_file)

            errors = []
            num_threads = 20
            ops_per_thread = 40

            def reader_task(tid: int):
                try:
                    for i in range(ops_per_thread):
                        key = f"Phrase Entry Number {i % 50}"
                        val = FullSentenceTranslator.translate(key)
                        if val != f"Bản dịch câu mẫu số {i % 50}":
                            errors.append(f"Reader {tid} got invalid val '{val}' for '{key}'")
                        time.sleep(0.001)
                except Exception as e:
                    errors.append(f"Reader {tid} exception: {str(e)}")

            def writer_task(tid: int):
                try:
                    for i in range(ops_per_thread):
                        with FullSentenceTranslator._lock:
                            key = f"Thread Novel Phrase {tid} Iteration {i}"
                            val = f"Giá trị mới tạo luồng {tid} lần {i}"
                            FullSentenceTranslator._cache[key] = val
                            FullSentenceTranslator._save_cache()
                        time.sleep(0.001)
                except Exception as e:
                    errors.append(f"Writer {tid} exception: {str(e)}")

            threads = []
            for t_i in range(10):
                threads.append(threading.Thread(target=reader_task, args=(t_i,)))
            for t_i in range(10, num_threads):
                threads.append(threading.Thread(target=writer_task, args=(t_i,)))

            for t in threads:
                t.start()
            for t in threads:
                t.join(timeout=15.0)

            # Check that file exists and is not truncated
            self.assertTrue(os.path.exists(temp_cache_file))
            file_size = os.path.getsize(temp_cache_file)
            self.assertGreater(file_size, 0, "Cache file was truncated to 0 bytes!")

            # Verify JSON file validity
            with open(temp_cache_file, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            self.assertIsInstance(loaded, dict)
            self.assertGreaterEqual(len(loaded), 50)

            self.assertEqual(len(errors), 0, f"Multithreaded cache stress test encountered errors: {errors}")

        finally:
            FullSentenceTranslator.set_cache_file(self.default_cache)
            shutil.rmtree(temp_dir, ignore_errors=True)

    def test_challenge_3_atomic_persistence_file_integrity_during_rapid_writes(self):
        """
        Verify that external readers reading translation_cache.json directly while
        _save_cache() is in high contention NEVER encounter a partially written or empty file.
        """
        temp_dir = tempfile.mkdtemp(prefix="cache_atomic_test_")
        temp_cache_file = os.path.join(temp_dir, "atomic_cache.json")
        try:
            initial_data = {f"Key_{i}": f"Value_{i}" * 50 for i in range(100)}
            with open(temp_cache_file, "w", encoding="utf-8") as f:
                json.dump(initial_data, f, ensure_ascii=False, indent=2)

            FullSentenceTranslator.set_cache_file(temp_cache_file)

            stop_event = threading.Event()
            read_errors = []

            def disk_reader():
                while not stop_event.is_set():
                    try:
                        sz = os.path.getsize(temp_cache_file)
                        if sz == 0:
                            read_errors.append("Read file size 0 (truncated!)")
                        with open(temp_cache_file, "r", encoding="utf-8") as f:
                            data = json.load(f)
                            if not isinstance(data, dict) or len(data) < 100:
                                read_errors.append(f"Incomplete data read: len={len(data)}")
                    except Exception as e:
                        read_errors.append(f"Direct disk reader error: {type(e).__name__}: {str(e)}")
                    time.sleep(0.0005)

            reader_threads = [threading.Thread(target=disk_reader) for _ in range(5)]
            for rt in reader_threads:
                rt.start()

            # Execute rapid writes in main thread
            for w in range(50):
                with FullSentenceTranslator._lock:
                    FullSentenceTranslator._cache[f"RapidKey_{w}"] = f"RapidVal_{w}"
                    FullSentenceTranslator._save_cache()
                time.sleep(0.002)

            stop_event.set()
            for rt in reader_threads:
                rt.join(timeout=5.0)

            self.assertEqual(
                len(read_errors),
                0,
                f"Atomic persistence failed! Encountered {len(read_errors)} read corruptions during concurrent save:\n"
                + "\n".join(read_errors[:10])
            )

        finally:
            FullSentenceTranslator.set_cache_file(self.default_cache)
            shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
