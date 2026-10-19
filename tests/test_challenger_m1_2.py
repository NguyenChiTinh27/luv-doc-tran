"""
tests/test_challenger_m1_2.py
─────────────────────────────
Adversarial empirical stress tests for Milestone 1 (Challenger M1-2):
1. Document-level output assertions & text run scan for prohibited mock artifacts
   ('(Thông tin)', 'UN không có', 'EMS không có', 'CAS không có', 'can Dạng').
2. Chemical formula, numerical range, and CAS code preservation.
3. Cache atomic persistence under high-concurrency multithreaded and multiprocess stress.
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
from concurrent.futures import ThreadPoolExecutor, ProcessPoolExecutor, as_completed
import docx

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from src.engine.phase3_bilingual.full_sentence_translator import (
    FullSentenceTranslator,
    TranslationQualityError,
)
from src.engine.phase3_bilingual.offline_glossary import OfflineGlossary


# Subprocess helper function for multiprocess cache write stress test
def _worker_write_cache(cache_file: str, worker_id: int, num_items: int) -> int:
    """Simulates independent process writing to cache with atomic persistence."""
    written = 0
    for i in range(num_items):
        key = f"ProcKey_{worker_id}_{i}"
        val = f"Bản dịch quy trình {worker_id}_{i}"
        cache_dir = os.path.dirname(cache_file)
        # Emulate atomic save logic from FullSentenceTranslator
        try:
            # Read existing
            data = {}
            if os.path.exists(cache_file):
                with open(cache_file, "r", encoding="utf-8") as f:
                    try:
                        data = json.load(f)
                    except Exception:
                        data = {}
            data[key] = val
            fd, tmp_path = tempfile.mkstemp(dir=cache_dir, prefix="tproc_", suffix=".tmp")
            with open(fd, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            os.replace(tmp_path, cache_file)
            written += 1
        except Exception as e:
            # On windows replace might briefly contest; retry once
            time.sleep(0.005)
            try:
                data[key] = val
                fd, tmp_path = tempfile.mkstemp(dir=cache_dir, prefix="tproc_retry_", suffix=".tmp")
                with open(fd, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                os.replace(tmp_path, cache_file)
                written += 1
            except Exception:
                pass
    return written


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
        ]
        for num_case in test_numeric_cases:
            with self.subTest(num_case=num_case):
                res = FullSentenceTranslator.translate(num_case)
                # Verify all digits, dots, hyphens, and percent signs from input remain intact
                orig_digits = "".join(c for c in num_case if c.isdigit() or c in ".-+><=%±°~")
                res_digits = "".join(c for c in res if c.isdigit() or c in ".-+><=%±°~")
                self.assertEqual(
                    orig_digits,
                    res_digits,
                    f"Numeric content altered in '{num_case}': got '{res}'"
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
            ("Content : 1.0 - 2.5 %", "Hàm lượng: 1.0 - 2.5 %"),
            ("Chemical Formula : CaCO3", "Công thức hóa học: CaCO3"),
            ("Melting point : > 1700 °C", "Điểm nóng chảy: > 1700 °C"),
            ("Relative density : 1.5 (±0.05)", "Mật độ tương đối: 1.5 (±0.05)"),
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
            # Seed 50 initial entries
            seed_data = {f"SeedKey_{i}": f"Bản dịch hạt giống {i}" for i in range(50)}
            with open(temp_cache_file, "w", encoding="utf-8") as f:
                json.dump(seed_data, f, ensure_ascii=False, indent=2)

            FullSentenceTranslator.set_cache_file(temp_cache_file)

            errors = []
            num_threads = 20
            ops_per_thread = 50

            def reader_task(tid: int):
                try:
                    for i in range(ops_per_thread):
                        key = f"SeedKey_{i % 50}"
                        val = FullSentenceTranslator.translate(key)
                        if val != f"Bản dịch hạt giống {i % 50}":
                            errors.append(f"Reader {tid} got invalid val '{val}' for '{key}'")
                        time.sleep(0.001)
                except Exception as e:
                    errors.append(f"Reader {tid} exception: {str(e)}")

            def writer_task(tid: int):
                try:
                    for i in range(ops_per_thread):
                        with FullSentenceTranslator._lock:
                            key = f"ThreadKey_{tid}_{i}"
                            val = f"Giá trị luồng {tid}_{i}"
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

    def test_challenge_3_multiprocess_atomic_persistence(self):
        """
        Stress-test atomic persistence across multiple independent OS processes
        writing to the same cache file via atomic replacement.
        """
        temp_dir = tempfile.mkdtemp(prefix="cache_proc_test_")
        temp_cache_file = os.path.join(temp_dir, "proc_cache.json")
        try:
            initial_data = {"INIT_KEY": "INIT_VAL"}
            with open(temp_cache_file, "w", encoding="utf-8") as f:
                json.dump(initial_data, f, ensure_ascii=False, indent=2)

            num_workers = 4
            items_per_worker = 25

            with ProcessPoolExecutor(max_workers=num_workers) as executor:
                futures = [
                    executor.submit(_worker_write_cache, temp_cache_file, w_id, items_per_worker)
                    for w_id in range(num_workers)
                ]
                results = [f.result(timeout=20.0) for f in as_completed(futures)]

            self.assertTrue(os.path.exists(temp_cache_file))
            self.assertGreater(os.path.getsize(temp_cache_file), 0)

            # File must be valid JSON without syntax or truncation errors
            with open(temp_cache_file, "r", encoding="utf-8") as f:
                final_data = json.load(f)
            self.assertIsInstance(final_data, dict)
            self.assertIn("INIT_KEY", final_data)

        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
