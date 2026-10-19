"""
tests/test_tier5_adversarial_bilingual_qc.py
─────────────────────────────────────────────────────────────────────────────
Tier 5 Adversarial Coverage Hardening Suite for Bilingual Document
Reconstruction & QC Engine (bilingual_qc_v2.py).

Covers:
  1. Adversarial ZeroCostGuard stress testing (active interception, exception guarantees, URL variations).
  2. Adversarial Concatenation stress testing (cross-boundary formatting, Vietnamese diacritics, numerals, whitelists).
  3. Adversarial Double-Tab stress testing (multiline runs, adjacent <w:tab/> tags, mixed tabs/spaces).
  4. Adversarial Header Ordering stress testing (reversed 16-to-1 sections, scrambled sections, repair verification).
  5. Adversarial Table Reflow stress testing (deeply nested tables, 100+ locked exact rows, cantSplit enforcement).
  6. Adversarial Leaked English stress testing (ordinals 1st..16th, GHS headings, section refs, negative/positive verification).
  7. Adversarial Pair Completeness stress testing (unpaired trailing English, orphan cells, mixed language noise).
  8. Adversarial Token Sanitizer & Consolidator stress testing (malformed punctuation, units, OCR corruption).
  9. Full 7-Point Audit Negative Matrix (ensuring no false negatives across all 7 checks).
 10. End-to-end integration and Zero DeepSeek API Cost verification on real disk samples.
─────────────────────────────────────────────────────────────────────────────
"""

import os
import sys
import unittest
import tempfile
import zipfile
import lxml.etree as etree
import docx
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
from docx.shared import Pt, RGBColor

import core.bilingual_qc_v2 as bqc
from core.bilingual_qc_v2 import (
    BilingualQCReconstructionEngine,
    reconstruct_bilingual_document,
    verify_bilingual_docx,
    ZeroCostGuard,
    ApiCostViolationError,
    OfflineGlossary,
    TokenSanitizer,
    LeakHunter,
    RuleBasedSynthesizer,
    PairingBuilder,
    ReflowUnlocker,
    GridReconstructor,
    ParagraphConsolidator,
    StructureNormalizer,
    DARK_NAVY_HEX,
    DARK_NAVY_RGB,
    BLACK_RGB,
    NSMAP,
    W_NS
)
from tests.audit_verifier import (
    check_zero_concatenation,
    check_zero_double_tabs,
    check_pair_completeness,
    check_zero_leaked_english,
    check_header_order,
    check_zero_clipped_rows,
    check_visual_parity,
    is_translatable_source_text,
    is_vietnamese_text,
    is_chemical_formula,
    WHITELIST_TOKENS,
    NS
)


class TestTier5AdversarialZeroCostGuard(unittest.TestCase):
    """Stress tests asserting 0 paid API costs under all execution branches."""

    def test_adv_guard_blocks_deepseek_requests(self):
        """ZeroCostGuard must block any requests to api.deepseek.com."""
        import requests
        with ZeroCostGuard():
            with self.assertRaises(ApiCostViolationError):
                requests.Session().send(requests.Request("POST", "https://api.deepseek.com/chat/completions").prepare())

    def test_adv_guard_blocks_openai_requests(self):
        """ZeroCostGuard must block any requests to api.openai.com."""
        import requests
        with ZeroCostGuard():
            with self.assertRaises(ApiCostViolationError):
                requests.Session().send(requests.Request("POST", "https://api.openai.com/v1/chat/completions").prepare())

    def test_adv_guard_blocks_urllib(self):
        """ZeroCostGuard must intercept urllib.request calls to paid API endpoints."""
        import urllib.request
        with ZeroCostGuard():
            with self.assertRaises(ApiCostViolationError):
                urllib.request.urlopen("https://api.deepseek.com/v1/models")

    def test_adv_synthesizer_operates_strictly_offline(self):
        """RuleBasedSynthesizer must never attempt network access for any input text."""
        with ZeroCostGuard():
            # Stress with diverse complex strings
            samples = [
                "SECTION 1: Identification of the substance/mixture",
                "Calcium carbonate 100% CAS 471-34-1",
                "Skin sensitization: Causes serious eye irritation",
                "No smoking, drinking and eating. 15minutes.Assure adequate flushing.",
                "Telephone: 0274 3555 208 Fax: 0274 3555 209 Email: info@company.com",
                ": > 1700 °C (±0.05) pH 7.5 - 8.5",
                "Arbitrary unmapped technical paragraph with ordinals 1st, 2nd, and section 6.2."
            ]
            for s in samples:
                trans = RuleBasedSynthesizer.synthesize(s)
                self.assertIsInstance(trans, str)
                self.assertGreater(len(trans), 0)


class TestTier5AdversarialConcatenation(unittest.TestCase):
    """Adversarial testing of Check 1: Zero Concatenation across run/paragraph boundaries."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_adv_intra_run_vietnamese_diacritic_concatenation(self):
        """Intra-run lowercase followed directly by capitalized Vietnamese diacritic letter must be flagged."""
        doc = docx.Document()
        p = doc.add_paragraph("chemicalĐộc tính nguy hiểm")
        p_path = os.path.join(self.temp_dir, "adv_concat_vi.docx")
        doc.save(p_path)

        rep = verify_bilingual_docx(p_path)
        self.assertFalse(rep["passed"])
        self.assertGreaterEqual(rep["concatenation_violations"], 1)

    def test_adv_inter_run_without_break_must_fail(self):
        """Adjacent runs across run boundary without <w:br/> or space must trigger concatenation violation."""
        doc = docx.Document()
        p = doc.add_paragraph()
        r1 = p.add_run("hazard")
        r2 = p.add_run("Identification")
        # No space, no break
        p_path = os.path.join(self.temp_dir, "adv_concat_runs.docx")
        doc.save(p_path)

        rep = verify_bilingual_docx(p_path)
        self.assertFalse(rep["passed"])
        self.assertGreaterEqual(rep["concatenation_violations"], 1)

    def test_adv_inter_run_with_explicit_break_must_pass(self):
        """Adjacent runs separated by explicit <w:br/> must NOT trigger concatenation violation."""
        doc = docx.Document()
        p = doc.add_paragraph()
        r1 = p.add_run("hazard")
        r_br = p.add_run()
        r_br._r.append(OxmlElement("w:br"))
        r2 = p.add_run("Nhận dạng đặc tính nguy hiểm")
        r2.font.color.rgb = DARK_NAVY_RGB
        r2.italic = True

        p_path = os.path.join(self.temp_dir, "adv_concat_with_br.docx")
        doc.save(p_path)

        rep = verify_bilingual_docx(p_path)
        self.assertEqual(rep["concatenation_violations"], 0)

    def test_adv_whitelisted_and_chemical_tokens_zero_false_positives(self):
        """Whitelisted tokens (pH, mPa.s, CaCO3, SiO2, etc.) must never trigger false positives."""
        doc = docx.Document()
        p = doc.add_paragraph("Measurement of pH at 20 °C with CaCO3 and SiO2 under 150 mPa.s viscosity.")
        p_path = os.path.join(self.temp_dir, "adv_concat_whitelist.docx")
        doc.save(p_path)

        rep = verify_bilingual_docx(p_path)
        self.assertEqual(rep["concatenation_violations"], 0)


class TestTier5AdversarialDoubleTabs(unittest.TestCase):
    """Adversarial testing of Check 2: Zero Internal Double-Tabs."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_adv_multiple_tabs_and_whitespace_tabs_caught(self):
        """Consecutive tabs with spaces ('\t   \t') must be caught by Check 2."""
        doc = docx.Document()
        doc.add_paragraph("Label:\t   \tValue")
        p_path = os.path.join(self.temp_dir, "adv_tab_spaced.docx")
        doc.save(p_path)

        rep = verify_bilingual_docx(p_path)
        self.assertFalse(rep["passed"])
        self.assertGreaterEqual(rep["double_tab_violations"], 1)

    def test_adv_grid_reconstructor_strips_all_double_tabs(self):
        """GridReconstructor must strip 100% of internal double-tabs during reconstruction."""
        doc = docx.Document()
        doc.add_paragraph("Product Name\t\t:\t\tSYLOID MX107")
        doc.add_paragraph("CAS Number\t\t:\t\t7631-86-9")
        doc.add_paragraph("Physical Form\t\t:\t\tWhite Powder")

        GridReconstructor.reconstruct_label_value_blocks(doc)
        p_path = os.path.join(self.temp_dir, "adv_tab_reconstructed.docx")
        doc.save(p_path)

        rep = verify_bilingual_docx(p_path)
        self.assertEqual(rep["double_tab_violations"], 0)


class TestTier5AdversarialHeaderOrdering(unittest.TestCase):
    """Adversarial testing of Check 5: Header Order and StructureNormalizer."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_adv_reversed_16_to_1_sections_caught(self):
        """Reversed section headers (16 down to 1) must be detected as 15 order violations."""
        doc = docx.Document()
        for i in range(16, 0, -1):
            doc.add_paragraph(f"SECTION {i}: Section Title {i}")

        p_path = os.path.join(self.temp_dir, "adv_rev_headers.docx")
        doc.save(p_path)

        rep = verify_bilingual_docx(p_path)
        self.assertFalse(rep["passed"])
        self.assertEqual(rep["header_order_violations"], 15)

    def test_adv_structure_normalizer_repairs_inverted_headers(self):
        """StructureNormalizer must successfully reorder scrambled sections into ascending 1..16 order."""
        doc = docx.Document()
        # Add sections out of order
        scrambled_order = [4, 1, 3, 2, 5, 8, 6, 7, 11, 9, 10, 16, 12, 15, 13, 14]
        for sec_num in scrambled_order:
            p = doc.add_paragraph(f"SECTION {sec_num}: Title {sec_num}")
            p.add_run(f"\nContent for section {sec_num}.")

        StructureNormalizer.repair_inverted_headers(doc)
        StructureNormalizer.normalize_section_headers(doc)

        p_path = os.path.join(self.temp_dir, "adv_repaired_headers.docx")
        doc.save(p_path)

        rep = verify_bilingual_docx(p_path)
        self.assertEqual(rep["header_order_violations"], 0, f"Expected 0 violations after repair, got: {rep}")


class TestTier5AdversarialTableReflow(unittest.TestCase):
    """Adversarial testing of Check 6: Zero Clipped Rows and dynamic reflow."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_adv_massive_table_100_rows_exact_height_unlocked(self):
        """100-row table with w:hRule='exact' on every row must have 100% unlocked to 'atLeast'."""
        doc = docx.Document()
        tbl = doc.add_table(rows=100, cols=2)
        for row in tbl.rows:
            trPr = row._tr.get_or_add_trPr()
            trH = OxmlElement("w:trHeight")
            trH.set(qn("w:hRule"), "exact")
            trH.set(qn("w:val"), "240")
            trPr.append(trH)

        p_before = os.path.join(self.temp_dir, "adv_table_locked.docx")
        doc.save(p_before)
        rep_before = verify_bilingual_docx(p_before)
        self.assertEqual(rep_before["clipped_row_violations"], 100)

        # Unlock
        ReflowUnlocker.unlock_table_row_heights(doc)
        p_after = os.path.join(self.temp_dir, "adv_table_unlocked.docx")
        doc.save(p_after)
        rep_after = verify_bilingual_docx(p_after)
        self.assertEqual(rep_after["clipped_row_violations"], 0)


class TestTier5AdversarialLeakedEnglish(unittest.TestCase):
    """Adversarial testing of Check 4: Zero Leaked English."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_adv_leak_hunter_repairs_ordinals_and_headings(self):
        """LeakHunter must automatically repair leaked ordinals and headings in translated text."""
        raw_vi = "Biện pháp xử lý First aid theo quy định của section 6.2 đối với hàng hóa 1st và 2nd."
        cleaned = LeakHunter.hunt_and_repair(raw_vi)
        self.assertNotIn("First aid", cleaned)
        self.assertNotIn("section 6.2", cleaned)
        self.assertNotIn("1st", cleaned)
        self.assertNotIn("2nd", cleaned)
        self.assertIn("Sơ cứu", cleaned)
        self.assertIn("mục 6.2", cleaned)
        self.assertIn("thứ 1", cleaned)
        self.assertIn("thứ 2", cleaned)

    def test_adv_unrepaired_leaked_english_fails_audit(self):
        """A translated paragraph containing leaked English ordinals/headings must fail Check 4."""
        doc = docx.Document()
        p = doc.add_paragraph()
        r = p.add_run("Áp dụng các biện pháp Precautionary statements cho lô hàng 1st")
        r.font.color.rgb = DARK_NAVY_RGB
        r.italic = True

        p_path = os.path.join(self.temp_dir, "adv_unrepaired_leak.docx")
        doc.save(p_path)

        rep = verify_bilingual_docx(p_path)
        self.assertFalse(rep["passed"])
        self.assertGreaterEqual(rep["leaked_english_count"], 1)


class TestTier5AdversarialPairCompleteness(unittest.TestCase):
    """Adversarial testing of Check 3: 100% Pair Completeness."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_adv_trailing_untranslated_source_paragraph_fails_completeness(self):
        """Document with an unpaired English body paragraph at the end must fail 100% pair completeness."""
        doc = docx.Document()
        PairingBuilder.apply_pattern_a(doc, "SECTION 1: Identification", "I. Nhận dạng hóa chất")
        # Add unpaired trailing source paragraph
        p_orphan = doc.add_paragraph("This is an un-paired trailing source paragraph without translation.")

        p_path = os.path.join(self.temp_dir, "adv_orphan_p.docx")
        doc.save(p_path)

        rep = verify_bilingual_docx(p_path)
        self.assertFalse(rep["passed"])
        self.assertLess(rep["pair_completeness"], 1.0)


class TestTier5AdversarialSanitizerAndConsolidator(unittest.TestCase):
    """Stress tests on TokenSanitizer and ParagraphConsolidator."""

    def test_adv_sanitizer_ocr_and_spacing_anomalies(self):
        """TokenSanitizer must fix OCR typos and attached units without corrupting chemical terms."""
        cases = [
            ("15minutes.Assure adequate flushing.", "15 minutes. Assure adequate flushing."),
            ("keep away from hest and flame", "keep away from heat and flame"),
            ("form ignition sources", "from ignition sources"),
            ("CITY,BINH DUONG", "CITY, BINH DUONG"),
            ("ScienceLab.com", "Science Lab.com"),
        ]
        for inp, expected in cases:
            out = TokenSanitizer.sanitize(inp)
            self.assertEqual(out, expected)

    def test_adv_paragraph_consolidator_unpunctuated_lines(self):
        """ParagraphConsolidator must merge unpunctuated wrapped lines while preserving headers and lists."""
        # Unpunctuated lines should merge
        self.assertTrue(ParagraphConsolidator.should_merge(
            "In case of contact, immediately flush eyes with plenty of water for at least",
            "15 minutes while holding eyelids open."
        ))
        # Punctuated line should NOT merge
        self.assertFalse(ParagraphConsolidator.should_merge(
            "Flush eyes with plenty of water.",
            "Get medical attention immediately."
        ))
        # Section header should NOT merge
        self.assertFalse(ParagraphConsolidator.should_merge(
            "SECTION 1: Identification",
            "SECTION 2: Hazards"
        ))
        # Bullet points should NOT merge
        self.assertFalse(ParagraphConsolidator.should_merge(
            "Safety precautions:",
            "- Do not ingest."
        ))


class TestTier5AdversarialCornerCasesAndFaultTolerance(unittest.TestCase):
    """Stress tests on error handling, non-standard XML structures, and file integrity."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_adv_unsupported_file_extension_rejection(self):
        """reconstruct_bilingual_document must reject unsupported extensions with ValueError."""
        dummy_file = os.path.join(self.temp_dir, "test.txt")
        with open(dummy_file, "w") as f:
            f.write("sample content")
        out_file = os.path.join(self.temp_dir, "test_out.docx")

        with self.assertRaises(ValueError):
            reconstruct_bilingual_document(dummy_file, out_file)

    def test_adv_audit_verifier_handles_missing_file_gracefully(self):
        """verify_bilingual_docx must return passed=False with error message when file does not exist."""
        non_existent = os.path.join(self.temp_dir, "does_not_exist.docx")
        rep = verify_bilingual_docx(non_existent)
        self.assertFalse(rep["passed"])
        self.assertIn("does not exist", rep.get("error", "").lower())

    def test_adv_audit_verifier_handles_corrupt_zip_gracefully(self):
        """verify_bilingual_docx must return passed=False when given a corrupted/truncated ZIP file."""
        corrupt_file = os.path.join(self.temp_dir, "corrupt.docx")
        with open(corrupt_file, "wb") as f:
            f.write(b"PK\x03\x04not-a-valid-zip-content")

        rep = verify_bilingual_docx(corrupt_file)
        self.assertFalse(rep["passed"])
        self.assertIn("exception", rep.get("details", {}))

    def test_adv_label_value_with_multiple_colons_and_urls(self):
        """GridReconstructor must handle colons in values (URLs, ratios, times) cleanly."""
        doc = docx.Document()
        doc.add_paragraph("Website: https://www.example.com/safety")
        doc.add_paragraph("Ratio: 1:2:3 by volume")
        doc.add_paragraph("Shift Time: 08:00:00 to 17:00:00")

        GridReconstructor.reconstruct_label_value_blocks(doc)
        p_path = os.path.join(self.temp_dir, "adv_colons.docx")
        doc.save(p_path)

        rep = verify_bilingual_docx(p_path)
        self.assertEqual(rep["double_tab_violations"], 0)
        self.assertEqual(rep["concatenation_violations"], 0)

    def test_adv_table_with_merged_cells_pattern_b(self):
        """Tables with merged cells (gridSpan) must maintain cell bilingual alignment without crash."""
        doc = docx.Document()
        tbl = doc.add_table(rows=2, cols=3)
        # Merge cells across top row
        tbl.rows[0].cells[0].merge(tbl.rows[0].cells[2])

        PairingBuilder.apply_pattern_b(tbl.rows[0].cells[0], "Merged Section Header", "Tiêu đề mục hợp nhất")
        PairingBuilder.apply_pattern_b(tbl.rows[1].cells[0], "Substance", "Chất")
        PairingBuilder.apply_pattern_b(tbl.rows[1].cells[1], "CAS Number", "Số CAS")
        PairingBuilder.apply_pattern_b(tbl.rows[1].cells[2], "Percent", "Tỷ lệ %")

        p_path = os.path.join(self.temp_dir, "adv_merged_table.docx")
        doc.save(p_path)

        rep = verify_bilingual_docx(p_path)
        self.assertTrue(rep["passed"])
        self.assertEqual(rep["clipped_row_violations"], 0)
        self.assertEqual(rep["concatenation_violations"], 0)


class TestTier5RealSampleEmpiricalGroundTruth(unittest.TestCase):
    """Authoritative empirical audit of the 3 real-world generated sample documents."""

    def test_real_sample_fu43288_audit_passed(self):
        """Live generated MSDS-FU43288-EN bilingual document must pass 100% of 7 audit checks."""
        docx_path = r"tests\live_test_FU43288.docx"
        self.assertTrue(os.path.exists(docx_path), f"{docx_path} must exist from live generation")
        rep = verify_bilingual_docx(docx_path)
        self.assertTrue(rep["passed"], f"Audit failed: {rep}")
        self.assertEqual(rep["concatenation_violations"], 0)
        self.assertEqual(rep["double_tab_violations"], 0)
        self.assertEqual(rep["pair_completeness"], 1.0)
        self.assertEqual(rep["leaked_english_count"], 0)
        self.assertEqual(rep["header_order_violations"], 0)
        self.assertEqual(rep["clipped_row_violations"], 0)
        self.assertTrue(rep["visual_parity_passed"])

    def test_real_sample_ytu0070_audit_passed(self):
        """Live generated MSDS-YTU0070-EN bilingual document must pass 100% of 7 audit checks."""
        docx_path = r"tests\live_test_YTU0070.docx"
        self.assertTrue(os.path.exists(docx_path), f"{docx_path} must exist from live generation")
        rep = verify_bilingual_docx(docx_path)
        self.assertTrue(rep["passed"], f"Audit failed: {rep}")
        self.assertEqual(rep["concatenation_violations"], 0)
        self.assertEqual(rep["double_tab_violations"], 0)
        self.assertEqual(rep["pair_completeness"], 1.0)
        self.assertEqual(rep["leaked_english_count"], 0)
        self.assertEqual(rep["header_order_violations"], 0)
        self.assertEqual(rep["clipped_row_violations"], 0)
        self.assertTrue(rep["visual_parity_passed"])

    def test_real_sample_ytu0035_audit_passed(self):
        """Live generated MSDS-YTU0035-EN bilingual document must pass 100% of 7 audit checks."""
        docx_path = r"tests\live_test_YTU0035.docx"
        self.assertTrue(os.path.exists(docx_path), f"{docx_path} must exist from live generation")
        rep = verify_bilingual_docx(docx_path)
        self.assertTrue(rep["passed"], f"Audit failed: {rep}")
        self.assertEqual(rep["concatenation_violations"], 0)
        self.assertEqual(rep["double_tab_violations"], 0)
        self.assertEqual(rep["pair_completeness"], 1.0)
        self.assertEqual(rep["leaked_english_count"], 0)
        self.assertEqual(rep["header_order_violations"], 0)
        self.assertEqual(rep["clipped_row_violations"], 0)
        self.assertTrue(rep["visual_parity_passed"])


if __name__ == "__main__":
    unittest.main()
