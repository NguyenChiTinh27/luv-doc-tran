"""
tests/test_keyword_cleanse.py
──────────────────────────────
Automated test suite verifying 100% eradication of chemical, MSDS, CAS, ISO, and GHS keywords
from all frontend assets (in static/ and templates/) in accordance with ORIGINAL_REQUEST.md §R1
and TEST_INFRA.md § Feature Inventory #5.

Target keywords to assert 0 occurrences:
- "hóa chất", "hoá chất", "hoa chat", "An Toàn Hóa Chất"
- "Mã CAS", "ma cas", "CAS Number", standalone "\\bCAS\\b" (excluding words like "uppercase")
- "ISO & GHS", "Tiêu chuẩn ISO & GHS", "ISO &amp; GHS", standalone "\\bGHS\\b"
- "MSDS", "MSDS/SDS", "\\bmsds\\b"
"""

from __future__ import annotations

import os
import re
import unittest
from typing import Dict, List, Tuple

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STATIC_DIR = os.path.join(PROJECT_ROOT, "static")
TEMPLATES_DIR = os.path.join(PROJECT_ROOT, "templates")

FRONTEND_EXTENSIONS = {".html", ".htm", ".js", ".css", ".json", ".svg", ".txt"}

# Compiled Regex Patterns
RE_HOA_CHAT = re.compile(r"(hóa chất|hoá chất|hoa chat)", re.IGNORECASE)
RE_CAS = re.compile(r"(mã\s+cas|ma\s+cas|\bcas\s*:\s*\d|\bcas\s+number\b|\bCAS\b)")
RE_ISO_GHS = re.compile(r"(iso\s*&\s*ghs|iso\s*&amp;\s*ghs|tiêu chuẩn iso|\bghs\b)", re.IGNORECASE)
RE_MSDS = re.compile(r"(\bmsds\b)", re.IGNORECASE)

# Combined master pattern from TEST_INFRA.md
RE_MASTER_AUDIT = re.compile(
    r"(hóa chất|hoá chất|hoa chat|mã\s+cas|ma\s+cas|\bcas\s*:\s*\d|\bcas\s+number\b|\bCAS\b|iso\s*&\s*ghs|iso\s*&amp;\s*ghs|tiêu chuẩn iso|\bghs\b|\bmsds\b)",
    re.IGNORECASE,
)


def get_all_frontend_files() -> List[str]:
    """Collect all files from static/ and templates/ with frontend extensions."""
    files_to_scan: List[str] = []
    target_dirs = [STATIC_DIR, TEMPLATES_DIR]

    for tdir in target_dirs:
        if not os.path.exists(tdir):
            continue
        for root, _, files in os.walk(tdir):
            for f in files:
                ext = os.path.splitext(f)[1].lower()
                if ext in FRONTEND_EXTENSIONS:
                    files_to_scan.append(os.path.join(root, f))

    return sorted(files_to_scan)


def scan_files_for_pattern(pattern: re.Pattern) -> List[Tuple[str, int, str, str]]:
    """Scan all frontend files for a regex pattern. Returns list of (rel_path, line_no, matched_text, line_content)."""
    violations: List[Tuple[str, int, str, str]] = []
    files = get_all_frontend_files()

    for fpath in files:
        rel_path = os.path.relpath(fpath, PROJECT_ROOT)
        try:
            with open(fpath, "r", encoding="utf-8", errors="ignore") as fp:
                for line_no, line in enumerate(fp, start=1):
                    match = pattern.search(line)
                    if match:
                        violations.append((rel_path, line_no, match.group(0), line.strip()))
        except Exception as exc:
            violations.append((rel_path, 0, f"READ_ERROR: {exc}", ""))

    return violations


class TestKeywordCleanse(unittest.TestCase):
    """Automated keyword sweep test asserting 0 occurrences of chemical/MSDS/CAS/ISO/GHS terms in frontend assets."""

    def test_00_frontend_files_exist(self):
        """Verify frontend assets exist and are discoverable for scanning."""
        files = get_all_frontend_files()
        self.assertGreater(len(files), 0, "No frontend files found to scan in static/ or templates/")
        scanned_names = [os.path.basename(f) for f in files]
        self.assertIn("index.html", scanned_names, "static/index.html must exist")

    def test_01_no_chemical_keywords(self):
        """Assert zero occurrences of 'hóa chất' / 'hoá chất' / 'hoa chat' across all frontend files."""
        violations = scan_files_for_pattern(RE_HOA_CHAT)
        if violations:
            msg_lines = [f"Found {len(violations)} chemical keyword violation(s):"]
            for path, line_no, matched, line_text in violations:
                msg_lines.append(f"  {path}:{line_no} [{matched}] -> {line_text[:100]}")
            self.fail("\n".join(msg_lines))

    def test_02_no_cas_number_keywords(self):
        """Assert zero occurrences of 'Mã CAS' / 'CAS Number' / standalone CAS in frontend files."""
        violations = scan_files_for_pattern(RE_CAS)
        if violations:
            msg_lines = [f"Found {len(violations)} CAS keyword violation(s):"]
            for path, line_no, matched, line_text in violations:
                msg_lines.append(f"  {path}:{line_no} [{matched}] -> {line_text[:100]}")
            self.fail("\n".join(msg_lines))

    def test_03_no_iso_and_ghs_keywords(self):
        """Assert zero occurrences of 'ISO & GHS' / 'Tiêu chuẩn ISO' / 'GHS' in frontend files."""
        violations = scan_files_for_pattern(RE_ISO_GHS)
        if violations:
            msg_lines = [f"Found {len(violations)} ISO/GHS keyword violation(s):"]
            for path, line_no, matched, line_text in violations:
                msg_lines.append(f"  {path}:{line_no} [{matched}] -> {line_text[:100]}")
            self.fail("\n".join(msg_lines))

    def test_04_no_msds_keywords(self):
        """Assert zero occurrences of 'MSDS' / 'msds' in frontend files."""
        violations = scan_files_for_pattern(RE_MSDS)
        if violations:
            msg_lines = [f"Found {len(violations)} MSDS keyword violation(s):"]
            for path, line_no, matched, line_text in violations:
                msg_lines.append(f"  {path}:{line_no} [{matched}] -> {line_text[:100]}")
            self.fail("\n".join(msg_lines))

    def test_05_master_audit_zero_tolerance(self):
        """Audit all frontend files against combined master pattern (TEST_INFRA.md specification)."""
        violations = scan_files_for_pattern(RE_MASTER_AUDIT)
        if violations:
            msg_lines = [
                f"Master keyword cleanse audit failed with {len(violations)} violation(s):"
            ]
            for path, line_no, matched, line_text in violations:
                msg_lines.append(f"  {path}:{line_no} -> match: '{matched}' in line: '{line_text[:120]}'")
            self.fail("\n".join(msg_lines))

    def test_06_brand_identity_repositioning(self):
        """Verify frontend contains universal translation branding."""
        index_path = os.path.join(STATIC_DIR, "index.html")
        if not os.path.exists(index_path):
            self.skipTest("index.html does not exist yet")
        with open(index_path, "r", encoding="utf-8") as f:
            content = f.read()

        # Must contain brand name LUV-DOC-TRAN
        self.assertIn("LUV-DOC-TRAN", content, "Branding 'LUV-DOC-TRAN' must be present")

    def test_07_no_doanh_nghiep_keywords(self):
        """Assert zero occurrences of 'doanh nghiệp' / 'Doanh nghiệp' across frontend HTML/JS assets."""
        pattern = re.compile(r"doanh\s*nghiệp", re.IGNORECASE)
        violations = scan_files_for_pattern(pattern)
        if violations:
            msg_lines = [f"Found {len(violations)} 'doanh nghiệp' violation(s):"]
            for path, line_no, matched, line_text in violations:
                msg_lines.append(f"  {path}:{line_no} [{matched}] -> {line_text[:100]}")
            self.fail("\n".join(msg_lines))

    def test_08_no_deepseek_keywords(self):
        """Assert zero occurrences of 'deepseek' / 'DeepSeek' in user-facing index.html and app.js."""
        pattern = re.compile(r"deepseek", re.IGNORECASE)
        violations = scan_files_for_pattern(pattern)
        if violations:
            msg_lines = [f"Found {len(violations)} 'deepseek' violation(s):"]
            for path, line_no, matched, line_text in violations:
                msg_lines.append(f"  {path}:{line_no} [{matched}] -> {line_text[:100]}")
            self.fail("\n".join(msg_lines))

    def test_09_no_ban_dung_or_free_trial(self):
        """Assert zero occurrences of 'bản dùng' or 'dùng thử' in static/index.html."""
        pattern = re.compile(r"(bản\s*dùng|dùng\s*thử)", re.IGNORECASE)
        index_path = os.path.join(STATIC_DIR, "index.html")
        violations = []
        with open(index_path, "r", encoding="utf-8") as fp:
            for line_no, line in enumerate(fp, 1):
                if pattern.search(line):
                    violations.append((line_no, line.strip()))
        self.assertEqual(len(violations), 0, f"Found trial/ban dung text in index.html: {violations}")

    def test_10_uses_key_kich_hoat(self):
        """Assert 'Key Kích Hoạt' is prominently used on UI."""
        index_path = os.path.join(STATIC_DIR, "index.html")
        with open(index_path, "r", encoding="utf-8") as fp:
            content = fp.read()
        self.assertIn("Key Kích Hoạt", content)

    def test_11_only_key_input_no_name_field(self):
        """Assert modal only requires Key input without requiring user name."""
        index_path = os.path.join(STATIC_DIR, "index.html")
        with open(index_path, "r", encoding="utf-8") as fp:
            content = fp.read()
        self.assertIn('id="act-key"', content)
        self.assertNotIn('id="act-reg-name"', content)

    def test_12_no_ed25519_cloudflare_subheading(self):
        """Assert zero occurrences of 'Xác thực mật mã Ed25519' in static/index.html."""
        index_path = os.path.join(STATIC_DIR, "index.html")
        with open(index_path, "r", encoding="utf-8") as fp:
            content = fp.read()
        self.assertNotIn("Xác thực mật mã Ed25519", content)
        self.assertNotIn("modal-subheading", content)


if __name__ == "__main__":
    unittest.main()
