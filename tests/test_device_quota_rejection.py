"""
tests/test_device_quota_rejection.py
───────────────────────────────────
Unit and integration tests for device limit enforcement across LUV-DOC-TRAN and TitanPDF.
Specifically validates: "quá số lượng thì sẽ bị từ chối"
"""

import unittest
import json
import os
import sys
from unittest.mock import patch, MagicMock
import urllib.error

# Add parent directory
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from core.license_client import LUVLicenseClient


class TestDeviceQuotaRejection(unittest.TestCase):
    """Verify that when active devices exceed max devices quota, activation is strictly rejected."""

    def test_01_device_limit_rejection_in_license_client(self):
        """When Cloudflare Worker returns error: device_limit, activate_key raises ValueError with explicit rejection."""
        mock_http_error = urllib.error.HTTPError(
            url="https://luv-license-api.luv-sds.workers.dev/v1/activate",
            code=403,
            msg="Forbidden",
            hdrs={},
            fp=MagicMock(read=lambda: json.dumps({
                "error": "device_limit",
                "message": "Key has reached maximum device limit."
            }).encode("utf-8"))
        )

        with patch("urllib.request.urlopen", side_effect=mock_http_error):
            with self.assertRaises(ValueError) as ctx:
                LUVLicenseClient.activate_key("LUV-12345-67890-ABCDE-FGHIJ", "Test User", "dev-999")
            
            err_msg = str(ctx.exception)
            self.assertIn("vượt quá số lượng thiết bị cho phép", err_msg.lower())
            self.assertIn("từ chối", err_msg.lower())

    def test_02_titanpdf_verify_key_online_device_limit_rejection(self):
        """TitanPDF's verify_key_online rejects keys when active_devices >= max_devices."""
        titan_file = os.path.join(r"c:\Users\ntc27\OneDrive\文件\TitanPDF", "core", "license_manager.py")
        if not os.path.exists(titan_file):
            self.skipTest("TitanPDF license_manager.py not found")

        import importlib.util
        spec = importlib.util.spec_from_file_location("titan_license_manager", titan_file)
        titan_mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(titan_mod)
        LUVLicenseManager = titan_mod.LUVLicenseManager

        # Mock database with a key that reached limit (1/1 máy)
        mock_db = [{
            "id": "mock-test-id",
            "license_key": "LUV-TEST-FULL-KEY",
            "registered_to": "Khách Đầy Máy",
            "status": "active",
            "max_devices": 1,
            "active_devices": 1,
            "devices": "1/1 máy"
        }]

        with patch.object(LUVLicenseManager, "_read_db", return_value=mock_db):
            result = LUVLicenseManager.verify_key_online("LUV-TEST-FULL-KEY")
            self.assertFalse(result.get("valid"))
            self.assertEqual(result.get("status"), "device_limit")
            self.assertIn("từ chối", result.get("reason", "").lower())
            self.assertIn("1/1", result.get("reason", ""))

    def test_03_titanpdf_handover_message_contains_device_rejection_notice(self):
        """Handover message explicitly reminds client about device limit and rejection."""
        titan_file = os.path.join(r"c:\Users\ntc27\OneDrive\文件\TitanPDF", "core", "license_manager.py")
        if not os.path.exists(titan_file):
            self.skipTest("TitanPDF license_manager.py not found")

        import importlib.util
        spec = importlib.util.spec_from_file_location("titan_license_manager_h", titan_file)
        titan_mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(titan_mod)
        LUVLicenseManager = titan_mod.LUVLicenseManager

        msg = LUVLicenseManager.format_handover_message(
            key="LUV-XXXXX-XXXXX",
            registered_to="Công Ty ABC",
            tier="Enterprise",
            expires="Vĩnh viễn",
            max_devices=2
        )
        self.assertIn("2 máy", msg)
        self.assertIn("Quá số lượng sẽ bị từ chối", msg)


if __name__ == "__main__":
    unittest.main()
