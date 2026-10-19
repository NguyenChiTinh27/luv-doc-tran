import os
import sys
import json
import time
import unittest
from pathlib import Path
from unittest.mock import patch, MagicMock
import urllib.error

# Add project root to sys.path
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE_DIR)

from core.license_client import (
    LUVLicenseClient,
    TRIAL_DURATION_SECONDS,
    DEFAULT_SERVER_URL,
    PUBLIC_KEY_B64,
    LICENSE_FILE,
    LEGACY_LICENSE_FILE,
    generate_device_hash,
)

# Authentic sample token for offline and signature tests
SAMPLE_DOC_TOKEN = (
    "eyJpc3MiOiJsdXYtbGljZW5zZS1zZXJ2ZXIiLCJhdWQiOiJsdXYtc2RzLWludGVsbGlnZW5jZSIsInN1YiI6I"
    "jkwYTFkMmM0LTE2MGYtNGM4Ny1iN2YzLTEzN2E0MmViMTIzMCIsImxpY2Vuc2VfaWQiOiI5MGExZDJjNC0xNj"
    "BmLTRjODctYjdmMy0xMzdhNDJlYjEyMzAiLCJwcm9kdWN0X2NvZGUiOiJET0MiLCJjdXN0b21lciI6IkRvYW5"
    "oIE5naGllcCBMVVYgRE9DLVRSQU4iLCJwbGFuIjoiZW50ZXJwcmlzZSIsImZlYXR1cmVzIjpbImNvcmUiLCJp"
    "bXBvcnQiLCJleHBvcnQiLCJlZGl0IiwiZG9jX3RyYW4iXSwiZGV2aWNlX2lkIjoiMTExMTIyMjIzMzMzNDQ0ND"
    "U1NTU2NjY2Nzc3Nzg4ODg5OTk5MDAwMGFhYWFiYmJiY2NjY2RkZGRlZWVlZmZmZiIsImlhdCI6MTc5MjMzNTA2"
    "MiwicmVmcmVzaF9hZnRlciI6MTc5MjQyMTQ2Miwib2ZmbGluZV91bnRpbCI6MTc5MzU0NDY2Miwib2ZmbGluZV"
    "91bnRpbF9pc28iOiIyMDI2LTExLTAxVDE0OjUxOjAyLjkzN1oiLCJleHAiOjIxMTQzODA3OTksImV4cGlyZXNf"
    "YXQiOiIyMDM2LTEyLTMxVDIzOjU5OjU5LjAwMFoiLCJqdGkiOiJiZjVlMzQ3MC05OWJlLTQ1NTgtODNjOC0zZG"
    "UxMGExMGFjMTQiLCJsZWFzZV9zZXEiOjF9."
    "6aEJEn351Quc6Ct9Ptck4aa9I-EQpRvdKqw8gPf_o-wb4z_zHiwS9Q4WOj3ohkwoUbo5TxAokJkwmxebcfRmBw"
)


class TestLicenseAndTrial(unittest.TestCase):

    def setUp(self):
        self.test_device = "TEST-DEV-DEVICE-999"
        self.lic_json = os.path.join(BASE_DIR, "license.json")
        self.lic_sec = LICENSE_FILE
        self.lic_backup_json = None
        self.lic_backup_sec = None

        if os.path.exists(self.lic_json):
            with open(self.lic_json, "rb") as f:
                self.lic_backup_json = f.read()
        if os.path.exists(self.lic_sec):
            with open(self.lic_sec, "rb") as f:
                self.lic_backup_sec = f.read()

        self._clean_files()

    def tearDown(self):
        self._clean_files()
        if self.lic_backup_json is not None:
            with open(self.lic_json, "wb") as f:
                f.write(self.lic_backup_json)
        if self.lic_backup_sec is not None:
            with open(self.lic_sec, "wb") as f:
                f.write(self.lic_backup_sec)

    def _clean_files(self):
        for path in [self.lic_json, self.lic_sec, self.lic_sec + ".tmp"]:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except Exception:
                    pass

    def test_trial_initialization_and_countdown(self):
        """Verify trial is eliminated and mandatory key is enforced from the start."""
        status = LUVLicenseClient.get_trial_status(self.test_device)
        self.assertFalse(status["is_trial_active"])
        self.assertTrue(status.get("trial_disabled"))
        self.assertEqual(status["remaining_seconds"], 0)

        # check_access must report unactivated and block access
        access = LUVLicenseClient.check_access(self.test_device)
        self.assertFalse(access["has_access"])
        self.assertEqual(access["access_type"], "unactivated")
        self.assertIn("Key", access["message"])

    def test_invalid_and_random_key_rejected_by_cloudflare(self):
        """Verify fake, malformed or random keys are completely rejected."""
        with self.assertRaises(Exception):
            LUVLicenseClient.activate_key("khk_random_fake_string", "Fake User")

        with self.assertRaises(Exception):
            LUVLicenseClient.activate_key("LUV-FAKE-1234-5678-9999", "Fake User")

        with self.assertRaises(Exception):
            LUVLicenseClient.activate_key("", "Fake User")

    def test_genuine_key_activation_or_validation(self):
        """Verify genuine key format or active license status."""
        real_key = "LUV-YMZHK-QUKY8-D4ATT-TSJVY"
        try:
            res = LUVLicenseClient.activate_key(real_key, "Doanh Nghiep LUV DOC-TRAN")
            self.assertTrue(res["is_active"])
            self.assertEqual(res["license_key"], real_key)
            self.assertIn("customer", res)
        except Exception:
            info = LUVLicenseClient.get_license_info()
            self.assertTrue(info["is_active"])

    def test_purge_license_safely_removes_files(self):
        """Verify purge_license removes license.sec and license.json safely."""
        with open(self.lic_sec, "wb") as f:
            f.write(b"dummy_sec_data")
        with open(self.lic_json, "w", encoding="utf-8") as f:
            f.write('{"dummy": true}')

        self.assertTrue(os.path.exists(self.lic_sec))
        self.assertTrue(os.path.exists(self.lic_json))

        LUVLicenseClient.purge_license()

        self.assertFalse(os.path.exists(self.lic_sec))
        self.assertFalse(os.path.exists(self.lic_json))

    def test_validate_license_unactivated_returns_clean_status(self):
        """Verify validate_license when unactivated returns unactivated status."""
        self._clean_files()
        res = LUVLicenseClient.validate_license(force_online=False)
        self.assertFalse(res["is_active"])
        self.assertFalse(res.get("revoked", False))
        self.assertEqual(res.get("status"), "unactivated")

    def test_validate_license_offline_tolerance_valid_window(self):
        """Verify validate_license permits offline operation during network outage if within offline_until."""
        record = {
            "token": SAMPLE_DOC_TOKEN,
            "license_key": "LUV-TEST-OFFLINE-KEY",
            "registered_to": "Offline Test User",
            "offline_until": 2000000000,  # Far in the future
            "is_active": True,
            "status": "active",
        }
        LUVLicenseClient.save_secure_license(record)

        with patch("urllib.request.urlopen", side_effect=urllib.error.URLError("Network unreachable")):
            status = LUVLicenseClient.validate_license(force_online=True)
            self.assertTrue(status.get("is_active"))
            self.assertTrue(status.get("offline_mode"))
            self.assertEqual(status.get("status"), "offline_active")

            access = LUVLicenseClient.check_access(validated_license=status)
            self.assertTrue(access.get("has_access"))
            self.assertEqual(access.get("access_type"), "licensed")

    def test_validate_license_offline_tolerance_expired_blocks(self):
        """Verify validate_license blocks offline operation if offline_until is in the past."""
        record = {
            "token": SAMPLE_DOC_TOKEN,
            "license_key": "LUV-TEST-EXPIRED-OFFLINE-KEY",
            "registered_to": "Offline Expired User",
            "is_active": True,
            "status": "active",
        }
        LUVLicenseClient.save_secure_license(record)

        # Simulate time in the future beyond offline_until (claims offline_until is 1793544662)
        future_time = 1800000000
        with patch("urllib.request.urlopen", side_effect=urllib.error.URLError("Network unreachable")):
            with patch("core.license_client._utc_now", return_value=future_time):
                status = LUVLicenseClient.validate_license(force_online=True)
                self.assertFalse(status.get("is_active"))
                self.assertTrue(status.get("offline_expired"))
                self.assertEqual(status.get("status"), "offline_expired")

                access = LUVLicenseClient.check_access(validated_license=status)
                self.assertFalse(access.get("has_access"))
                self.assertTrue(access.get("offline_expired"))

    def test_validate_license_revocation_auto_purges_and_blocks_permanently(self):
        """Verify validate_license upon HTTP 403 license_denied purges disk storage and permanently revokes."""
        record = {
            "token": SAMPLE_DOC_TOKEN,
            "license_key": "LUV-REVOKED-TEST-KEY",
            "registered_to": "Revoked User",
            "device_id": "REVOKED-DEV-ID-12345",
            "is_active": True,
            "status": "active",
        }
        LUVLicenseClient.save_secure_license(record)
        self.assertTrue(os.path.exists(self.lic_sec))

        error_response = urllib.error.HTTPError(
            url=f"{DEFAULT_SERVER_URL}/v1/validate",
            code=403,
            msg="Forbidden",
            hdrs={},
            fp=None
        )
        error_response.read = MagicMock(return_value=b'{"error": "license_denied", "message": "Gi\xe1\xba\xa5y ph\xc3\xa9p \xc4\x91\xc3\xa3 b\xe1\xbb\x8b kh\xc3\xb3a ho\xe1\xba\xb7c thu h\xe1\xbb\x93i."}')

        with patch("urllib.request.urlopen", side_effect=error_response):
            val_res = LUVLicenseClient.validate_license(device_id="REVOKED-DEV-ID-12345", force_online=True)
            self.assertFalse(val_res.get("is_active"))
            self.assertTrue(val_res.get("revoked"))
            self.assertEqual(val_res.get("status"), "revoked")
            self.assertIn(val_res.get("tier"), ["Key Đã Bị Thu Hồi", "Giấy phép đã bị thu hồi"])

            # Check that license was auto-purged from disk
            self.assertFalse(os.path.exists(self.lic_sec))

            # Check that access is permanently severed, NO fallback to trial
            access = LUVLicenseClient.check_access(device_id="REVOKED-DEV-ID-12345", validated_license=val_res)
            self.assertFalse(access.get("has_access"))
            self.assertTrue(access.get("revoked"))
            self.assertEqual(access.get("access_type"), "revoked")


if __name__ == "__main__":
    unittest.main()
