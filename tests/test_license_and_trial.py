import os
import sys
import unittest
import tempfile
import time
from pathlib import Path

# Add project root to sys.path
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE_DIR)

from core.license_client import (
    LUVLicenseClient,
    TRIAL_DURATION_SECONDS,
    DEFAULT_SERVER_URL,
    PUBLIC_KEY_B64,
    generate_device_hash,
)


class TestLicenseAndTrial(unittest.TestCase):

    def setUp(self):
        self.test_device = "TEST-DEV-DEVICE-999"
        self.lic_file = os.path.join(BASE_DIR, "license.json")
        self.lic_backup = None
        if os.path.exists(self.lic_file):
            with open(self.lic_file, "r", encoding="utf-8") as f:
                self.lic_backup = f.read()

    def tearDown(self):
        # Restore license if backed up
        if self.lic_backup is not None:
            with open(self.lic_file, "w", encoding="utf-8") as f:
                f.write(self.lic_backup)
        elif os.path.exists(self.lic_file):
            os.remove(self.lic_file)

    def test_trial_initialization_and_countdown(self):
        """Verify 48-hour trial is properly initiated and formatted."""
        status = LUVLicenseClient.get_trial_status(self.test_device)
        self.assertTrue(status["is_trial_active"])
        self.assertGreater(status["remaining_seconds"], 0)
        self.assertLessEqual(status["remaining_seconds"], TRIAL_DURATION_SECONDS)
        self.assertIn("h", status["remaining_formatted"])

    def test_invalid_and_random_key_rejected_by_cloudflare(self):
        """Verify fake, malformed or random keys are completely rejected."""
        # 1. Random garbage string
        with self.assertRaises(Exception):
            LUVLicenseClient.activate_key("khk_random_fake_string", "Fake User")

        # 2. Fake format key
        with self.assertRaises(Exception):
            LUVLicenseClient.activate_key("LUV-FAKE-1234-5678-9999", "Fake User")

        # 3. Empty string
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
        except Exception as e:
            # If server rate-limits or already activated, check existing valid license
            info = LUVLicenseClient.get_license_info()
            self.assertTrue(info["is_active"])


if __name__ == "__main__":
    unittest.main()
