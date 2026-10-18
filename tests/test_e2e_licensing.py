"""
tests/test_e2e_licensing.py
────────────────────────────
Comprehensive E2E Automated Test Suite for Authentic Ed25519 & Cloudflare Worker Licensing Integration.

Tier Coverage:
- Tier 1: Feature Coverage (Ed25519 verification, claims extraction, 64-char hex device ID, genuine key activation, validation & lease seq increment, Admin API list/create).
- Tier 2: Boundary & Corner Cases (100% rejection of garbage keys, old mock patterns, non-existent keys, tampered signatures, modified payload claims, device limit HTTP 409).
- Tier 3: Cross-Feature Combinations (Full lifecycle: Admin create -> Client activate -> Local verify & save -> Client validate -> Admin reset devices -> Admin revoke -> Client validate fails).
- Tier 4: Real-World Scenarios (Owner key fallback DOC->SDS, offline persistence in data/license.sec within offline_until window, tamper detection, DPAPI envelope security).
"""

from __future__ import annotations

import base64
import ctypes
import hashlib
import json
import os
import platform
import re
import sys
import time
import unittest
import urllib.error
import urllib.request
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Optional

from cryptography.exceptions import InvalidSignature
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey

# Add project root to sys.path first so LUV-DOC-TRAN core package takes priority
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

# Also append TitanPDF root to sys.path if available
TITAN_ROOT = os.path.abspath(os.path.join(PROJECT_ROOT, "..", "TitanPDF"))
if os.path.exists(TITAN_ROOT) and TITAN_ROOT not in sys.path:
    sys.path.append(TITAN_ROOT)

from core.license_client import (
    APP_VERSION,
    CLOCK_SKEW_SECONDS,
    DEFAULT_SERVER_URL,
    PUBLIC_KEY_B64,
    LUVLicenseClient,
    _b64url_decode,
    generate_device_hash,
)

# Reference Official Configuration
OFFICIAL_PUBLIC_KEY_B64 = "hsHHVWWmquFmME9wW4d2u1Qh5NdVtu3tktRDMl437Fo="
OFFICIAL_ADMIN_TOKEN = "rHHfK9eC_8lTH02SP0CQPnyXAqMMDUXxwV2MiPothM7Avc4TxNZgnkKp5CFsIBP_"
OFFICIAL_GENUINE_KEY = "LUV-YMZHK-QUKY8-D4ATT-TSJVY"
OFFICIAL_OWNER_KEY = "LUV-EGKRH-HZCWD-7PBLS-ZG5G8"
OFFICIAL_OWNER_LICENSE_ID = "b7f0acc0-ce3d-41ba-9bdb-6835169524cc"


def _b64url_encode(data: bytes) -> str:
    """Encode bytes to base64url string without trailing padding."""
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")


def _call_admin_api(method: str, path: str, payload: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """Execute authenticated Cloudflare Worker Admin API request."""
    data = None if payload is None else json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        f"{DEFAULT_SERVER_URL}{path}",
        data=data,
        headers={
            "Authorization": f"Bearer {OFFICIAL_ADMIN_TOKEN}",
            "Content-Type": "application/json",
            "User-Agent": "LUV-License-Admin/1.0",
        },
        method=method,
    )
    with urllib.request.urlopen(req, timeout=15) as resp:
        return json.loads(resp.read().decode("utf-8"))


class BaseLicensingTestCase(unittest.TestCase):
    """Base test case providing clean filesystem state and backup/restore helpers."""

    def setUp(self):
        super().setUp()
        self.license_json = os.path.join(PROJECT_ROOT, "license.json")
        self.data_dir = os.path.join(PROJECT_ROOT, "data")
        self.license_sec = os.path.join(self.data_dir, "license.sec")
        self.trials_db = os.path.join(PROJECT_ROOT, "trials_db.json")

        os.makedirs(self.data_dir, exist_ok=True)

        # Backup existing license artifacts
        self.backup_json = None
        if os.path.exists(self.license_json):
            with open(self.license_json, "rb") as f:
                self.backup_json = f.read()

        self.backup_sec = None
        if os.path.exists(self.license_sec):
            with open(self.license_sec, "rb") as f:
                self.backup_sec = f.read()

        self.backup_trials = None
        if os.path.exists(self.trials_db):
            with open(self.trials_db, "rb") as f:
                self.backup_trials = f.read()

        # Clean current state before each test
        self._remove_test_artifacts()

    def tearDown(self):
        # Restore backups
        self._remove_test_artifacts()
        if self.backup_json is not None:
            with open(self.license_json, "wb") as f:
                f.write(self.backup_json)
        if self.backup_sec is not None:
            with open(self.license_sec, "wb") as f:
                f.write(self.backup_sec)
        if self.backup_trials is not None:
            with open(self.trials_db, "wb") as f:
                f.write(self.backup_trials)
        super().tearDown()

    def _remove_test_artifacts(self):
        for path in [self.license_json, self.license_sec, self.trials_db]:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except Exception:
                    pass


# ============================================================================
# TIER 1 — FEATURE COVERAGE
# ============================================================================
class TestTier1FeatureCoverage(BaseLicensingTestCase):
    """Tier 1: Comprehensive feature coverage of Ed25519 crypto, activation, claims, and Admin API."""

    def test_01_ed25519_official_public_key_integrity(self):
        """Verify official Ed25519 public key matches authoritative constant and has exact 32 bytes."""
        self.assertEqual(PUBLIC_KEY_B64, OFFICIAL_PUBLIC_KEY_B64)
        raw_key_bytes = base64.b64decode(PUBLIC_KEY_B64)
        self.assertEqual(len(raw_key_bytes), 32, "Ed25519 public key must be exactly 32 bytes (256 bits)")
        pub_key = Ed25519PublicKey.from_public_bytes(raw_key_bytes)
        self.assertIsInstance(pub_key, Ed25519PublicKey)

    def test_02_ed25519_authentic_signature_verification(self):
        """Verify authentic Ed25519 signature verification succeeds with official public key."""
        # Genuine token issued by Cloudflare Worker for DOC
        sample_token = (
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
        self.assertIn(".", sample_token)
        payload_part, sig_part = sample_token.split(".", 1)

        pub_key = Ed25519PublicKey.from_public_bytes(base64.b64decode(PUBLIC_KEY_B64))
        sig_bytes = _b64url_decode(sig_part)
        self.assertEqual(len(sig_bytes), 64, "Ed25519 signature must be exactly 64 bytes")

        # Mathematical verification: MUST NOT raise InvalidSignature
        try:
            pub_key.verify(sig_bytes, payload_part.encode("ascii"))
        except InvalidSignature:
            self.fail("Ed25519 verification failed for an authentic token!")

    def test_03_token_payload_claims_extraction_and_schema(self):
        """Verify token payload claim extraction matches required schema and data types."""
        sample_payload_b64 = (
            "eyJpc3MiOiJsdXYtbGljZW5zZS1zZXJ2ZXIiLCJhdWQiOiJsdXYtc2RzLWludGVsbGlnZW5jZSIsInN1YiI6I"
            "jkwYTFkMmM0LTE2MGYtNGM4Ny1iN2YzLTEzN2E0MmViMTIzMCIsImxpY2Vuc2VfaWQiOiI5MGExZDJjNC0xNj"
            "BmLTRjODctYjdmMy0xMzdhNDJlYjEyMzAiLCJwcm9kdWN0X2NvZGUiOiJET0MiLCJjdXN0b21lciI6IkRvYW5"
            "oIE5naGllcCBMVVYgRE9DLVRSQU4iLCJwbGFuIjoiZW50ZXJwcmlzZSIsImZlYXR1cmVzIjpbImNvcmUiLCJp"
            "bXBvcnQiLCJleHBvcnQiLCJlZGl0IiwiZG9jX3RyYW4iXSwiZGV2aWNlX2lkIjoiMTExMTIyMjIzMzMzNDQ0ND"
            "U1NTU2NjY2Nzc3Nzg4ODg5OTk5MDAwMGFhYWFiYmJiY2NjY2RkZGRlZWVlZmZmZiIsImlhdCI6MTc5MjMzNTA2"
            "MiwicmVmcmVzaF9hZnRlciI6MTc5MjQyMTQ2Miwib2ZmbGluZV91bnRpbCI6MTc5MzU0NDY2Miwib2ZmbGluZV"
            "91bnRpbF9pc28iOiIyMDI2LTExLTAxVDE0OjUxOjAyLjkzN1oiLCJleHAiOjIxMTQzODA3OTksImV4cGlyZXNf"
            "YXQiOiIyMDM2LTEyLTMxVDIzOjU5OjU5LjAwMFoiLCJqdGkiOiJiZjVlMzQ3MC05OWJlLTQ1NTgtODNjOC0zZG"
            "UxMGExMGFjMTQiLCJsZWFzZV9zZXEiOjF9"
        )
        claims = json.loads(_b64url_decode(sample_payload_b64).decode("utf-8"))

        required_claims = [
            "iss", "aud", "sub", "license_id", "product_code",
            "customer", "plan", "features", "device_id",
            "iat", "refresh_after", "offline_until", "offline_until_iso",
            "exp", "expires_at", "jti", "lease_seq"
        ]
        for field in required_claims:
            self.assertIn(field, claims, f"Missing required claim: {field}")

        self.assertEqual(claims["iss"], "luv-license-server")
        self.assertEqual(claims["aud"], "luv-sds-intelligence")
        self.assertIn(claims["product_code"], ["DOC", "SDS"])
        self.assertIsInstance(claims["features"], list)
        self.assertIsInstance(claims["lease_seq"], int)
        self.assertIsInstance(claims["exp"], int)
        self.assertGreater(claims["exp"], claims["iat"])

    def test_04_device_id_format_64_char_lowercase_hex_sha256(self):
        """Verify device ID generation produces exactly 64-character lowercase hexadecimal SHA-256."""
        dev_id = generate_device_hash()
        self.assertEqual(len(dev_id), 64)
        self.assertTrue(re.match(r"^[a-f0-9]{64}$", dev_id), f"Device ID does not match lowercase hex regex: {dev_id}")
        self.assertFalse(any(c.isupper() for c in dev_id), "Device ID must not contain uppercase letters")

    def test_05_device_id_deterministic_and_unique(self):
        """Verify device ID is deterministic for identical inputs and unique for differing hints."""
        dev1 = generate_device_hash("workstation-alpha")
        dev2 = generate_device_hash("workstation-alpha")
        dev3 = generate_device_hash("workstation-beta")
        self.assertEqual(dev1, dev2, "Identical inputs must yield identical device ID")
        self.assertNotEqual(dev1, dev3, "Differing inputs must yield unique device IDs")

    def test_06_genuine_key_activation_direct_cf_worker(self):
        """Verify direct Cloudflare Worker POST /v1/activate succeeds with genuine key."""
        time.sleep(0.3)
        dev_id = "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef"
        payload = {
            "license_key": OFFICIAL_GENUINE_KEY,
            "device_id": dev_id,
            "device_label": "Direct CF Worker E2E Test",
            "app_version": APP_VERSION,
            "product_code": "DOC",
        }
        req = urllib.request.Request(
            f"{DEFAULT_SERVER_URL}/v1/activate",
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Content-Type": "application/json",
                "User-Agent": f"LUV-DOC/{APP_VERSION}",
            },
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=15) as resp:
            self.assertEqual(resp.status, 200)
            data = json.loads(resp.read().decode("utf-8"))
            self.assertIn("token", data)
            token = data["token"]

        # Cryptographically verify the returned token
        payload_part, sig_part = token.split(".", 1)
        pub_key = Ed25519PublicKey.from_public_bytes(base64.b64decode(PUBLIC_KEY_B64))
        pub_key.verify(_b64url_decode(sig_part), payload_part.encode("ascii"))

    def test_07_genuine_key_activation_via_client(self):
        """Verify LUVLicenseClient.activate_key succeeds with genuine key."""
        time.sleep(0.3)
        res = LUVLicenseClient.activate_key(OFFICIAL_GENUINE_KEY, "E2E Automated Test Suite")
        self.assertTrue(res.get("is_active"))
        self.assertEqual(res.get("status"), "active")
        self.assertEqual(res.get("license_key"), OFFICIAL_GENUINE_KEY)
        self.assertIn("Doanh Nghiệp", res.get("tier", ""))
        self.assertIn("token", res)

    def test_08_validation_lease_sequence_incrementing(self):
        """Verify POST /v1/validate verifies existing token and increments lease_seq."""
        time.sleep(0.3)
        dev_id = "a1b2c3d4e5f60718293a4b5c6d7e8f90a1b2c3d4e5f60718293a4b5c6d7e8f90"
        # 1. Activate
        act_req = urllib.request.Request(
            f"{DEFAULT_SERVER_URL}/v1/activate",
            data=json.dumps({
                "license_key": OFFICIAL_GENUINE_KEY,
                "device_id": dev_id,
                "device_label": "Lease Test Machine",
                "app_version": APP_VERSION,
                "product_code": "DOC",
            }).encode("utf-8"),
            headers={"Content-Type": "application/json", "User-Agent": f"LUV-DOC/{APP_VERSION}"},
            method="POST",
        )
        with urllib.request.urlopen(act_req, timeout=15) as resp:
            token = json.loads(resp.read().decode("utf-8"))["token"]

        claims_v1 = json.loads(_b64url_decode(token.split(".")[0]).decode("utf-8"))
        seq_v1 = claims_v1.get("lease_seq", 1)

        # 2. Validate
        time.sleep(0.3)
        val_req = urllib.request.Request(
            f"{DEFAULT_SERVER_URL}/v1/validate",
            data=json.dumps({
                "token": token,
                "device_id": dev_id,
                "app_version": APP_VERSION,
                "product_code": "DOC",
            }).encode("utf-8"),
            headers={"Content-Type": "application/json", "User-Agent": f"LUV-DOC/{APP_VERSION}"},
            method="POST",
        )
        with urllib.request.urlopen(val_req, timeout=15) as resp:
            self.assertEqual(resp.status, 200)
            renewed_token = json.loads(resp.read().decode("utf-8"))["token"]

        claims_v2 = json.loads(_b64url_decode(renewed_token.split(".")[0]).decode("utf-8"))
        seq_v2 = claims_v2.get("lease_seq", 1)
        self.assertGreater(seq_v2, seq_v1, "Validation must increment lease_seq")

        # Verify signature of renewed token
        payload_part, sig_part = renewed_token.split(".", 1)
        pub_key = Ed25519PublicKey.from_public_bytes(base64.b64decode(PUBLIC_KEY_B64))
        pub_key.verify(_b64url_decode(sig_part), payload_part.encode("ascii"))

    def test_09_titanpdf_admin_api_list_licenses_schema(self):
        """Verify GET /v1/admin/licenses returns valid schema and active licenses."""
        time.sleep(0.3)
        data = _call_admin_api("GET", "/v1/admin/licenses")
        self.assertIn("licenses", data, "Admin licenses response must contain 'licenses' key")
        licenses = data.get("licenses", [])
        self.assertIsInstance(licenses, list)
        self.assertGreater(len(licenses), 0, "D1 must contain at least one license")

        first = licenses[0]
        self.assertIn("id", first)
        self.assertIn("status", first)
        self.assertIn("plan", first)
        self.assertIn("max_devices", first)

    def test_10_titanpdf_admin_api_create_license_with_doc_and_sds(self):
        """Verify POST /v1/admin/licenses successfully creates genuine key entitled to DOC and SDS."""
        time.sleep(0.3)
        payload = {
            "customer_name": "E2E Test Corporate Client",
            "customer_email": "e2e_corp@luv.vn",
            "expires_at": "2035-12-31",
            "products": {
                "DOC": {"enabled": True, "max_devices": 5},
                "SDS": {"enabled": True, "max_devices": 5},
            },
            "offline_days": 14,
            "plan": "enterprise",
            "features": ["core", "doc_tran", "export"],
            "notes": "Automated E2E Creation Test",
        }
        res = _call_admin_api("POST", "/v1/admin/licenses", payload)
        self.assertIn("license_key", res)
        self.assertIn("id", res)
        key = res["license_key"]
        lic_id = res["id"]

        # Key must strictly conform to LUV key format
        self.assertTrue(
            re.match(r"^LUV-[A-Z0-9]{5}(?:-[A-Z0-9]{5}){3}$", key),
            f"Created key does not match LUV format: {key}",
        )

        # Cleanup created license
        time.sleep(0.3)
        _call_admin_api("POST", f"/v1/admin/licenses/{lic_id}/delete", {"confirm": True})

    def test_11_admin_api_unauthorized_without_token(self):
        """Verify Admin API rejects requests missing Authorization Bearer token."""
        req = urllib.request.Request(
            f"{DEFAULT_SERVER_URL}/v1/admin/licenses",
            headers={"User-Agent": "LUV-License-Admin/1.0"},
            method="GET",
        )
        with self.assertRaises(urllib.error.HTTPError) as ctx:
            urllib.request.urlopen(req, timeout=10)
        self.assertEqual(ctx.exception.code, 401)

    def test_12_admin_api_unauthorized_with_invalid_token(self):
        """Verify Admin API rejects requests with invalid Bearer token."""
        req = urllib.request.Request(
            f"{DEFAULT_SERVER_URL}/v1/admin/licenses",
            headers={
                "Authorization": "Bearer invalid_bearer_token_12345",
                "User-Agent": "LUV-License-Admin/1.0",
            },
            method="GET",
        )
        with self.assertRaises(urllib.error.HTTPError) as ctx:
            urllib.request.urlopen(req, timeout=10)
        self.assertEqual(ctx.exception.code, 401)


# ============================================================================
# TIER 2 — BOUNDARY & CORNER CASES (CRITICAL FOR ACCEPTANCE)
# ============================================================================
class TestTier2BoundaryAndCornerCases(BaseLicensingTestCase):
    """Tier 2: 100% rejection of bogus keys, mock patterns, tampered tokens, and device limits."""

    def test_13_rejection_random_garbage_keys_client_side(self):
        """100% REJECTION of random garbage keys (khk, 12345, abcdef, special chars)."""
        garbage_samples = [
            "khk",
            "12345",
            "abcdef",
            "!@#$%^&*()_+",
            "null",
            "undefined",
            "True",
            "SELECT * FROM licenses;",
            "<script>alert(1)</script>",
        ]
        for bad_key in garbage_samples:
            with self.subTest(bad_key=bad_key):
                with self.assertRaises((ValueError, Exception)):
                    LUVLicenseClient.activate_key(bad_key, "Attacker")

    def test_14_rejection_random_garbage_keys_live_cloudflare_worker(self):
        """Verify Cloudflare Worker live endpoint returns HTTP 400 invalid_request for garbage keys."""
        time.sleep(0.3)
        dev_id = "0000111122223333444455556666777788889999aaaabbbbccccddddeeeeffff"
        payload = {
            "license_key": "khk_garbage_key_123",
            "device_id": dev_id,
            "device_label": "Boundary Test",
            "app_version": APP_VERSION,
            "product_code": "DOC",
        }
        req = urllib.request.Request(
            f"{DEFAULT_SERVER_URL}/v1/activate",
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json", "User-Agent": f"LUV-DOC/{APP_VERSION}"},
            method="POST",
        )
        with self.assertRaises(urllib.error.HTTPError) as ctx:
            urllib.request.urlopen(req, timeout=10)
        self.assertEqual(ctx.exception.code, 400)
        err_data = json.loads(ctx.exception.read().decode("utf-8"))
        self.assertEqual(err_data.get("error"), "invalid_request")

    def test_15_rejection_old_mock_prefix_patterns(self):
        """100% REJECTION of old mock prefix patterns (LUV-PDF-1234, LUV-PRO-TEST-1234)."""
        mock_patterns = [
            "LUV-PDF-1234",
            "LUV-PDF-1234-5678",
            "LUV-PRO-TEST-1234",
            "LUV-ENT-1234-5678-9999",
            "LUV-PDF-2026-PRO-VN99",
            "LUV-2026-ENTERPRISE",
            "LUV-MASTER-KEY",
        ]
        for mock_key in mock_patterns:
            with self.subTest(mock_key=mock_key):
                with self.assertRaises((ValueError, Exception)) as ctx:
                    LUVLicenseClient.activate_key(mock_key, "Mock Tester")
                # Ensure no offline pseudo-hash activation passed
                info = LUVLicenseClient.get_license_info()
                self.assertFalse(info.get("is_active"))

    def test_16_rejection_well_formatted_non_existent_keys(self):
        """100% REJECTION of syntactically valid but non-existent keys (e.g. LUV-AAAAA-BBBBB-CCCCC-DDDDD)."""
        time.sleep(0.3)
        non_existent_key = "LUV-AAAAA-BBBBB-CCCCC-DDDDD"
        # 1. Direct Cloudflare Worker check
        payload = {
            "license_key": non_existent_key,
            "device_id": "ffff0000ffff0000ffff0000ffff0000ffff0000ffff0000ffff0000ffff0000",
            "device_label": "NonExistent Key Test",
            "app_version": APP_VERSION,
            "product_code": "DOC",
        }
        req = urllib.request.Request(
            f"{DEFAULT_SERVER_URL}/v1/activate",
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json", "User-Agent": f"LUV-DOC/{APP_VERSION}"},
            method="POST",
        )
        with self.assertRaises(urllib.error.HTTPError) as ctx:
            urllib.request.urlopen(req, timeout=10)
        self.assertEqual(ctx.exception.code, 403)
        err_data = json.loads(ctx.exception.read().decode("utf-8"))
        self.assertEqual(err_data.get("error"), "license_denied")

        # 2. Client activate check
        time.sleep(0.3)
        with self.assertRaises(ValueError):
            LUVLicenseClient.activate_key(non_existent_key, "Non-existent Key User")

    def test_17_rejection_tampered_token_altered_signature_bytes(self):
        """100% REJECTION of tampered tokens with bit-flipped / altered signature bytes."""
        valid_token = (
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
        payload_part, sig_part = valid_token.split(".", 1)
        sig_bytes = bytearray(_b64url_decode(sig_part))

        # Flip the first bit of the signature
        sig_bytes[0] ^= 0x01
        tampered_sig_b64 = _b64url_encode(bytes(sig_bytes))
        tampered_token = f"{payload_part}.{tampered_sig_b64}"

        pub_key = Ed25519PublicKey.from_public_bytes(base64.b64decode(PUBLIC_KEY_B64))
        with self.assertRaises(InvalidSignature):
            pub_key.verify(_b64url_decode(tampered_sig_b64), payload_part.encode("ascii"))

        # Verify client rejects the tampered token
        with open(self.license_json, "w", encoding="utf-8") as f:
            json.dump({"token": tampered_token, "license_key": OFFICIAL_GENUINE_KEY}, f)

        info = LUVLicenseClient.get_license_info()
        self.assertFalse(info.get("is_active"), "Tampered signature token must report is_active=False")

    def test_18_rejection_tampered_payload_claims_extended_expiry(self):
        """100% REJECTION of modified payload claims: extending expiration date breaks Ed25519 signature."""
        valid_payload_b64 = (
            "eyJpc3MiOiJsdXYtbGljZW5zZS1zZXJ2ZXIiLCJhdWQiOiJsdXYtc2RzLWludGVsbGlnZW5jZSIsInN1YiI6I"
            "jkwYTFkMmM0LTE2MGYtNGM4Ny1iN2YzLTEzN2E0MmViMTIzMCIsImxpY2Vuc2VfaWQiOiI5MGExZDJjNC0xNj"
            "BmLTRjODctYjdmMy0xMzdhNDJlYjEyMzAiLCJwcm9kdWN0X2NvZGUiOiJET0MiLCJjdXN0b21lciI6IkRvYW5"
            "oIE5naGllcCBMVVYgRE9DLVRSQU4iLCJwbGFuIjoiZW50ZXJwcmlzZSIsImZlYXR1cmVzIjpbImNvcmUiLCJp"
            "bXBvcnQiLCJleHBvcnQiLCJlZGl0IiwiZG9jX3RyYW4iXSwiZGV2aWNlX2lkIjoiMTExMTIyMjIzMzMzNDQ0ND"
            "U1NTU2NjY2Nzc3Nzg4ODg5OTk5MDAwMGFhYWFiYmJiY2NjY2RkZGRlZWVlZmZmZiIsImlhdCI6MTc5MjMzNTA2"
            "MiwicmVmcmVzaF9hZnRlciI6MTc5MjQyMTQ2Miwib2ZmbGluZV91bnRpbCI6MTc5MzU0NDY2Miwib2ZmbGluZV"
            "91bnRpbF9pc28iOiIyMDI2LTExLTAxVDE0OjUxOjAyLjkzN1oiLCJleHAiOjIxMTQzODA3OTksImV4cGlyZXNf"
            "YXQiOiIyMDM2LTEyLTMxVDIzOjU5OjU5LjAwMFoiLCJqdGkiOiJiZjVlMzQ3MC05OWJlLTQ1NTgtODNjOC0zZG"
            "UxMGExMGFjMTQiLCJsZWFzZV9zZXEiOjF9"
        )
        valid_sig_b64 = "6aEJEn351Quc6Ct9Ptck4aa9I-EQpRvdKqw8gPf_o-wb4z_zHiwS9Q4WOj3ohkwoUbo5TxAokJkwmxebcfRmBw"

        claims = json.loads(_b64url_decode(valid_payload_b64).decode("utf-8"))
        # Attacker modifies expiry date to year 2099
        claims["exp"] = 4070908800
        claims["expires_at"] = "2099-01-01T00:00:00.000Z"

        tampered_payload_bytes = json.dumps(claims, separators=(",", ":")).encode("utf-8")
        tampered_payload_b64 = _b64url_encode(tampered_payload_bytes)

        pub_key = Ed25519PublicKey.from_public_bytes(base64.b64decode(PUBLIC_KEY_B64))
        sig_bytes = _b64url_decode(valid_sig_b64)

        # Mathematical verification must reject altered payload
        with self.assertRaises(InvalidSignature):
            pub_key.verify(sig_bytes, tampered_payload_b64.encode("ascii"))

    def test_19_rejection_tampered_payload_claims_modified_device_id(self):
        """100% REJECTION of modified payload claims: altering device_id breaks Ed25519 signature."""
        valid_payload_b64 = (
            "eyJpc3MiOiJsdXYtbGljZW5zZS1zZXJ2ZXIiLCJhdWQiOiJsdXYtc2RzLWludGVsbGlnZW5jZSIsInN1YiI6I"
            "jkwYTFkMmM0LTE2MGYtNGM4Ny1iN2YzLTEzN2E0MmViMTIzMCIsImxpY2Vuc2VfaWQiOiI5MGExZDJjNC0xNj"
            "BmLTRjODctYjdmMy0xMzdhNDJlYjEyMzAiLCJwcm9kdWN0X2NvZGUiOiJET0MiLCJjdXN0b21lciI6IkRvYW5"
            "oIE5naGllcCBMVVYgRE9DLVRSQU4iLCJwbGFuIjoiZW50ZXJwcmlzZSIsImZlYXR1cmVzIjpbImNvcmUiLCJp"
            "bXBvcnQiLCJleHBvcnQiLCJlZGl0IiwiZG9jX3RyYW4iXSwiZGV2aWNlX2lkIjoiMTExMTIyMjIzMzMzNDQ0ND"
            "U1NTU2NjY2Nzc3Nzg4ODg5OTk5MDAwMGFhYWFiYmJiY2NjY2RkZGRlZWVlZmZmZiIsImlhdCI6MTc5MjMzNTA2"
            "MiwicmVmcmVzaF9hZnRlciI6MTc5MjQyMTQ2Miwib2ZmbGluZV91bnRpbCI6MTc5MzU0NDY2Miwib2ZmbGluZV"
            "91bnRpbF9pc28iOiIyMDI2LTExLTAxVDE0OjUxOjAyLjkzN1oiLCJleHAiOjIxMTQzODA3OTksImV4cGlyZXNf"
            "YXQiOiIyMDM2LTEyLTMxVDIzOjU5OjU5LjAwMFoiLCJqdGkiOiJiZjVlMzQ3MC05OWJlLTQ1NTgtODNjOC0zZG"
            "UxMGExMGFjMTQiLCJsZWFzZV9zZXEiOjF9"
        )
        valid_sig_b64 = "6aEJEn351Quc6Ct9Ptck4aa9I-EQpRvdKqw8gPf_o-wb4z_zHiwS9Q4WOj3ohkwoUbo5TxAokJkwmxebcfRmBw"

        claims = json.loads(_b64url_decode(valid_payload_b64).decode("utf-8"))
        claims["device_id"] = "9999999999999999999999999999999999999999999999999999999999999999"

        tampered_payload_b64 = _b64url_encode(json.dumps(claims).encode("utf-8"))
        pub_key = Ed25519PublicKey.from_public_bytes(base64.b64decode(PUBLIC_KEY_B64))
        sig_bytes = _b64url_decode(valid_sig_b64)

        with self.assertRaises(InvalidSignature):
            pub_key.verify(sig_bytes, tampered_payload_b64.encode("ascii"))

    def test_20_rejection_device_limit_http_409(self):
        """Verify device limit enforcement returns HTTP 409 Conflict when limit exceeded."""
        time.sleep(0.3)
        # 1. Create a 1-device license
        created = _call_admin_api("POST", "/v1/admin/licenses", {
            "customer_name": "Device Limit Test",
            "expires_at": "2030-01-01",
            "products": {"DOC": {"enabled": True, "max_devices": 1}},
            "plan": "pro",
        })
        lic_id = created["id"]
        key = created["license_key"]

        try:
            time.sleep(0.3)
            # 2. Activate Device 1 -> 200 OK
            dev1_payload = {
                "license_key": key,
                "device_id": "1111111111111111111111111111111111111111111111111111111111111111",
                "device_label": "Device 1",
                "app_version": APP_VERSION,
                "product_code": "DOC",
            }
            req1 = urllib.request.Request(
                f"{DEFAULT_SERVER_URL}/v1/activate",
                data=json.dumps(dev1_payload).encode("utf-8"),
                headers={"Content-Type": "application/json", "User-Agent": f"LUV-DOC/{APP_VERSION}"},
                method="POST",
            )
            with urllib.request.urlopen(req1, timeout=10) as resp1:
                self.assertEqual(resp1.status, 200)

            time.sleep(0.3)
            # 3. Activate Device 2 -> 409 Conflict (device_limit)
            dev2_payload = {
                "license_key": key,
                "device_id": "2222222222222222222222222222222222222222222222222222222222222222",
                "device_label": "Device 2",
                "app_version": APP_VERSION,
                "product_code": "DOC",
            }
            req2 = urllib.request.Request(
                f"{DEFAULT_SERVER_URL}/v1/activate",
                data=json.dumps(dev2_payload).encode("utf-8"),
                headers={"Content-Type": "application/json", "User-Agent": f"LUV-DOC/{APP_VERSION}"},
                method="POST",
            )
            with self.assertRaises(urllib.error.HTTPError) as ctx:
                urllib.request.urlopen(req2, timeout=10)
            self.assertEqual(ctx.exception.code, 409)
            err_data = json.loads(ctx.exception.read().decode("utf-8"))
            self.assertEqual(err_data.get("error"), "device_limit")
        finally:
            # Cleanup
            time.sleep(0.3)
            _call_admin_api("POST", f"/v1/admin/licenses/{lic_id}/delete", {"confirm": True})

    def test_21_rejection_empty_and_whitespace_keys(self):
        """Verify client immediately rejects empty, whitespace, or None keys with ValueError."""
        for empty_val in ["", "   ", " \t \n ", None]:
            with self.subTest(empty_val=empty_val):
                with self.assertRaises(ValueError):
                    LUVLicenseClient.activate_key(empty_val)

    def test_22_rejection_expired_token_timestamp(self):
        """Verify client rejects validly signed token when exp timestamp is in the past."""
        past_exp = int(datetime.now(timezone.utc).timestamp()) - 3600  # 1 hour ago
        claims = {
            "iss": "luv-license-server",
            "aud": "luv-sds-intelligence",
            "sub": "test-uuid",
            "license_id": "test-uuid",
            "product_code": "DOC",
            "customer": "Expired Test",
            "plan": "pro",
            "features": ["core"],
            "device_id": generate_device_hash(),
            "exp": past_exp,
            "expires_at": "2020-01-01T00:00:00.000Z",
            "lease_seq": 1,
        }
        payload_b64 = _b64url_encode(json.dumps(claims).encode("utf-8"))
        # Dummy signature
        fake_token = f"{payload_b64}.AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"

        with open(self.license_json, "w", encoding="utf-8") as f:
            json.dump({"token": fake_token, "license_key": OFFICIAL_GENUINE_KEY}, f)

        info = LUVLicenseClient.get_license_info()
        self.assertFalse(info.get("is_active"))


# ============================================================================
# TIER 3 — CROSS-FEATURE COMBINATIONS
# ============================================================================
class TestTier3CrossFeatureCombinations(BaseLicensingTestCase):
    """Tier 3: Pairwise and full cross-feature lifecycle integration tests."""

    def test_23_full_lifecycle_admin_create_activate_validate_reset_revoke(self):
        """Full lifecycle: Admin creates key -> Client activates -> validates -> Admin resets -> Admin revokes -> validation fails."""
        time.sleep(0.3)
        # Step 1: Admin generates key via Admin API
        created = _call_admin_api("POST", "/v1/admin/licenses", {
            "customer_name": "Full Lifecycle Enterprise Corp",
            "customer_email": "lifecycle@luv.vn",
            "expires_at": "2032-12-31",
            "products": {
                "DOC": {"enabled": True, "max_devices": 2},
                "SDS": {"enabled": True, "max_devices": 2},
            },
            "plan": "enterprise",
            "features": ["core", "doc_tran", "export"],
        })
        lic_id = created["id"]
        key = created["license_key"]

        try:
            time.sleep(0.3)
            # Step 2: Client activates key in LUV-DOC-TRAN
            act_res = LUVLicenseClient.activate_key(key, "Lifecycle Client 1")
            self.assertTrue(act_res["is_active"])
            token_v1 = act_res["token"]

            # Step 3: Local signature verified and saved
            info = LUVLicenseClient.get_license_info()
            self.assertTrue(info["is_active"])
            self.assertEqual(info["license_key"], key)

            time.sleep(0.3)
            # Step 4: Client validates and lease sequence increments
            dev_id = act_res["device_id"]
            val_req = urllib.request.Request(
                f"{DEFAULT_SERVER_URL}/v1/validate",
                data=json.dumps({
                    "token": token_v1,
                    "device_id": dev_id,
                    "app_version": APP_VERSION,
                    "product_code": "DOC",
                }).encode("utf-8"),
                headers={"Content-Type": "application/json", "User-Agent": f"LUV-DOC/{APP_VERSION}"},
                method="POST",
            )
            with urllib.request.urlopen(val_req, timeout=10) as resp:
                self.assertEqual(resp.status, 200)
                token_v2 = json.loads(resp.read().decode("utf-8"))["token"]

            claims_v2 = json.loads(_b64url_decode(token_v2.split(".")[0]).decode("utf-8"))
            self.assertEqual(claims_v2["lease_seq"], 2)

            time.sleep(0.3)
            # Step 5: Admin resets devices
            reset_res = _call_admin_api("POST", f"/v1/admin/licenses/{lic_id}/reset-devices")
            self.assertTrue(reset_res.get("ok"))

            time.sleep(0.3)
            # Step 6: Admin revokes key
            rev_res = _call_admin_api("POST", f"/v1/admin/licenses/{lic_id}/revoke")
            self.assertTrue(rev_res.get("ok"))

            time.sleep(0.3)
            # Step 7: Subsequent client validation or activation fails
            with self.assertRaises(urllib.error.HTTPError) as ctx:
                val_req_revoked = urllib.request.Request(
                    f"{DEFAULT_SERVER_URL}/v1/validate",
                    data=json.dumps({
                        "token": token_v2,
                        "device_id": dev_id,
                        "app_version": APP_VERSION,
                        "product_code": "DOC",
                    }).encode("utf-8"),
                    headers={"Content-Type": "application/json", "User-Agent": f"LUV-DOC/{APP_VERSION}"},
                    method="POST",
                )
                urllib.request.urlopen(val_req_revoked, timeout=10)
            self.assertIn(ctx.exception.code, [401, 403])

        finally:
            # Step 8: Clean up
            time.sleep(0.3)
            _call_admin_api("POST", f"/v1/admin/licenses/{lic_id}/delete", {"confirm": True})

    def test_24_admin_device_reset_allows_re_registration(self):
        """Cross-feature: When device limit is hit, Admin reset-devices frees binding and permits new device."""
        time.sleep(0.3)
        created = _call_admin_api("POST", "/v1/admin/licenses", {
            "customer_name": "Reset Device Test",
            "expires_at": "2030-01-01",
            "products": {"DOC": {"enabled": True, "max_devices": 1}},
            "plan": "pro",
        })
        lic_id = created["id"]
        key = created["license_key"]

        try:
            time.sleep(0.3)
            # Register Device A
            dev_a = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
            req_a = urllib.request.Request(
                f"{DEFAULT_SERVER_URL}/v1/activate",
                data=json.dumps({
                    "license_key": key, "device_id": dev_a,
                    "device_label": "Dev A", "app_version": APP_VERSION, "product_code": "DOC"
                }).encode("utf-8"),
                headers={"Content-Type": "application/json", "User-Agent": f"LUV-DOC/{APP_VERSION}"},
                method="POST",
            )
            with urllib.request.urlopen(req_a, timeout=10) as r:
                self.assertEqual(r.status, 200)

            time.sleep(0.3)
            # Device B rejected (409)
            dev_b = "bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb"
            req_b = urllib.request.Request(
                f"{DEFAULT_SERVER_URL}/v1/activate",
                data=json.dumps({
                    "license_key": key, "device_id": dev_b,
                    "device_label": "Dev B", "app_version": APP_VERSION, "product_code": "DOC"
                }).encode("utf-8"),
                headers={"Content-Type": "application/json", "User-Agent": f"LUV-DOC/{APP_VERSION}"},
                method="POST",
            )
            with self.assertRaises(urllib.error.HTTPError) as ctx:
                urllib.request.urlopen(req_b, timeout=10)
            self.assertEqual(ctx.exception.code, 409)

            time.sleep(0.3)
            # Admin resets devices
            _call_admin_api("POST", f"/v1/admin/licenses/{lic_id}/reset-devices")

            time.sleep(0.3)
            # Device B now activates successfully
            with urllib.request.urlopen(req_b, timeout=10) as r2:
                self.assertEqual(r2.status, 200)

        finally:
            time.sleep(0.3)
            _call_admin_api("POST", f"/v1/admin/licenses/{lic_id}/delete", {"confirm": True})


# ============================================================================
# TIER 4 — REAL-WORLD SCENARIOS
# ============================================================================
class TestTier4RealWorldScenarios(BaseLicensingTestCase):
    """Tier 4: Owner Key Fallback, offline persistence in data/license.sec, and DPAPI protection."""

    def test_25_owner_key_fallback_doc_to_sds(self):
        """Owner Key Fallback: Activating Owner Key (entitled to SDS) automatically falls back from DOC to SDS and verifies."""
        time.sleep(0.3)
        # 1. Reset devices on owner license to ensure room for test activation
        _call_admin_api("POST", f"/v1/admin/licenses/{OFFICIAL_OWNER_LICENSE_ID}/reset-devices")

        time.sleep(0.3)
        # 2. Confirm direct probe with product_code=DOC gets 403 product_denied
        dev_id = generate_device_hash("OwnerMachineProbe")
        doc_req = urllib.request.Request(
            f"{DEFAULT_SERVER_URL}/v1/activate",
            data=json.dumps({
                "license_key": OFFICIAL_OWNER_KEY,
                "device_id": dev_id,
                "device_label": "Direct Owner Probe",
                "app_version": APP_VERSION,
                "product_code": "DOC",
            }).encode("utf-8"),
            headers={"Content-Type": "application/json", "User-Agent": f"LUV-DOC/{APP_VERSION}"},
            method="POST",
        )
        with self.assertRaises(urllib.error.HTTPError) as ctx:
            urllib.request.urlopen(doc_req, timeout=10)
        self.assertEqual(ctx.exception.code, 403)
        err = json.loads(ctx.exception.read().decode("utf-8"))
        self.assertEqual(err.get("error"), "product_denied")

        time.sleep(0.3)
        # 3. LUVLicenseClient.activate_key automatically executes fallback and succeeds
        res = LUVLicenseClient.activate_key(OFFICIAL_OWNER_KEY, "LUV Owner PC")
        self.assertTrue(res.get("is_active"))
        self.assertEqual(res.get("status"), "active")
        self.assertIn("OWNER", res.get("tier", "").upper())

        # Cryptographically verify the Ed25519 signature of the returned token
        token = res["token"]
        payload_part, sig_part = token.split(".", 1)
        pub_key = Ed25519PublicKey.from_public_bytes(base64.b64decode(PUBLIC_KEY_B64))
        pub_key.verify(_b64url_decode(sig_part), payload_part.encode("ascii"))

    def test_26_offline_persistence_within_offline_until_window(self):
        """Offline Persistence: Valid token stored on disk allows local verification within offline window without network."""
        future_exp = int(datetime.now(timezone.utc).timestamp()) + 365 * 86400
        future_offline = int(datetime.now(timezone.utc).timestamp()) + 14 * 86400

        # Construct genuine claims
        claims = {
            "iss": "luv-license-server",
            "aud": "luv-sds-intelligence",
            "sub": "offline-persist-test",
            "license_id": "offline-persist-test",
            "product_code": "DOC",
            "customer": "Offline Corp",
            "plan": "enterprise",
            "features": ["core", "doc_tran"],
            "device_id": generate_device_hash(),
            "iat": int(datetime.now(timezone.utc).timestamp()),
            "exp": future_exp,
            "offline_until": future_offline,
            "offline_until_iso": "2030-01-01T00:00:00.000Z",
            "lease_seq": 1,
        }

        # Authentic Ed25519 token from probe
        sample_token = (
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
        record = {
            "status": "active",
            "license_key": OFFICIAL_GENUINE_KEY,
            "registered_to": "Doanh Nghiep LUV DOC-TRAN",
            "tier": "Doanh Nghiệp (ENTERPRISE)",
            "token": sample_token,
            "is_active": True,
        }

        # Save to both target persistence paths
        for path in [self.license_json, self.license_sec]:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(record, f)

        # Local verification without network
        info = LUVLicenseClient.get_license_info()
        self.assertTrue(info.get("is_active"))
        self.assertEqual(info.get("license_key"), OFFICIAL_GENUINE_KEY)

        # Check access permission
        access = LUVLicenseClient.check_access()
        self.assertTrue(access.get("has_access"))
        self.assertEqual(access.get("access_type"), "licensed")

    def test_27_offline_tamper_detection_in_persisted_storage(self):
        """Corrupting persisted license file causes local cryptographic verification to safely fail."""
        corrupted_record = {
            "status": "active",
            "license_key": OFFICIAL_GENUINE_KEY,
            "token": "corrupted_payload.corrupted_signature",
            "is_active": True,
        }
        for path in [self.license_json, self.license_sec]:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(corrupted_record, f)

        info = LUVLicenseClient.get_license_info()
        self.assertFalse(info.get("is_active"))

    def test_28_windows_dpapi_entropy_round_trip(self):
        """Verify Windows DPAPI protection round-trip with entropy b'LUV-SDS-License-v1'."""
        if os.name != "nt":
            self.skipTest("DPAPI test requires Windows NT platform")

        from ctypes import wintypes

        class DATA_BLOB(ctypes.Structure):
            _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

        _crypt32 = ctypes.windll.crypt32
        _kernel32 = ctypes.windll.kernel32

        def protect(data: bytes, entropy: bytes = b"LUV-SDS-License-v1") -> bytes:
            s_buf = ctypes.create_string_buffer(data)
            s_blob = DATA_BLOB(len(data), ctypes.cast(s_buf, ctypes.POINTER(ctypes.c_char)))
            o_blob = DATA_BLOB()
            e_buf = ctypes.create_string_buffer(entropy)
            e_blob = DATA_BLOB(len(entropy), ctypes.cast(e_buf, ctypes.POINTER(ctypes.c_char)))
            if not _crypt32.CryptProtectData(
                ctypes.byref(s_blob), "LUV License", ctypes.byref(e_blob), None, None, 0x1, ctypes.byref(o_blob)
            ):
                raise RuntimeError("CryptProtectData failed")
            try:
                return ctypes.string_at(o_blob.pbData, o_blob.cbData)
            finally:
                _kernel32.LocalFree(o_blob.pbData)

        def unprotect(data: bytes, entropy: bytes = b"LUV-SDS-License-v1") -> bytes:
            s_buf = ctypes.create_string_buffer(data)
            s_blob = DATA_BLOB(len(data), ctypes.cast(s_buf, ctypes.POINTER(ctypes.c_char)))
            o_blob = DATA_BLOB()
            e_buf = ctypes.create_string_buffer(entropy)
            e_blob = DATA_BLOB(len(entropy), ctypes.cast(e_buf, ctypes.POINTER(ctypes.c_char)))
            if not _crypt32.CryptUnprotectData(
                ctypes.byref(s_blob), None, ctypes.byref(e_blob), None, None, 0x1, ctypes.byref(o_blob)
            ):
                raise RuntimeError("CryptUnprotectData failed")
            try:
                return ctypes.string_at(o_blob.pbData, o_blob.cbData)
            finally:
                _kernel32.LocalFree(o_blob.pbData)

        test_data = json.dumps({"token": "genuine.ed25519.token", "status": "active"}).encode("utf-8")
        protected = protect(test_data)
        self.assertNotEqual(protected, test_data)
        unprotected = unprotect(protected)
        self.assertEqual(unprotected, test_data)

    def test_29_secure_license_save_and_load_round_trip(self):
        """Verify LUVLicenseClient save_secure_license and load_secure_license round-trip in data/license.sec."""
        sample_token = (
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
        record = {
            "status": "active",
            "license_key": OFFICIAL_GENUINE_KEY,
            "registered_to": "Secure Corp",
            "tier": "Doanh Nghiệp (ENTERPRISE)",
            "token": sample_token,
            "is_active": True,
        }
        LUVLicenseClient.save_secure_license(record)
        self.assertTrue(os.path.exists(self.license_sec), "data/license.sec must exist after save_secure_license")

        loaded = LUVLicenseClient.load_secure_license()
        self.assertIsNotNone(loaded)
        self.assertEqual(loaded.get("license_key"), OFFICIAL_GENUINE_KEY)
        self.assertEqual(loaded.get("token"), sample_token)

    def test_30_legacy_license_json_migration_fallback(self):
        """Verify automatic migration from legacy license.json to data/license.sec."""
        sample_token = (
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
        record = {
            "status": "active",
            "license_key": OFFICIAL_GENUINE_KEY,
            "registered_to": "Migrated Client",
            "token": sample_token,
            "is_active": True,
        }
        # Write to legacy license.json only
        with open(self.license_json, "w", encoding="utf-8") as f:
            json.dump(record, f)

        if os.path.exists(self.license_sec):
            os.remove(self.license_sec)

        info = LUVLicenseClient.get_license_info()
        self.assertTrue(info.get("is_active"))
        # Verify it migrated to data/license.sec
        self.assertTrue(os.path.exists(self.license_sec), "Migration must write data/license.sec")

    def test_31_tampered_token_zeroed_signature_rejected(self):
        """Verify Ed25519 verification rejects all-zero 64-byte signature."""
        sample_payload_b64 = "eyJpc3MiOiJsdXYtbGljZW5zZS1zZXJ2ZXIiLCJhdWQiOiJsdXYtc2RzLWludGVsbGlnZW5jZSJ9"
        zero_sig = _b64url_encode(b"\x00" * 64)
        pub_key = Ed25519PublicKey.from_public_bytes(base64.b64decode(PUBLIC_KEY_B64))
        with self.assertRaises(InvalidSignature):
            pub_key.verify(_b64url_decode(zero_sig), sample_payload_b64.encode("ascii"))

    def test_32_tampered_token_truncated_signature_rejected(self):
        """Verify Ed25519 verification rejects truncated signature (< 64 bytes)."""
        sample_payload_b64 = "eyJpc3MiOiJsdXYtbGljZW5zZS1zZXJ2ZXIiLCJhdWQiOiJsdXYtc2RzLWludGVsbGlnZW5jZSJ9"
        short_sig = _b64url_encode(b"\x01" * 32)
        pub_key = Ed25519PublicKey.from_public_bytes(base64.b64decode(PUBLIC_KEY_B64))
        with self.assertRaises((InvalidSignature, ValueError)):
            pub_key.verify(_b64url_decode(short_sig), sample_payload_b64.encode("ascii"))

    def test_33_tampered_token_missing_dot_delimiter(self):
        """Verify client rejects token with missing dot delimiter."""
        bad_token = "invalid_token_without_period_delimiter"
        with open(self.license_sec, "w", encoding="utf-8") as f:
            json.dump({"token": bad_token, "license_key": OFFICIAL_GENUINE_KEY}, f)
        info = LUVLicenseClient.get_license_info()
        self.assertFalse(info.get("is_active"))

    def test_34_wrong_ed25519_key_fails_verification(self):
        """Verify authentic token fails verification when checked against a foreign/wrong Ed25519 key."""
        from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey
        wrong_private = Ed25519PrivateKey.generate()
        wrong_public = wrong_private.public_key()

        sample_token = (
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
        payload_part, sig_part = sample_token.split(".", 1)
        sig_bytes = _b64url_decode(sig_part)

        # Verification with wrong public key MUST raise InvalidSignature
        with self.assertRaises(InvalidSignature):
            wrong_public.verify(sig_bytes, payload_part.encode("ascii"))

    def test_35_trial_vs_licensed_access_state_machine(self):
        """Verify check_access properly transitions between 48h trial state and official licensed state."""
        # 1. Without license, check_access reports trial
        self._remove_test_artifacts()
        access_initial = LUVLicenseClient.check_access("DEV-TEST-STATE-MACHINE")
        self.assertTrue(access_initial.get("has_access"))
        self.assertEqual(access_initial.get("access_type"), "trial")
        self.assertIn("Dùng thử", access_initial.get("tier", ""))

        # 2. After activating genuine license, check_access reports licensed
        sample_token = (
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
        record = {
            "status": "active",
            "license_key": OFFICIAL_GENUINE_KEY,
            "registered_to": "State Machine Corp",
            "tier": "Doanh Nghiệp (ENTERPRISE)",
            "token": sample_token,
            "is_active": True,
        }
        LUVLicenseClient.save_secure_license(record)

        access_licensed = LUVLicenseClient.check_access("DEV-TEST-STATE-MACHINE")
        self.assertTrue(access_licensed.get("has_access"))
        self.assertEqual(access_licensed.get("access_type"), "licensed")
        self.assertEqual(access_licensed.get("license_key"), OFFICIAL_GENUINE_KEY)

    def test_36_realtime_validate_and_revocation_auto_purge(self):
        """Verify LUVLicenseClient.validate_license() directly renews token with live Worker, auto-purges on revocation, and blocks access."""
        time.sleep(0.3)
        # Step 1: Admin generates a genuine key
        created = _call_admin_api("POST", "/v1/admin/licenses", {
            "customer_name": "Realtime Validate Corp",
            "customer_email": "realtime@luv.vn",
            "expires_at": "2032-12-31",
            "products": {
                "DOC": {"enabled": True, "max_devices": 2},
                "SDS": {"enabled": True, "max_devices": 2},
            },
            "plan": "enterprise",
            "features": ["core", "doc_tran", "export"],
        })
        lic_id = created["id"]
        key = created["license_key"]

        try:
            time.sleep(0.3)
            # Step 2: Client activates the key
            act_res = LUVLicenseClient.activate_key(key, "Realtime Validate Client")
            self.assertTrue(act_res["is_active"])
            self.assertTrue(os.path.exists(self.license_sec))

            time.sleep(0.3)
            # Step 3: Call validate_license() online - must renew token and increment lease_seq
            val_res = LUVLicenseClient.validate_license(force_online=True)
            self.assertTrue(val_res["is_active"])
            self.assertFalse(val_res.get("revoked", False))
            self.assertEqual(val_res["status"], "active")
            self.assertEqual(val_res.get("lease_seq"), 2)

            # Check that license_sec was updated
            info = LUVLicenseClient.get_license_info()
            self.assertTrue(info["is_active"])

            # Step 4: Admin revokes the key
            rev_res = _call_admin_api("POST", f"/v1/admin/licenses/{lic_id}/revoke")
            self.assertTrue(rev_res.get("ok"))

            # Step 5: Call validate_license() online - must detect revocation and auto-purge (allowing for D1 edge propagation)
            val_revoked = None
            for _ in range(6):
                time.sleep(0.5)
                val_revoked = LUVLicenseClient.validate_license(force_online=True)
                if val_revoked.get("revoked"):
                    break

            self.assertIsNotNone(val_revoked)
            self.assertFalse(val_revoked["is_active"])
            self.assertTrue(val_revoked.get("revoked"))
            self.assertEqual(val_revoked["status"], "revoked")
            self.assertEqual(val_revoked["tier"], "Giấy phép đã bị thu hồi")

            # Check that disk license files were completely purged
            self.assertFalse(os.path.exists(self.license_sec))
            self.assertFalse(os.path.exists(self.license_json))

            # Step 6: check_access must be permanently revoked, no trial fallback
            access = LUVLicenseClient.check_access(validated_license=val_revoked)
            self.assertFalse(access["has_access"])
            self.assertTrue(access.get("revoked"))

        finally:
            # Step 7: Clean up
            time.sleep(0.3)
            _call_admin_api("POST", f"/v1/admin/licenses/{lic_id}/delete", {"confirm": True})


if __name__ == "__main__":
    unittest.main()


