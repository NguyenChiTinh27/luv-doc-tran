"""
tests/test_online_revocation.py
────────────────────────────────
Comprehensive test suite for Real-time Cloudflare Worker Ed25519 Online License Validation,
Token Renewal, Auto-Purge on Revocation, and Offline Grace Period Tolerance.

Covers Requirements:
- ORIGINAL_REQUEST.md §R3 (Online real-time validation, instant revocation purge, offline_until tolerance)
- PROJECT.md § Milestone 1 & Feature Inventory (#1, #2, #3, #4)
- TEST_INFRA.md § Feature Inventory (#1, #2, #3, #4) & Tier 4 Scenarios (#2, #3)

Test Categories:
1. `TestPurgeLicense`: Idempotent and clean purge of data/license.sec, data/license.sec.tmp, and license.json.
2. `TestOnlineValidationAndRenewal`: Real-time validation via POST /v1/validate, Ed25519 signature checks,
   token renewal, lease_seq increment, and header/payload contract verification.
3. `TestOnlineRevocationHandling`: HTTP 403/401/404 handling, license_denied error, disk auto-purge,
   revocation status propagation to check_access() and /api/translate (HTTP 403).
4. `TestNetworkOutageAndOfflineGrace`: URLError handling, within offline_until access granted,
   beyond offline_until access denied.
5. `TestAdversarialAndEdgeCases`: Corrupted responses, clock skew boundaries, tampered signatures.
"""

from __future__ import annotations

import base64
import io
import json
import os
import sys
import time
import unittest
import urllib.error
import urllib.request
from datetime import datetime, timezone, timedelta
from typing import Any, Dict, Optional
from unittest.mock import MagicMock, patch

from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey, Ed25519PublicKey
from fastapi.testclient import TestClient

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from core.license_client import (
    APP_VERSION,
    CLOCK_SKEW_SECONDS,
    DEFAULT_SERVER_URL,
    PUBLIC_KEY_B64,
    LUVLicenseClient,
    _b64url_decode,
    generate_device_hash,
)
from server import app

OFFICIAL_PUBLIC_KEY_B64 = "hsHHVWWmquFmME9wW4d2u1Qh5NdVtu3tktRDMl437Fo="
OFFICIAL_SAMPLE_TOKEN = "eyJpc3MiOiJsdXYtbGljZW5zZS1zZXJ2ZXIiLCJhdWQiOiJsdXYtc2RzLWludGVsbGlnZW5jZSIsInN1YiI6IjkwYTFkMmM0LTE2MGYtNGM4Ny1iN2YzLTEzN2E0MmViMTIzMCIsImxpY2Vuc2VfaWQiOiI5MGExZDJjNC0xNjBmLTRjODctYjdmMy0xMzdhNDJlYjEyMzAiLCJwcm9kdWN0X2NvZGUiOiJET0MiLCJjdXN0b21lciI6IkRvYW5oIE5naGllcCBMVVYgRE9DLVRSQU4iLCJwbGFuIjoiZW50ZXJwcmlzZSIsImZlYXR1cmVzIjpbImNvcmUiLCJpbXBvcnQiLCJleHBvcnQiLCJlZGl0IiwiZG9jX3RyYW4iXSwiZGV2aWNlX2lkIjoiMTExMTIyMjIzMzMzNDQ0NDU1NTU2NjY2Nzc3Nzg4ODg5OTk5MDAwMGFhYWFiYmJiY2NjY2RkZGRlZWVlZmZmZiIsImlhdCI6MTc5MjMzNTA2MiwicmVmcmVzaF9hZnRlciI6MTc5MjQyMTQ2Miwib2ZmbGluZV91bnRpbCI6MTc5MzU0NDY2Miwib2ZmbGluZV91bnRpbF9pc28iOiIyMDI2LTExLTAxVDE0OjUxOjAyLjkzN1oiLCJleHAiOjIxMTQzODA3OTksImV4cGlyZXNfYXQiOiIyMDM2LTEyLTMxVDIzOjU5OjU5LjAwMFoiLCJqdGkiOiJiZjVlMzQ3MC05OWJlLTQ1NTgtODNjOC0zZGUxMGExMGFjMTQiLCJsZWFzZV9zZXEiOjF9.6aEJEn351Quc6Ct9Ptck4aa9I-EQpRvdKqw8gPf_o-wb4z_zHiwS9Q4WOj3ohkwoUbo5TxAokJkwmxebcfRmBw"


def _b64url_encode(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")


def create_test_signed_token(claims: Dict[str, Any], private_key: Ed25519PrivateKey) -> str:
    raw_payload = json.dumps(claims, separators=(",", ":"), ensure_ascii=False).encode("utf-8")
    b64_payload = _b64url_encode(raw_payload)
    sig_bytes = private_key.sign(b64_payload.encode("ascii"))
    b64_sig = _b64url_encode(sig_bytes)
    return f"{b64_payload}.{b64_sig}"


class BaseOnlineRevocationTestCase(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.test_private_key = Ed25519PrivateKey.generate()
        cls.test_public_key = cls.test_private_key.public_key()
        cls.test_pub_b64 = base64.b64encode(cls.test_public_key.public_bytes_raw()).decode("ascii")

    def setUp(self):
        super().setUp()
        self.data_dir = os.path.join(PROJECT_ROOT, "data")
        self.license_sec = os.path.join(self.data_dir, "license.sec")
        self.license_tmp = os.path.join(self.data_dir, "license.sec.tmp")
        self.legacy_json = os.path.join(PROJECT_ROOT, "license.json")
        self.trials_db = os.path.join(PROJECT_ROOT, "trials_db.json")

        os.makedirs(self.data_dir, exist_ok=True)

        self._backup_sec = None
        if os.path.exists(self.license_sec):
            with open(self.license_sec, "rb") as f:
                self._backup_sec = f.read()

        self._backup_json = None
        if os.path.exists(self.legacy_json):
            with open(self.legacy_json, "rb") as f:
                self._backup_json = f.read()

        self._backup_trials = None
        if os.path.exists(self.trials_db):
            with open(self.trials_db, "rb") as f:
                self._backup_trials = f.read()

        self._remove_license_files()
        self.client = TestClient(app)

    def tearDown(self):
        self._remove_license_files()
        if self._backup_sec is not None:
            with open(self.license_sec, "wb") as f:
                f.write(self._backup_sec)
        if self._backup_json is not None:
            with open(self.legacy_json, "wb") as f:
                f.write(self._backup_json)
        if self._backup_trials is not None:
            with open(self.trials_db, "wb") as f:
                f.write(self._backup_trials)
        super().tearDown()

    def _remove_license_files(self):
        LUVLicenseClient._is_revoked = False
        LUVLicenseClient._cached_claims = None
        for path in (self.license_sec, self.license_tmp, self.legacy_json, self.trials_db):
            if os.path.exists(path):
                try:
                    os.remove(path)
                except Exception:
                    pass

    def _make_dummy_active_record(
        self,
        token: Optional[str] = None,
        offline_until_ts: Optional[int] = None,
        exp_ts: Optional[int] = None
    ) -> Dict[str, Any]:
        """Create a valid license record saved to data/license.sec."""
        now = int(time.time())
        exp = exp_ts if exp_ts is not None else now + 31536000
        off_ts = offline_until_ts if offline_until_ts is not None else now + 1209600

        if token is None:
            claims = {
                "iss": "luv-license-server",
                "aud": "luv-sds-intelligence",
                "sub": "test-uuid-sub",
                "license_id": "test-license-id-1234",
                "product_code": "DOC",
                "customer": "Doanh Nghiep Test E2E",
                "plan": "enterprise",
                "features": ["core", "import", "export", "edit", "doc_tran"],
                "device_id": generate_device_hash(),
                "iat": now - 3600,
                "offline_until": off_ts,
                "offline_until_iso": datetime.fromtimestamp(off_ts, tz=timezone.utc).isoformat(),
                "exp": exp,
                "expires_at": datetime.fromtimestamp(exp, tz=timezone.utc).isoformat(),
                "lease_seq": 1
            }
            token = create_test_signed_token(claims, self.test_private_key)

        record = {
            "status": "active",
            "license_id": "test-license-id-1234",
            "license_key": "LUV-TESTK-EY123-45678-99999",
            "registered_to": "Doanh Nghiep Test E2E",
            "customer": "Doanh Nghiep Test E2E",
            "tier": "Doanh Nghiệp (ENTERPRISE)",
            "expires_at": datetime.fromtimestamp(exp, tz=timezone.utc).isoformat(),
            "offline_until": off_ts,
            "offline_until_iso": datetime.fromtimestamp(off_ts, tz=timezone.utc).isoformat(),
            "features": ["core", "import", "export", "edit", "doc_tran"],
            "token": token,
            "activated_at": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "device_id": generate_device_hash(),
            "is_active": True
        }
        LUVLicenseClient.save_secure_license(record)
        return record


class TestPurgeLicense(BaseOnlineRevocationTestCase):
    def test_01_purge_license_deletes_all_local_license_files(self):
        """Assert purge_license() deletes data/license.sec, temporary file, and legacy license.json."""
        self._make_dummy_active_record()
        self.assertTrue(os.path.exists(self.license_sec), "license.sec must exist prior to purge")
        with open(self.license_tmp, "wb") as f:
            f.write(b"temporary data")
        with open(self.legacy_json, "w", encoding="utf-8") as f:
            f.write('{"token": "legacy"}')
        self.assertTrue(os.path.exists(self.license_tmp))
        self.assertTrue(os.path.exists(self.legacy_json))
        self.assertTrue(hasattr(LUVLicenseClient, "purge_license"), "LUVLicenseClient must implement purge_license()")
        LUVLicenseClient.purge_license()
        self.assertFalse(os.path.exists(self.license_sec), "data/license.sec must be purged")
        self.assertFalse(os.path.exists(self.license_tmp), "data/license.sec.tmp must be purged")
        self.assertFalse(os.path.exists(self.legacy_json), "license.json must be purged")

    def test_02_purge_license_is_idempotent(self):
        """Assert calling purge_license() multiple times when no files exist does not raise exception."""
        self._remove_license_files()
        try:
            LUVLicenseClient.purge_license()
            LUVLicenseClient.purge_license()
        except Exception as exc:
            self.fail(f"purge_license() must be idempotent and not raise error: {exc}")


class TestOnlineValidationAndRenewal(BaseOnlineRevocationTestCase):
    def test_03_validate_license_when_no_stored_license(self):
        """Assert validate_license() returns unactivated when no license is present without network calls."""
        self._remove_license_files()
        self.assertTrue(hasattr(LUVLicenseClient, "validate_license"), "LUVLicenseClient must implement validate_license()")
        with patch("urllib.request.urlopen") as mock_urlopen:
            res = LUVLicenseClient.validate_license()
            mock_urlopen.assert_not_called()
        self.assertIsInstance(res, dict)
        self.assertFalse(res.get("is_active", False))
        self.assertIn("Chưa kích hoạt", res.get("tier", ""))

    def test_04_validate_license_successful_renewal_and_token_update(self):
        """Verify successful online validation updates token, offline_until, and lease_seq."""
        with patch("core.license_client.PUBLIC_KEY_B64", self.test_pub_b64):
            self._make_dummy_active_record()
            now = int(time.time())
            renewed_claims = {
                "iss": "luv-license-server",
                "aud": "luv-sds-intelligence",
                "sub": "test-uuid-sub",
                "license_id": "test-license-id-1234",
                "product_code": "DOC",
                "customer": "Doanh Nghiep Renewed",
                "plan": "enterprise",
                "features": ["core", "import", "export", "edit", "doc_tran"],
                "device_id": generate_device_hash(),
                "iat": now,
                "offline_until": now + 604800,
                "offline_until_iso": datetime.fromtimestamp(now + 604800, tz=timezone.utc).isoformat(),
                "exp": now + 31536000,
                "expires_at": datetime.fromtimestamp(now + 31536000, tz=timezone.utc).isoformat(),
                "lease_seq": 2
            }
            renewed_token = create_test_signed_token(renewed_claims, self.test_private_key)
            resp_payload = {"token": renewed_token, "valid": True}

            mock_resp = MagicMock()
            mock_resp.status = 200
            mock_resp.read.return_value = json.dumps(resp_payload).encode("utf-8")
            mock_resp.__enter__.return_value = mock_resp
            mock_resp.__exit__.return_value = None

            captured_requests = []

            def fake_urlopen(req, *args, **kwargs):
                captured_requests.append(req)
                return mock_resp

            with patch("urllib.request.urlopen", side_effect=fake_urlopen):
                result = LUVLicenseClient.validate_license(force_online=True)

            self.assertTrue(result.get("is_active"))
            self.assertEqual(result.get("status"), "active")
            self.assertFalse(result.get("revoked", False))
            self.assertEqual(len(captured_requests), 1)
            req = captured_requests[0]
            self.assertEqual(req.full_url, f"{DEFAULT_SERVER_URL}/v1/validate")
            self.assertEqual(req.get_header("User-agent"), f"LUV-DOC/{APP_VERSION}")
            self.assertEqual(req.get_header("Content-type"), "application/json")
            sent_payload = json.loads(req.data.decode("utf-8"))
            self.assertIn("token", sent_payload)
            self.assertEqual(sent_payload.get("app_version"), APP_VERSION)
            self.assertIn("device_id", sent_payload)

            saved = LUVLicenseClient.load_secure_license()
            self.assertIsNotNone(saved)
            self.assertEqual(saved.get("token"), renewed_token)

    def test_05_validate_license_rejects_tampered_signature_in_renewal(self):
        """Assert validate_license() rejects a renewed token with invalid Ed25519 signature."""
        with patch("core.license_client.PUBLIC_KEY_B64", self.test_pub_b64):
            self._make_dummy_active_record()
            now = int(time.time())
            untrusted_priv_key = Ed25519PrivateKey.generate()
            fake_claims = {
                "iss": "luv-license-server",
                "customer": "Attacker",
                "plan": "enterprise",
                "exp": now + 86400
            }
            tampered_token = create_test_signed_token(fake_claims, untrusted_priv_key)

            mock_resp = MagicMock()
            mock_resp.status = 200
            mock_resp.code = 200
            mock_resp.read.return_value = json.dumps({"token": tampered_token}).encode("utf-8")
            mock_resp.__enter__.return_value = mock_resp

            with patch("urllib.request.urlopen", return_value=mock_resp):
                result = LUVLicenseClient.validate_license()

            self.assertNotEqual(result.get("token"), tampered_token)
            saved = LUVLicenseClient.load_secure_license()
            self.assertIsNotNone(saved)
            self.assertNotEqual(saved.get("token"), tampered_token)


class TestOnlineRevocationHandling(BaseOnlineRevocationTestCase):
    def test_06_revocation_on_http_403_license_denied_purges_disk_and_sets_revoked(self):
        """When Worker returns HTTP 403 / license_denied, verify data/license.sec is deleted and status is revoked."""
        with patch("core.license_client.PUBLIC_KEY_B64", self.test_pub_b64):
            self._make_dummy_active_record()
            self.assertTrue(os.path.exists(self.license_sec))

            err_body = json.dumps({"error": "license_denied", "message": "Key revoked"}).encode("utf-8")
            http_err = urllib.error.HTTPError(
                url=f"{DEFAULT_SERVER_URL}/v1/validate",
                code=403,
                msg="Forbidden",
                hdrs=MagicMock(),
                fp=io.BytesIO(err_body)
            )

            with patch("urllib.request.urlopen", side_effect=http_err):
                res = LUVLicenseClient.validate_license()

            self.assertFalse(res.get("is_active", True))
            self.assertTrue(res.get("revoked", False))
            self.assertEqual(res.get("status"), "revoked")
            self.assertFalse(os.path.exists(self.license_sec), "License file must be purged from disk")

    def test_07_revocation_on_http_401_and_404_purges_disk(self):
        """Verify HTTP 401 Unauthorized and HTTP 404 Not Found also trigger auto-purge."""
        with patch("core.license_client.PUBLIC_KEY_B64", self.test_pub_b64):
            for status_code in (401, 404):
                self._make_dummy_active_record()
                self.assertTrue(os.path.exists(self.license_sec))

                err_body = json.dumps({"error": "device_not_registered"}).encode("utf-8")
                http_err = urllib.error.HTTPError(
                    url=f"{DEFAULT_SERVER_URL}/v1/validate",
                    code=status_code,
                    msg="Error",
                    hdrs=MagicMock(),
                    fp=io.BytesIO(err_body)
                )

                with patch("urllib.request.urlopen", side_effect=http_err):
                    res = LUVLicenseClient.validate_license()

                self.assertFalse(res.get("is_active", True))
                self.assertFalse(os.path.exists(self.license_sec), f"File must be purged on HTTP {status_code}")

    def test_08_check_access_blocked_after_revocation(self):
        """Assert check_access() returns has_access=False and revoked=True when license is revoked."""
        with patch("core.license_client.PUBLIC_KEY_B64", self.test_pub_b64):
            self._make_dummy_active_record()

            err_body = json.dumps({"error": "license_denied"}).encode("utf-8")
            http_err = urllib.error.HTTPError(
                url=f"{DEFAULT_SERVER_URL}/v1/validate",
                code=403,
                msg="Forbidden",
                hdrs=MagicMock(),
                fp=io.BytesIO(err_body)
            )

            with patch("urllib.request.urlopen", side_effect=http_err):
                access = LUVLicenseClient.check_access()

            self.assertFalse(access.get("has_access", True))
            self.assertTrue(access.get("revoked", False))
            self.assertIn("thu hồi", access.get("message", "").lower())
            self.assertFalse(os.path.exists(self.license_sec))

    def test_09_api_license_status_reports_revocation(self):
        """GET /api/license/status calls validate_license and returns revoked state."""
        with patch("core.license_client.PUBLIC_KEY_B64", self.test_pub_b64):
            self._make_dummy_active_record()

            err_body = json.dumps({"error": "license_denied"}).encode("utf-8")
            http_err = urllib.error.HTTPError(
                url=f"{DEFAULT_SERVER_URL}/v1/validate",
                code=403,
                msg="Forbidden",
                hdrs=MagicMock(),
                fp=io.BytesIO(err_body)
            )

            with patch("urllib.request.urlopen", side_effect=http_err):
                resp = self.client.get("/api/license/status")

            self.assertEqual(resp.status_code, 200)
            data = resp.json()
            self.assertFalse(data.get("is_active", True))
            self.assertTrue(data.get("revoked", False))
            self.assertIn("thu hồi", data.get("message", "").lower())

    def test_10_api_translate_locks_translation_with_403_when_revoked(self):
        """POST /api/translate rejects with HTTP 403 Forbidden when license is revoked."""
        with patch("core.license_client.PUBLIC_KEY_B64", self.test_pub_b64):
            self._make_dummy_active_record()

            err_body = json.dumps({"error": "license_denied"}).encode("utf-8")
            http_err = urllib.error.HTTPError(
                url=f"{DEFAULT_SERVER_URL}/v1/validate",
                code=403,
                msg="Forbidden",
                hdrs=MagicMock(),
                fp=io.BytesIO(err_body)
            )

            with patch("urllib.request.urlopen", side_effect=http_err):
                LUVLicenseClient.validate_license()

            fake_file = io.BytesIO(b"Fake content")
            resp = self.client.post(
                "/api/translate",
                files={"file": ("test.docx", fake_file, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
                data={
                    "source_lang": "auto",
                    "target_lang": "vi",
                    "output_mode": "bilingual",
                    "provider": "google",
                    "engine_mode": "auto",
                    "device_id": generate_device_hash()
                }
            )

            self.assertEqual(resp.status_code, 403, "/api/translate must return HTTP 403 Forbidden when revoked")
            res_json = resp.json()
            self.assertFalse(res_json.get("success", True))
            self.assertTrue(res_json.get("revoked", False))
            self.assertIn("thu hồi", res_json.get("message", ""))


class TestNetworkOutageAndOfflineGrace(BaseOnlineRevocationTestCase):
    def test_11_network_outage_within_offline_grace_period_grants_access(self):
        """When network outage occurs and offline_until is in future, access remains active and file not deleted."""
        with patch("core.license_client.PUBLIC_KEY_B64", self.test_pub_b64):
            now = int(time.time())
            future_offline_ts = now + 604800
            self._make_dummy_active_record(offline_until_ts=future_offline_ts)

            network_err = urllib.error.URLError("Network unreachable")
            with patch("urllib.request.urlopen", side_effect=network_err):
                result = LUVLicenseClient.validate_license(force_online=True)

            self.assertTrue(result.get("is_active"))
            self.assertFalse(result.get("revoked", False))
            self.assertTrue(result.get("offline_mode", False) or result.get("status") == "offline_active")
            self.assertTrue(os.path.exists(self.license_sec), "License file must NOT be purged during network outage")

    def test_12_network_outage_beyond_offline_grace_period_denies_access(self):
        """When network outage occurs and offline_until has expired, access is blocked."""
        with patch("core.license_client.PUBLIC_KEY_B64", self.test_pub_b64):
            now = int(time.time())
            past_offline_ts = now - 3600
            self._make_dummy_active_record(offline_until_ts=past_offline_ts)

            network_err = urllib.error.URLError("Network unreachable")
            with patch("urllib.request.urlopen", side_effect=network_err):
                result = LUVLicenseClient.validate_license(force_online=True)

            self.assertFalse(result.get("is_active", True), "Access must be denied when offline_until expired")
            self.assertTrue(result.get("offline_expired", False) or "hết hạn" in result.get("tier", "").lower())

    def test_13_force_online_flag_respects_valid_offline_cache(self):
        """When force_online=False, validate_license uses verified local cache if within offline grace period."""
        with patch("core.license_client.PUBLIC_KEY_B64", self.test_pub_b64):
            now = int(time.time())
            future_offline_ts = now + 604800
            self._make_dummy_active_record(offline_until_ts=future_offline_ts)

            with patch("urllib.request.urlopen") as mock_urlopen:
                res = LUVLicenseClient.validate_license(force_online=False)
                mock_urlopen.assert_not_called()

            self.assertTrue(res.get("is_active"))


class TestAdversarialAndCryptoIntegrity(BaseOnlineRevocationTestCase):
    def test_14_official_sample_token_oracle_verification(self):
        """Verify the official sample token verifies cleanly against official PUBLIC_KEY_B64."""
        self.assertIn(".", OFFICIAL_SAMPLE_TOKEN)
        payload_part, sig_part = OFFICIAL_SAMPLE_TOKEN.split(".", 1)
        pub_key = Ed25519PublicKey.from_public_bytes(base64.b64decode(OFFICIAL_PUBLIC_KEY_B64))
        sig_bytes = _b64url_decode(sig_part)
        pub_key.verify(sig_bytes, payload_part.encode("ascii"))
        claims = json.loads(_b64url_decode(payload_part).decode("utf-8"))
        self.assertEqual(claims.get("product_code"), "DOC")
        self.assertEqual(claims.get("plan"), "enterprise")
        self.assertEqual(claims.get("lease_seq"), 1)

    def test_15_server_returns_malformed_html_error_on_revocation(self):
        """When Cloudflare Worker returns HTML error page (e.g. 403 Cloudflare WAF block), handles gracefully."""
        with patch("core.license_client.PUBLIC_KEY_B64", self.test_pub_b64):
            self._make_dummy_active_record()
            html_error = b"<html><head><title>403 Forbidden</title></head><body>Cloudflare Access Denied</body></html>"
            http_err = urllib.error.HTTPError(
                url=f"{DEFAULT_SERVER_URL}/v1/validate",
                code=403,
                msg="Forbidden",
                hdrs=MagicMock(),
                fp=io.BytesIO(html_error)
            )

            with patch("urllib.request.urlopen", side_effect=http_err):
                res = LUVLicenseClient.validate_license()

            self.assertFalse(res.get("is_active", True))
            self.assertTrue(res.get("revoked", False))
            self.assertFalse(os.path.exists(self.license_sec))


if __name__ == "__main__":
    unittest.main()
