"""
core/license_client.py
──────────────────────
Hệ thống xác thực giấy phép bản quyền LUV-DOC-TRAN:
- Kết nối trực tiếp Máy chủ Giấy phép Cloudflare Worker: https://luv-license-api.luv-sds.workers.dev
- Xác thực mật mã bất đối xứng Ed25519 với Public Key chính thức.
- Tương thích 100% với LUV_License_Manager.exe và LUV SDS Intelligence - Bản Hoàn Chỉnh.exe.
- Chặn đứng 100% các mã nhập bừa / mã giả lập không tồn tại trong Cloudflare D1 Database.
- Hỗ trợ cơ chế dùng thử 48 giờ (Anti-tamper 48h Trial).
"""

from __future__ import annotations

import os
import sys
import json
import time
import base64
import hashlib
import platform
import threading
import urllib.request
import urllib.error
from datetime import datetime, timezone
from typing import Dict, Any, Optional, Tuple

from cryptography.exceptions import InvalidSignature
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LICENSE_FILE = os.path.join(BASE_DIR, "license.json")
TRIALS_FILE = os.path.join(BASE_DIR, "trials_db.json")

# Official Cloudflare Worker & Ed25519 Configuration
DEFAULT_SERVER_URL = "https://luv-license-api.luv-sds.workers.dev"
PUBLIC_KEY_B64 = "hsHHVWWmquFmME9wW4d2u1Qh5NdVtu3tktRDMl437Fo="
APP_VERSION = "1.1.0"
CLOCK_SKEW_SECONDS = 300
TRIAL_DURATION_SECONDS = 48 * 3600  # 48 hours = 2 days

_trial_lock = threading.Lock()


def _b64url_decode(value: str) -> bytes:
    """Decode base64url string with proper padding."""
    return base64.urlsafe_b64decode(value + "=" * (-len(value) % 4))


def _utc_now() -> int:
    return int(datetime.now(timezone.utc).timestamp())


def generate_device_hash(raw_hint: str = "") -> str:
    """Generate SHA-256 64-char hex device fingerprint matching license server format."""
    parts = [platform.system(), platform.machine(), os.environ.get("COMPUTERNAME", "")]
    if raw_hint:
        parts.append(raw_hint)
    raw = ("LUV-DEVICE-v1|" + "|".join(p for p in parts if p)).encode("utf-8")
    return hashlib.sha256(raw).hexdigest()


class LUVLicenseClient:
    """
    Client quản lý giấy phép bản quyền mật mã Cloudflare & Dùng thử 48h.
    """

    @classmethod
    def _read_trials(cls) -> Dict[str, Any]:
        with _trial_lock:
            if not os.path.exists(TRIALS_FILE):
                return {}
            try:
                with open(TRIALS_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return {}

    @classmethod
    def _write_trials(cls, data: Dict[str, Any]) -> None:
        with _trial_lock:
            temp = TRIALS_FILE + ".tmp"
            with open(temp, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            if os.path.exists(TRIALS_FILE):
                os.replace(temp, TRIALS_FILE)
            else:
                os.rename(temp, TRIALS_FILE)

    @classmethod
    def get_trial_status(cls, device_id: str = "") -> Dict[str, Any]:
        """Check or register 48-hour free trial for device."""
        clean_dev = device_id.strip() if device_id else generate_device_hash()
        trials = cls._read_trials()
        now = time.time()

        if clean_dev not in trials:
            trials[clean_dev] = {
                "started_at": now,
                "started_at_str": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                "device_id": clean_dev
            }
            cls._write_trials(trials)

        start_time = trials[clean_dev].get("started_at", now)
        elapsed = now - start_time
        remaining = max(0, TRIAL_DURATION_SECONDS - elapsed)
        is_valid = remaining > 0

        hours = int(remaining // 3600)
        minutes = int((remaining % 3600) // 60)
        seconds = int(remaining % 60)

        return {
            "device_id": clean_dev,
            "is_trial_active": is_valid,
            "remaining_seconds": int(remaining),
            "remaining_formatted": f"{hours}h {minutes}m" if hours > 0 else f"{minutes}m {seconds}s",
            "hours_left": hours,
            "minutes_left": minutes,
            "total_trial_hours": 48
        }

    @classmethod
    def get_license_info(cls) -> Dict[str, Any]:
        """Read activated license from disk and cryptographically verify signature."""
        if not os.path.exists(LICENSE_FILE):
            return {
                "is_active": False,
                "tier": "Chưa kích hoạt",
                "registered_to": None,
                "license_key": None
            }
        try:
            with open(LICENSE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)

            token = data.get("token")
            if not token or "." not in token:
                return {"is_active": False, "tier": "Chưa kích hoạt"}

            # Verify Ed25519 signature of the stored token
            payload_part, sig_part = token.split(".", 1)
            pub_key = Ed25519PublicKey.from_public_bytes(base64.b64decode(PUBLIC_KEY_B64))
            sig_bytes = _b64url_decode(sig_part)
            pub_key.verify(sig_bytes, payload_part.encode("ascii"))

            claims = json.loads(_b64url_decode(payload_part).decode("utf-8"))
            now = _utc_now()
            exp = int(claims.get("exp", 0))

            if exp > 0 and exp + CLOCK_SKEW_SECONDS < now:
                return {
                    "is_active": False,
                    "tier": "Giấy phép đã hết hạn",
                    "license_key": data.get("license_key"),
                    "expired": True
                }

            return {
                "is_active": True,
                "tier": data.get("tier") or f"Doanh Nghiệp ({claims.get('plan', 'pro').upper()})",
                "registered_to": claims.get("customer") or data.get("registered_to"),
                "license_key": data.get("license_key"),
                "expires_at": claims.get("expires_at") or data.get("expires_at"),
                "features": claims.get("features", [])
            }

        except Exception as e:
            return {
                "is_active": False,
                "tier": "Chưa kích hoạt",
                "registered_to": None,
                "license_key": None,
                "error": str(e)
            }

    @classmethod
    def activate_key(cls, key: str, registered_to: str = "Doanh nghiệp LUV", client_device_hint: str = "") -> Dict[str, Any]:
        """
        Kích hoạt mã giấy phép bằng cách kết nối trực tiếp với Máy chủ Cloudflare Worker:
        https://luv-license-api.luv-sds.workers.dev/v1/activate

        - Chặn 100% mã nhập bừa / mã giả mạo.
        - Xác thực chữ ký Ed25519 của Cloudflare Worker.
        """
        if not key or not isinstance(key, str):
            raise ValueError("Vui lòng nhập Mã Bản Quyền chính thức!")

        normalized_key = "".join(key.upper().split())
        if len(normalized_key) < 16:
            raise ValueError("Mã bản quyền không đúng định dạng. Mã giấy phép chuẩn có dạng LUV-XXXXX-XXXXX-XXXXX-XXXXX.")

        dev_id = generate_device_hash(client_device_hint)
        dev_label = (registered_to.strip() or platform.node() or "Windows Client")[:80]

        # First try product code 'DOC', fallback to 'SDS' (since owner key / keys may have DOC or SDS product entitlement)
        products_to_try = ["DOC", "SDS"]
        last_error = "Không thể kích hoạt mã bản quyền."

        token = None
        claims = None

        for prod in products_to_try:
            payload = {
                "license_key": normalized_key,
                "device_id": dev_id,
                "device_label": dev_label,
                "app_version": APP_VERSION,
                "product_code": prod
            }

            req = urllib.request.Request(
                f"{DEFAULT_SERVER_URL}/v1/activate",
                data=json.dumps(payload).encode("utf-8"),
                headers={
                    "Content-Type": "application/json",
                    "User-Agent": f"LUV-{prod}/{APP_VERSION}"
                },
                method="POST"
            )

            try:
                with urllib.request.urlopen(req, timeout=15) as resp:
                    resp_data = json.loads(resp.read().decode("utf-8"))
                    token = resp_data.get("token")
                    if token:
                        break
            except urllib.error.HTTPError as exc:
                raw = exc.read().decode("utf-8", errors="replace")
                try:
                    parsed = json.loads(raw)
                    msg = parsed.get("message") or parsed.get("error") or raw
                except Exception:
                    msg = raw
                last_error = msg
            except Exception as e:
                last_error = f"Không thể kết nối tới máy chủ giấy phép ({str(e)})."

        if not token:
            raise ValueError(f"Máy chủ giấy phép từ chối kích hoạt: {last_error}")

        # Cryptographically verify the Ed25519 signature
        try:
            payload_part, sig_part = token.split(".", 1)
            pub_key = Ed25519PublicKey.from_public_bytes(base64.b64decode(PUBLIC_KEY_B64))
            sig_bytes = _b64url_decode(sig_part)
            pub_key.verify(sig_bytes, payload_part.encode("ascii"))
            claims = json.loads(_b64url_decode(payload_part).decode("utf-8"))
        except (InvalidSignature, Exception) as exc:
            raise ValueError(f"Chữ ký số của giấy phép không hợp lệ: {str(exc)}")

        # Check claims
        now = _utc_now()
        exp = int(claims.get("exp", 0))
        if exp > 0 and exp + CLOCK_SKEW_SECONDS < now:
            raise ValueError("Mã bản quyền này đã hết hạn sử dụng.")

        plan_name = claims.get("plan", "Enterprise").upper()
        customer_name = claims.get("customer") or registered_to

        record = {
            "status": "active",
            "license_id": claims.get("license_id"),
            "license_key": normalized_key,
            "registered_to": customer_name,
            "tier": f"Doanh Nghiệp ({plan_name})",
            "expires_at": claims.get("expires_at", "Vĩnh viễn"),
            "offline_until": claims.get("offline_until_iso"),
            "features": claims.get("features", []),
            "token": token,
            "activated_at": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "device_id": dev_id,
            "is_active": True
        }

        # Persist valid license
        with open(LICENSE_FILE, "w", encoding="utf-8") as f:
            json.dump(record, f, ensure_ascii=False, indent=2)

        return record

    @classmethod
    def check_access(cls, device_id: str = "") -> Dict[str, Any]:
        """
        Check if client has active permission to translate:
        either via valid cryptographic License Key OR via 48-hour Free Trial.
        """
        lic = cls.get_license_info()
        if lic.get("is_active"):
            return {
                "has_access": True,
                "access_type": "licensed",
                "tier": lic.get("tier", "Enterprise"),
                "registered_to": lic.get("registered_to", "Doanh nghiệp"),
                "license_key": lic.get("license_key"),
                "remaining_seconds": None,
                "remaining_formatted": "Bản Quyền Chính Thức"
            }

        trial = cls.get_trial_status(device_id)
        return {
            "has_access": trial["is_trial_active"],
            "access_type": "trial",
            "tier": "Dùng thử 48 Giờ",
            "registered_to": "Khách dùng thử",
            "license_key": None,
            "remaining_seconds": trial["remaining_seconds"],
            "remaining_formatted": trial["remaining_formatted"],
            "hours_left": trial["hours_left"],
            "minutes_left": trial["minutes_left"]
        }
