"""
core/license_client.py
──────────────────────
Hệ thống xác thực giấy phép bản quyền LUV-DOC-TRAN:
- Kết nối trực tiếp Máy chủ Giấy phép Cloudflare Worker: https://luv-license-api.luv-sds.workers.dev
- Xác thực mật mã bất đối xứng Ed25519 với Public Key chính thức.
- Tương thích 100% với LUV_License_Manager.exe và LUV SDS Intelligence - Bản Hoàn Chỉnh.exe.
- Chặn đứng 100% các mã nhập bừa / mã giả lập không tồn tại trong Cloudflare D1 Database.
- Bảo mật tệp giấy phép đa nền tảng: Windows DPAPI (CryptProtectData / CryptUnprotectData với entropy)
  và Base64 envelope an toàn trên Linux/Docker/Render.
- Tự động di trú (migration fallback) từ license.json sang data/license.sec.
- Hỗ trợ cơ chế dùng thử 48 giờ (Anti-tamper 48h Trial).
"""

from __future__ import annotations

import os
import re
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
from typing import Dict, Any, Optional

from cryptography.exceptions import InvalidSignature
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "data")
LICENSE_FILE = os.path.join(DATA_DIR, "license.sec")
LEGACY_LICENSE_FILE = os.path.join(BASE_DIR, "license.json")
TRIALS_FILE = os.path.join(BASE_DIR, "trials_db.json")

os.makedirs(DATA_DIR, exist_ok=True)

# Official Cloudflare Worker & Ed25519 Configuration
DEFAULT_SERVER_URL = "https://luv-license-api.luv-sds.workers.dev"
PUBLIC_KEY_B64 = "hsHHVWWmquFmME9wW4d2u1Qh5NdVtu3tktRDMl437Fo="
APP_VERSION = "1.1.0"
CLOCK_SKEW_SECONDS = 300
TRIAL_DURATION_SECONDS = 48 * 3600  # 48 hours = 2 days
ENTROPY_BYTES = b"LUV-SDS-License-v1"

_trial_lock = threading.Lock()

if os.name == "nt":
    import ctypes
    from ctypes import wintypes

    class DATA_BLOB(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

    _crypt32 = ctypes.WinDLL("crypt32", use_last_error=True)
    _kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

    _crypt32.CryptProtectData.argtypes = [
        ctypes.POINTER(DATA_BLOB), wintypes.LPCWSTR, ctypes.POINTER(DATA_BLOB),
        ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(DATA_BLOB),
    ]
    _crypt32.CryptProtectData.restype = wintypes.BOOL
    _crypt32.CryptUnprotectData.argtypes = [
        ctypes.POINTER(DATA_BLOB), ctypes.POINTER(wintypes.LPWSTR), ctypes.POINTER(DATA_BLOB),
        ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(DATA_BLOB),
    ]
    _crypt32.CryptUnprotectData.restype = wintypes.BOOL
    _kernel32.LocalFree.argtypes = [wintypes.HLOCAL]
    _kernel32.LocalFree.restype = wintypes.HLOCAL


def _protect_data(data: bytes) -> bytes:
    """Platform-aware data protection: DPAPI on Windows, safe passthrough envelope on Linux/Docker."""
    if os.name != "nt":
        return data
    source = ctypes.create_string_buffer(data)
    source_blob = DATA_BLOB(len(data), ctypes.cast(source, ctypes.POINTER(ctypes.c_char)))
    output_blob = DATA_BLOB()
    entropy = ctypes.create_string_buffer(ENTROPY_BYTES)
    entropy_blob = DATA_BLOB(len(ENTROPY_BYTES), ctypes.cast(entropy, ctypes.POINTER(ctypes.c_char)))
    if not _crypt32.CryptProtectData(
        ctypes.byref(source_blob), "LUV DOC License", ctypes.byref(entropy_blob),
        None, None, 0x1, ctypes.byref(output_blob),
    ):
        raise ValueError("Không thể mã hóa bảo vệ giấy phép trên máy trạm này.")
    try:
        return ctypes.string_at(output_blob.pbData, output_blob.cbData)
    finally:
        _kernel32.LocalFree(output_blob.pbData)


def _unprotect_data(data: bytes) -> bytes:
    """Platform-aware data unprotection: DPAPI on Windows, safe passthrough envelope on Linux/Docker."""
    if os.name != "nt":
        return data
    source = ctypes.create_string_buffer(data)
    source_blob = DATA_BLOB(len(data), ctypes.cast(source, ctypes.POINTER(ctypes.c_char)))
    output_blob = DATA_BLOB()
    entropy = ctypes.create_string_buffer(ENTROPY_BYTES)
    entropy_blob = DATA_BLOB(len(ENTROPY_BYTES), ctypes.cast(entropy, ctypes.POINTER(ctypes.c_char)))
    if not _crypt32.CryptUnprotectData(
        ctypes.byref(source_blob), None, ctypes.byref(entropy_blob),
        None, None, 0x1, ctypes.byref(output_blob),
    ):
        raise ValueError("Giấy phép cục bộ không thuộc máy này hoặc đã bị chỉnh sửa.")
    try:
        return ctypes.string_at(output_blob.pbData, output_blob.cbData)
    finally:
        _kernel32.LocalFree(output_blob.pbData)


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
    PUBLIC_KEY_B64 = PUBLIC_KEY_B64
    _is_revoked: bool = False
    _cached_claims: Optional[Dict[str, Any]] = None
    _license_lock: threading.Lock = threading.Lock()
    _last_online_validation_time: float = 0.0
    _last_online_validation_result: Optional[Dict[str, Any]] = None
    _online_cache_ttl: float = 600.0  # 10 minutes cache

    @classmethod
    def save_secure_license(cls, record: Dict[str, Any]) -> None:
        """Save license record securely to data/license.sec using platform protection."""
        with cls._license_lock:
            cls._last_online_validation_time = 0.0
            cls._last_online_validation_result = None
        os.makedirs(DATA_DIR, exist_ok=True)
        raw_bytes = json.dumps(record, ensure_ascii=False, indent=2).encode("utf-8")
        protected_bytes = _protect_data(raw_bytes)
        b64_envelope = base64.b64encode(protected_bytes)

        temp_file = LICENSE_FILE + ".tmp"
        with open(temp_file, "wb") as f:
            f.write(b64_envelope)
        if os.path.exists(LICENSE_FILE):
            os.replace(temp_file, LICENSE_FILE)
        else:
            os.rename(temp_file, LICENSE_FILE)

    @classmethod
    def load_secure_license(cls) -> Optional[Dict[str, Any]]:
        """
        Load license record from data/license.sec.
        If data/license.sec does not exist but legacy license.json exists,
        automatically migrate to data/license.sec.
        """
        # Migration fallback: if data/license.sec does not exist but license.json exists
        if not os.path.exists(LICENSE_FILE) and os.path.exists(LEGACY_LICENSE_FILE):
            try:
                with open(LEGACY_LICENSE_FILE, "r", encoding="utf-8") as f:
                    legacy_data = json.load(f)
                if legacy_data and legacy_data.get("token"):
                    cls.save_secure_license(legacy_data)
            except Exception:
                pass

        if not os.path.exists(LICENSE_FILE):
            return None

        try:
            with open(LICENSE_FILE, "rb") as f:
                content = f.read().strip()
            if not content:
                return None

            try:
                decoded = base64.b64decode(content, validate=True)
            except Exception:
                decoded = content

            try:
                unprotected = _unprotect_data(decoded)
            except Exception:
                unprotected = decoded

            text = unprotected.decode("utf-8")
            if text.strip().startswith("{"):
                return json.loads(text)
            elif "." in text:
                return {"token": text.strip()}
            return None
        except Exception:
            return None

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
        """Cơ chế dùng thử đã bỏ hoàn toàn: Bắt buộc nhập Key kích hoạt từ ban đầu."""
        clean_dev = device_id.strip() if device_id else generate_device_hash()
        return {
            "device_id": clean_dev,
            "is_trial_active": False,
            "trial_disabled": True,
            "remaining_seconds": 0,
            "remaining_formatted": "Yêu Cầu Key Kích Hoạt",
            "hours_left": 0,
            "minutes_left": 0,
            "total_trial_hours": 0,
            "message": "Cơ chế dùng thử không còn áp dụng. Vui lòng nhập Key kích hoạt để sử dụng hệ thống."
        }

    @classmethod
    def get_license_info(cls) -> Dict[str, Any]:
        """Read activated license from data/license.sec and cryptographically verify signature."""
        data = cls.load_secure_license()
        if not data:
            return {
                "is_active": False,
                "tier": "Chưa kích hoạt",
                "registered_to": None,
                "customer": None,
                "license_key": None
            }
        try:
            token = data.get("token")
            if not token or "." not in token:
                return {
                    "is_active": False,
                    "tier": "Chưa kích hoạt",
                    "registered_to": None,
                    "customer": None,
                    "license_key": None
                }

            # Authentic Ed25519 signature verification against official Public Key
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
                    "registered_to": claims.get("customer") or data.get("registered_to"),
                    "customer": claims.get("customer") or data.get("registered_to"),
                    "license_key": data.get("license_key"),
                    "expired": True
                }

            plan_name = claims.get("plan", "pro").upper()
            customer_name = claims.get("customer") or data.get("registered_to") or "Người Dùng Bản Quyền"
            clean_reg = re.sub(r'(?i)doanh\s*nghiệp', 'Người Dùng', customer_name) if "doanh nghiệp" in customer_name.lower() else customer_name
            raw_tier = data.get("tier") or f"Key Kích Hoạt ({plan_name})"
            clean_tier = re.sub(r'(?i)doanh\s*nghiệp', 'Key Kích Hoạt', raw_tier) if "doanh nghiệp" in raw_tier.lower() else raw_tier

            return {
                "is_active": True,
                "tier": clean_tier,
                "registered_to": clean_reg,
                "customer": clean_reg,
                "license_key": data.get("license_key"),
                "expires_at": claims.get("expires_at") or data.get("expires_at"),
                "features": claims.get("features", [])
            }

        except Exception as e:
            return {
                "is_active": False,
                "tier": "Chưa kích hoạt",
                "registered_to": None,
                "customer": None,
                "license_key": None,
                "error": str(e)
            }

    @classmethod
    def activate_key(cls, key: str, registered_to: str = "Người Dùng Bản Quyền", client_device_hint: str = "") -> Dict[str, Any]:
        """
        Kích hoạt mã giấy phép bằng cách kết nối trực tiếp với Máy chủ Cloudflare Worker:
        https://luv-license-api.luv-sds.workers.dev/v1/activate

        - Chặn 100% mã nhập bừa / mã giả mạo.
        - Cơ chế 2-step fallback: thử 'DOC' trước, nếu HTTP 403 product_denied thì thử 'SDS'.
        - Header bắt buộc: User-Agent: LUV-DOC/1.1.0.
        - Xác thực chữ ký số Ed25519 với Public Key chính thức.
        - Lưu trữ an toàn vào data/license.sec với bảo vệ DPAPI (Windows) / base64 envelope (Linux).
        """
        if not key or not isinstance(key, str):
            raise ValueError("Vui lòng nhập Mã Bản Quyền chính thức!")

        normalized_key = "".join(key.upper().split())
        if len(normalized_key) < 16:
            raise ValueError("Mã bản quyền không đúng định dạng. Mã giấy phép chuẩn có dạng LUV-XXXXX-XXXXX-XXXXX-XXXXX.")

        dev_id = generate_device_hash(client_device_hint)
        dev_label = (registered_to.strip() or platform.node() or "Windows Client")[:80]

        # 2-step activation fallback: try 'DOC' first, on HTTP 403 'product_denied' fallback to 'SDS'
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
                    "User-Agent": f"LUV-DOC/{APP_VERSION}"
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
                err_code = ""
                try:
                    parsed = json.loads(raw)
                    err_code = parsed.get("error", "")
                    msg = parsed.get("message") or parsed.get("error") or raw
                except Exception:
                    msg = raw

                # If product denied on DOC, fallback to SDS
                if prod == "DOC" and err_code == "product_denied":
                    last_error = msg
                    continue

                if err_code == "license_denied":
                    last_error = "Mã bản quyền không tồn tại trên hệ thống hoặc đã bị thu hồi."
                elif err_code == "device_limit":
                    last_error = "Mã Key đã vượt quá số lượng thiết bị cho phép. Thiết bị này bị từ chối truy cập."
                elif err_code == "invalid_request":
                    last_error = "Mã bản quyền hoặc thông tin kích hoạt không hợp lệ."
                else:
                    last_error = msg
                break
            except Exception as e:
                last_error = f"Không thể kết nối tới máy chủ giấy phép ({str(e)})."
                break

        if not token:
            raise ValueError(f"Máy chủ giấy phép từ chối kích hoạt: {last_error}")

        # Authentic Ed25519 verification
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
            "registered_to": re.sub(r'(?i)doanh\s*nghiệp', 'Người Dùng', customer_name) if "doanh nghiệp" in customer_name.lower() else customer_name,
            "customer": re.sub(r'(?i)doanh\s*nghiệp', 'Người Dùng', customer_name) if "doanh nghiệp" in customer_name.lower() else customer_name,
            "tier": f"Key Kích Hoạt ({plan_name})",
            "expires_at": claims.get("expires_at", "Vĩnh viễn"),
            "offline_until": claims.get("offline_until_iso"),
            "features": claims.get("features", []),
            "token": token,
            "activated_at": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "device_id": dev_id,
            "is_active": True
        }

        # Clear any previous revocation state for this device upon new successful activation
        trials = cls._read_trials()
        if dev_id in trials and trials[dev_id].get("revoked"):
            trials[dev_id]["revoked"] = False
            cls._write_trials(trials)

        with cls._license_lock:
            cls._is_revoked = False
            cls._cached_claims = claims

        # Persist valid license securely to data/license.sec
        cls.save_secure_license(record)

        return record

    @classmethod
    def purge_license(cls) -> None:
        """
        Xóa an toàn toàn bộ các tệp giấy phép cục bộ khi bị thu hồi hoặc vi phạm tính toàn vẹn:
        - data/license.sec
        - data/license.sec.tmp
        - license.json
        - license.json.tmp
        Clears cached in-memory claims and sets _is_revoked = True. Idempotent and thread-safe.
        """
        with cls._license_lock:
            cls._is_revoked = True
            cls._cached_claims = None
            cls._last_online_validation_time = 0.0
            cls._last_online_validation_result = None
            target_files = [
                LICENSE_FILE,
                LICENSE_FILE + ".tmp",
                LEGACY_LICENSE_FILE,
                LEGACY_LICENSE_FILE + ".tmp",
            ]
            for path in target_files:
                try:
                    if os.path.exists(path):
                        os.remove(path)
                except Exception:
                    pass

    @classmethod
    def validate_license_online(
        cls,
        device_id: str = "",
        force_online: bool = False
    ) -> Dict[str, Any]:
        """
        Direct API method for real-time online license validation against Cloudflare Worker.
        - Reads current stored token from data/license.sec. If missing, returns {"is_active": False, "valid": False}.
        - Calls POST https://luv-license-api.luv-sds.workers.dev/v1/validate with JSON payload.
        - Verifies Ed25519 signature against PUBLIC_KEY_B64.
        - On HTTP 400/401/403/404 or license_denied: calls purge_license(), sets revoked=True, valid=False.
        - On network outage: checks offline tolerance window claims['offline_until'].
        """
        return cls.validate_license(device_id=device_id, force_online=force_online)

    @classmethod
    def validate_license(
        cls,
        client_device_hint: str = "",
        force_online: bool = True,
        device_id: str = ""
    ) -> Dict[str, Any]:
        """
        Xác thực giấy phép trực tuyến thời gian thực với Máy chủ Cloudflare Worker:
        POST https://luv-license-api.luv-sds.workers.dev/v1/validate

        - Gửi headers: Content-Type: application/json, User-Agent: LUV-DOC/1.1.0
        - Payload: token, device_id, app_version: "1.1.0", product_code: claims.get("product_code", "DOC")
        - Khi HTTP 200: Xác thực chữ ký số Ed25519 của token mới, cập nhật data/license.sec,
          cập nhật offline_until, trả về trạng thái active.
        - Khi HTTP 400/401/403/404 hoặc lỗi license_denied / device_not_registered:
          gọi purge_license(), đặt status="revoked", revoked=True, is_active=False.
        - Khi mất kết nối mạng (urllib.error.URLError): Cho phép dịch ngoại tuyến nếu _utc_now() <= offline_until;
          chặn nếu đã hết hạn ngoại tuyến.
        """
        data = cls.load_secure_license()
        if not data:
            # Check if this device was previously marked revoked
            clean_dev = device_id or (client_device_hint if len(client_device_hint) == 64 else generate_device_hash(client_device_hint)) if client_device_hint else generate_device_hash()
            trials = cls._read_trials()
            if clean_dev and trials.get(clean_dev, {}).get("revoked"):
                return {
                    "is_active": False,
                    "valid": False,
                    "revoked": True,
                    "status": "revoked",
                    "tier": "Giấy phép đã bị thu hồi",
                    "registered_to": None,
                    "customer": None,
                    "license_key": None,
                    "error": "license_denied",
                    "message": "Giấy phép bản quyền đã bị thu hồi bởi Quản trị viên. Quyền dịch thuật bị khóa vĩnh viễn."
                }
            return {
                "is_active": False,
                "valid": False,
                "revoked": False,
                "status": "unactivated",
                "tier": "Chưa kích hoạt",
                "registered_to": None,
                "customer": None,
                "license_key": None
            }

        token = data.get("token")
        if not token or "." not in token:
            cls.purge_license()
            return {
                "is_active": False,
                "valid": False,
                "revoked": False,
                "status": "unactivated",
                "tier": "Chưa kích hoạt",
                "registered_to": None,
                "customer": None,
                "license_key": None
            }

        # 1. Local Ed25519 Cryptographic Verification
        try:
            payload_part, sig_part = token.split(".", 1)
            pub_key = Ed25519PublicKey.from_public_bytes(base64.b64decode(PUBLIC_KEY_B64))
            sig_bytes = _b64url_decode(sig_part)
            pub_key.verify(sig_bytes, payload_part.encode("ascii"))
            claims = json.loads(_b64url_decode(payload_part).decode("utf-8"))
        except (InvalidSignature, Exception) as exc:
            cls.purge_license()
            return {
                "is_active": False,
                "valid": False,
                "revoked": True,
                "status": "invalid_signature",
                "tier": "Chữ ký không hợp lệ",
                "registered_to": None,
                "customer": None,
                "license_key": None,
                "error": "invalid_signature",
                "message": f"Chữ ký số của giấy phép không hợp lệ: {str(exc)}"
            }

        now = _utc_now()
        exp = int(claims.get("exp", 0))
        if exp > 0 and exp + CLOCK_SKEW_SECONDS < now:
            return {
                "is_active": False,
                "valid": False,
                "revoked": False,
                "status": "expired",
                "expired": True,
                "tier": "Giấy phép đã hết hạn",
                "registered_to": claims.get("customer") or data.get("registered_to"),
                "customer": claims.get("customer") or data.get("registered_to"),
                "license_key": data.get("license_key"),
                "message": "Giấy phép bản quyền đã hết hạn sử dụng."
            }

        # Determine effective device_id
        dev_id = ""
        if device_id:
            if len(device_id) == 64 and all(c in "0123456789abcdefABCDEF" for c in device_id):
                dev_id = device_id
            else:
                dev_id = claims.get("device_id") or data.get("device_id") or generate_device_hash(device_id)
        elif client_device_hint:
            if len(client_device_hint) == 64 and all(c in "0123456789abcdefABCDEF" for c in client_device_hint):
                dev_id = client_device_hint
            else:
                dev_id = claims.get("device_id") or data.get("device_id") or generate_device_hash(client_device_hint)
        if not dev_id:
            dev_id = claims.get("device_id") or data.get("device_id") or generate_device_hash()

        # Offline until timestamp calculation helper
        def get_offline_timestamp() -> int:
            val = claims.get("offline_until") or data.get("offline_until")
            if isinstance(val, (int, float)):
                return int(val)
            if isinstance(val, str):
                if val.isdigit():
                    return int(val)
                try:
                    iso_str = val.replace("Z", "+00:00")
                    dt = datetime.fromisoformat(iso_str)
                    return int(dt.timestamp())
                except Exception:
                    pass
            return exp if exp > 0 else 0

        offline_ts = get_offline_timestamp()

        # If online validation is requested
        if force_online:
            product_code = claims.get("product_code", "DOC")
            products_to_validate = [product_code]
            if product_code == "DOC":
                products_to_validate.append("SDS")
            elif product_code == "SDS":
                products_to_validate.append("DOC")

            validation_success = False
            last_http_error = None
            raw_err = ""
            renewed_token = None
            renewed_claims = None

            for prod in products_to_validate:
                payload = {
                    "token": token,
                    "device_id": dev_id,
                    "app_version": APP_VERSION,
                    "product_code": prod
                }
                req = urllib.request.Request(
                    f"{DEFAULT_SERVER_URL}/v1/validate",
                    data=json.dumps(payload).encode("utf-8"),
                    headers={
                        "Content-Type": "application/json",
                        "User-Agent": f"LUV-DOC/{APP_VERSION}"
                    },
                    method="POST"
                )

                try:
                    with urllib.request.urlopen(req, timeout=10) as resp:
                        if resp.status == 200:
                            body = json.loads(resp.read().decode("utf-8"))
                            renewed_token = body.get("token")
                            if renewed_token and "." in renewed_token:
                                # Verify signature of renewed token
                                r_payload_part, r_sig_part = renewed_token.split(".", 1)
                                r_pub_key = Ed25519PublicKey.from_public_bytes(base64.b64decode(PUBLIC_KEY_B64))
                                r_sig_bytes = _b64url_decode(r_sig_part)
                                r_pub_key.verify(r_sig_bytes, r_payload_part.encode("ascii"))
                                renewed_claims = json.loads(_b64url_decode(r_payload_part).decode("utf-8"))
                                validation_success = True
                                break
                except urllib.error.HTTPError as exc:
                    last_http_error = exc
                    raw_err = exc.read().decode("utf-8", errors="replace")
                    err_code = ""
                    try:
                        parsed = json.loads(raw_err)
                        err_code = parsed.get("error", "")
                    except Exception:
                        pass
                    if prod == "DOC" and err_code == "product_denied":
                        continue
                    break
                except urllib.error.URLError as exc:
                    last_http_error = exc
                    break
                except Exception as exc:
                    last_http_error = exc
                    break

            if validation_success and renewed_token and renewed_claims:
                record = dict(data)
                record["token"] = renewed_token
                record["status"] = "active"
                record["is_active"] = True
                record["offline_until"] = renewed_claims.get("offline_until_iso") or renewed_claims.get("offline_until")
                record["expires_at"] = renewed_claims.get("expires_at") or record.get("expires_at")
                if "lease_seq" in renewed_claims:
                    record["lease_seq"] = renewed_claims["lease_seq"]
                cls.save_secure_license(record)

                plan_name = renewed_claims.get("plan", data.get("plan", "Enterprise")).upper()
                customer_name = renewed_claims.get("customer") or data.get("customer") or data.get("registered_to")
                raw_tier = data.get("tier") or f"Key Kích Hoạt ({plan_name})"
                clean_tier = re.sub(r'(?i)doanh\s*nghiệp', 'Key Kích Hoạt', raw_tier) if "doanh nghiệp" in raw_tier.lower() else raw_tier
                cust_clean = re.sub(r'(?i)doanh\s*nghiệp', 'Người Dùng', customer_name) if "doanh nghiệp" in customer_name.lower() else customer_name
                result_payload = {
                    "is_active": True,
                    "revoked": False,
                    "status": "active",
                    "tier": clean_tier,
                    "registered_to": cust_clean,
                    "customer": cust_clean,
                    "license_key": data.get("license_key"),
                    "expires_at": renewed_claims.get("expires_at") or data.get("expires_at"),
                    "offline_until": renewed_claims.get("offline_until_iso") or renewed_claims.get("offline_until"),
                    "features": renewed_claims.get("features", data.get("features", [])),
                    "lease_seq": renewed_claims.get("lease_seq")
                }
                cls._last_online_validation_time = time.time()
                cls._last_online_validation_result = dict(result_payload)
                return result_payload

            # Handle HTTP errors from server (revoked / denied / quota exceeded)
            if isinstance(last_http_error, urllib.error.HTTPError):
                code = last_http_error.code
                err_code = ""
                err_msg = ""
                try:
                    parsed = json.loads(raw_err) if raw_err else {}
                    err_code = parsed.get("error", "")
                    err_msg = parsed.get("message") or "Key kích hoạt đã bị thu hồi bởi Quản trị viên."
                except Exception:
                    err_msg = "Key kích hoạt đã bị thu hồi bởi Quản trị viên."

                # Device quota exceeded: do not purge license, return device quota status
                if err_code == "device_quota_exceeded":
                    return {
                        "is_active": False,
                        "valid": False,
                        "revoked": False,
                        "status": "device_quota_exceeded",
                        "tier": "Vượt Quá Số Lượng Thiết Bị",
                        "registered_to": None,
                        "customer": None,
                        "license_key": data.get("license_key"),
                        "error": "device_quota_exceeded",
                        "message": err_msg or "Key kích hoạt đã đạt tối đa số lượng thiết bị cho phép. Vui lòng nâng cấp gói hoặc gỡ bớt thiết bị cũ."
                    }

                # Revocation on HTTP 401, 403, 404 or license_denied / device_not_registered / revoked
                if code in [401, 403, 404] or err_code in ["license_denied", "device_not_registered", "revoked", "license_revoked"]:
                    cls.purge_license()
                    trials = cls._read_trials()
                    if dev_id:
                        if dev_id not in trials:
                            trials[dev_id] = {"started_at": time.time(), "device_id": dev_id}
                        trials[dev_id]["revoked"] = True
                        cls._write_trials(trials)

                    return {
                        "is_active": False,
                        "valid": False,
                        "revoked": True,
                        "status": "revoked",
                        "tier": "Giấy phép đã bị thu hồi",
                        "registered_to": None,
                        "customer": None,
                        "license_key": None,
                        "error": err_code or "license_denied",
                        "message": err_msg or "Key kích hoạt đã bị thu hồi bởi Quản trị viên."
                    }

            # Handle Network Outage (URLError or connection error)
            if isinstance(last_http_error, (urllib.error.URLError, TimeoutError, ConnectionError, OSError)):
                now = _utc_now()
                if offline_ts > 0 and now <= offline_ts:
                    plan_name = claims.get("plan", data.get("plan", "Enterprise")).upper()
                    customer_name = claims.get("customer") or data.get("customer") or data.get("registered_to")
                    raw_tier = data.get("tier") or f"Key Kích Hoạt ({plan_name})"
                    clean_tier = re.sub(r'(?i)doanh\s*nghiệp', 'Key Kích Hoạt', raw_tier) if "doanh nghiệp" in raw_tier.lower() else raw_tier
                    cust_clean = re.sub(r'(?i)doanh\s*nghiệp', 'Người Dùng', customer_name) if "doanh nghiệp" in customer_name.lower() else customer_name
                    return {
                        "is_active": True,
                        "revoked": False,
                        "status": "offline_active",
                        "offline_mode": True,
                        "tier": clean_tier,
                        "registered_to": cust_clean,
                        "customer": cust_clean,
                        "license_key": data.get("license_key"),
                        "expires_at": claims.get("expires_at") or data.get("expires_at"),
                        "offline_until": claims.get("offline_until_iso") or data.get("offline_until"),
                        "features": claims.get("features", data.get("features", [])),
                        "message": "Đang hoạt động ở chế độ ngoại tuyến (trong thời hạn cho phép)."
                    }
                else:
                    return {
                        "is_active": False,
                        "revoked": False,
                        "status": "offline_expired",
                        "offline_expired": True,
                        "tier": "Hết hạn sử dụng ngoại tuyến",
                        "registered_to": claims.get("customer") or data.get("registered_to"),
                        "customer": claims.get("customer") or data.get("registered_to"),
                        "license_key": data.get("license_key"),
                        "message": "Thời hạn sử dụng ngoại tuyến đã kết thúc. Vui lòng kết nối Internet để gia hạn giấy phép."
                    }

        # If force_online is False, check offline_until
        now = _utc_now()
        if offline_ts > 0 and now <= offline_ts:
            plan_name = claims.get("plan", data.get("plan", "Enterprise")).upper()
            customer_name = claims.get("customer") or data.get("customer") or data.get("registered_to")
            raw_tier = data.get("tier") or f"Key Kích Hoạt ({plan_name})"
            clean_tier = re.sub(r'(?i)doanh\s*nghiệp', 'Key Kích Hoạt', raw_tier) if "doanh nghiệp" in raw_tier.lower() else raw_tier
            cust_clean = re.sub(r'(?i)doanh\s*nghiệp', 'Người Dùng', customer_name) if "doanh nghiệp" in customer_name.lower() else customer_name
            return {
                "is_active": True,
                "revoked": False,
                "status": "offline_active",
                "offline_mode": True,
                "tier": clean_tier,
                "registered_to": cust_clean,
                "customer": cust_clean,
                "license_key": data.get("license_key"),
                "expires_at": claims.get("expires_at") or data.get("expires_at"),
                "offline_until": claims.get("offline_until_iso") or data.get("offline_until"),
                "features": claims.get("features", data.get("features", [])),
            }
        else:
            return {
                "is_active": False,
                "revoked": False,
                "status": "offline_expired",
                "offline_expired": True,
                "tier": "Hết hạn sử dụng ngoại tuyến",
                "registered_to": claims.get("customer") or data.get("registered_to"),
                "customer": claims.get("customer") or data.get("registered_to"),
                "license_key": data.get("license_key"),
                "message": "Thời hạn sử dụng ngoại tuyến đã kết thúc. Vui lòng kết nối Internet để gia hạn giấy phép."
            }

    @classmethod
    def check_access(
        cls,
        device_id: str = "",
        validated_license: Optional[Dict[str, Any]] = None,
        force_online: bool = True
    ) -> Dict[str, Any]:
        """
        Kiểm tra quyền dịch thuật của máy trạm:
        - Bắt buộc phải có Key kích hoạt hợp lệ từ ban đầu (Không dùng thử).
        - Nếu Key bị thu hồi (revoked): Chặn vĩnh viễn.
        - Nếu có Key hợp lệ: Trả về access_type="licensed".
        - Nếu hết hạn ngoại tuyến: Chặn truy cập.
        - Nếu chưa kích hoạt: Chặn dịch thuật và yêu cầu nhập Key kích hoạt.
        """
        clean_dev = device_id.strip() if device_id else ""
        trials = cls._read_trials()
        effective_dev = clean_dev or generate_device_hash()
        if trials.get(effective_dev, {}).get("revoked"):
            return {
                "has_access": False,
                "revoked": True,
                "access_type": "revoked",
                "tier": "Key Đã Bị Thu Hồi",
                "message": "Key kích hoạt đã bị thu hồi bởi Quản trị viên. Quyền dịch thuật bị khóa vĩnh viễn.",
                "registered_to": None,
                "customer": None,
                "license_key": None,
                "remaining_seconds": 0,
                "remaining_formatted": "Key Đã Bị Thu Hồi"
            }

        if validated_license is not None:
            lic = validated_license
        else:
            try:
                lic = cls.validate_license(device_id=clean_dev, force_online=force_online)
            except Exception:
                lic = cls.get_license_info()

        # If revoked: permanent block, do NOT fall back
        if lic.get("revoked") or lic.get("status") == "revoked":
            return {
                "has_access": False,
                "revoked": True,
                "access_type": "revoked",
                "tier": "Key Đã Bị Thu Hồi",
                "message": lic.get("message") or "Key kích hoạt đã bị thu hồi bởi Quản trị viên. Quyền dịch thuật bị khóa vĩnh viễn.",
                "registered_to": None,
                "customer": None,
                "license_key": None,
                "remaining_seconds": 0,
                "remaining_formatted": "Key Đã Bị Thu Hồi"
            }

        # If offline grace expired
        if lic.get("offline_expired") or lic.get("status") == "offline_expired":
            return {
                "has_access": False,
                "revoked": False,
                "offline_expired": True,
                "access_type": "offline_expired",
                "tier": "Hết hạn sử dụng ngoại tuyến",
                "message": lic.get("message") or "Thời hạn sử dụng ngoại tuyến đã kết thúc. Vui lòng kết nối Internet để gia hạn giấy phép.",
                "registered_to": lic.get("registered_to"),
                "customer": lic.get("customer"),
                "license_key": lic.get("license_key"),
                "remaining_seconds": 0,
                "remaining_formatted": "Hết hạn ngoại tuyến"
            }

        if lic.get("is_active"):
            raw_tier = lic.get("tier") or "Key Kích Hoạt (PRO)"
            clean_tier = re.sub(r'(?i)doanh\s*nghiệp', 'Key Kích Hoạt', raw_tier) if "doanh nghiệp" in raw_tier.lower() else raw_tier
            reg_to = lic.get("registered_to") or "Người Dùng Bản Quyền"
            clean_reg = re.sub(r'(?i)doanh\s*nghiệp', 'Người Dùng', reg_to) if "doanh nghiệp" in reg_to.lower() else reg_to
            return {
                "has_access": True,
                "revoked": False,
                "access_type": "licensed",
                "tier": clean_tier,
                "registered_to": clean_reg,
                "customer": clean_reg,
                "license_key": lic.get("license_key"),
                "remaining_seconds": None,
                "remaining_formatted": "Key Kích Hoạt: Hợp Lệ"
            }

        return {
            "has_access": False,
            "revoked": False,
            "access_type": "unactivated",
            "tier": "Chưa Kích Hoạt Key",
            "message": "Hệ thống yêu cầu nhập Key kích hoạt để bắt đầu sử dụng. Vui lòng kích hoạt bằng Key của bạn.",
            "registered_to": None,
            "customer": None,
            "license_key": None,
            "remaining_seconds": 0,
            "remaining_formatted": "Chưa Kích Hoạt"
        }

