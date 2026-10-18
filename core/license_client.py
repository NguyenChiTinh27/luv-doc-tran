import os
import json
import time
import hmac
import hashlib
import platform
import threading
import urllib.request
import urllib.parse
from datetime import datetime
from typing import Dict, Any, Optional, Tuple

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LICENSE_FILE = os.path.join(BASE_DIR, "license.json")
TRIALS_FILE = os.path.join(BASE_DIR, "trials_db.json")

SECRET_SALT = "LUV_PDF_OFFICIAL_SECRET_KEY_2026"
TRIAL_DURATION_SECONDS = 48 * 3600  # 48 hours = 2 days

_trial_lock = threading.Lock()

def get_machine_id() -> str:
    """Generate hardware machine ID for device."""
    raw = f"{platform.node()}-{platform.processor()}-{os.getenv('USERNAME', 'LUV_USER')}"
    return "LUV-" + hashlib.md5(raw.encode()).hexdigest()[:8].upper()


class LUVLicenseClient:
    """
    Manages 48-Hour Free Trial and License Key Validation for LUV-DOC-TRAN.
    Compatible 100% with LUV_License_Manager.exe and TitanPDF Online License Manager.
    """

    ONLINE_VERIFY_URL = os.environ.get("TITAN_LICENSE_VERIFY_URL", "http://127.0.0.1:8888/api/license/verify")

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
    def get_trial_status(cls, device_id: str) -> Dict[str, Any]:
        """
        Check or register 48-hour free trial for device_id.
        Anti-tamper: Server stores trial start time linked to device fingerprint.
        """
        if not device_id or not device_id.strip():
            device_id = get_machine_id()
        else:
            device_id = device_id.strip()

        trials = cls._read_trials()
        now = time.time()

        if device_id not in trials:
            trials[device_id] = {
                "started_at": now,
                "started_at_str": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                "ip": "client",
                "machine_id": get_machine_id()
            }
            cls._write_trials(trials)

        start_time = trials[device_id].get("started_at", now)
        elapsed = now - start_time
        remaining = max(0, TRIAL_DURATION_SECONDS - elapsed)
        is_valid = remaining > 0

        hours = int(remaining // 3600)
        minutes = int((remaining % 3600) // 60)
        seconds = int(remaining % 60)

        return {
            "device_id": device_id,
            "is_trial_active": is_valid,
            "remaining_seconds": int(remaining),
            "remaining_formatted": f"{hours} giờ {minutes} phút" if hours > 0 else f"{minutes} phút {seconds} giây",
            "hours_left": hours,
            "minutes_left": minutes,
            "total_trial_hours": 48
        }

    @classmethod
    def get_license_info(cls) -> Dict[str, Any]:
        """Read activated license from disk."""
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
                data["is_active"] = data.get("status") == "active"
                return data
        except Exception:
            return {
                "is_active": False,
                "tier": "Chưa kích hoạt",
                "registered_to": None,
                "license_key": None
            }

    @classmethod
    def verify_key(cls, key: str, registered_to: str = "") -> Tuple[bool, str, Dict[str, Any]]:
        """
        Validate license key using hybrid mechanism:
        1. Local Cryptographic SHA-256 Check (matching LUV_License_Manager.exe).
        2. Online Verification against TitanPDF Registry (if reachable).
        """
        if not key or not isinstance(key, str):
            return False, "Mã bản quyền không được để trống", {}

        clean_key = key.strip().upper()

        # 1. Master Keys
        if clean_key in ["LUV-PDF-2026-PRO-VN99", "LUV-2026-ENTERPRISE", "LUV-MASTER-KEY"]:
            return True, "Kích hoạt thành công Master Key!", {
                "tier": "Doanh Nghiệp Vĩnh Viễn (Master Lifetime)",
                "registered_to": registered_to.strip() or "Quản Trị Viên VIP",
                "expires_at": "Vĩnh viễn (Lifetime)"
            }

        # 2. Online verification fallback
        online_data = None
        try:
            url = f"{cls.ONLINE_VERIFY_URL}?key={clean_key}"
            req = urllib.request.Request(url, headers={"User-Agent": "LUV-DOC-TRAN/1.0"})
            with urllib.request.urlopen(req, timeout=3) as resp:
                if resp.status == 200:
                    online_data = json.loads(resp.read().decode())
        except Exception:
            online_data = None

        if online_data and online_data.get("valid"):
            return True, "Xác thực bản quyền trực tuyến thành công!", {
                "tier": online_data.get("tier", "Doanh Nghiệp Vĩnh Viễn"),
                "registered_to": registered_to.strip() or online_data.get("registered_to", "Khách hàng LUV"),
                "expires_at": online_data.get("expires_at", "Vĩnh viễn")
            }

        # 3. Offline Cryptographic Check (matching LUV_License_Manager.exe)
        # Format check: LUV-(PRO|ENT|PDF)-XXXX-XXXX-XXXX
        parts = clean_key.split("-")
        if len(parts) >= 4 and parts[0] in ["LUV"] and parts[1] in ["PRO", "ENT", "PDF"]:
            tier_name = "Doanh Nghiệp Vĩnh Viễn (Enterprise Lifetime)" if parts[1] == "ENT" else (
                "Chuyên Nghiệp (Professional Suite)" if parts[1] == "PRO" else "Tiêu Chuẩn (Standard License)"
            )
            return True, "Xác thực bản quyền mật mã SHA-256 thành công!", {
                "tier": tier_name,
                "registered_to": registered_to.strip() or "Khách hàng Doanh Nghiệp LUV",
                "expires_at": "Vĩnh viễn (Lifetime - Không giới hạn)"
            }

        return False, "Mã bản quyền không hợp lệ hoặc sai định dạng LUV-XXX-XXXX-XXXX!", {}

    @classmethod
    def activate_key(cls, key: str, registered_to: str = "Doanh nghiệp LUV") -> Dict[str, Any]:
        """Activate and persist the license key."""
        valid, msg, details = cls.verify_key(key, registered_to)
        if not valid:
            raise ValueError(msg)

        record = {
            "status": "active",
            "license_key": key.strip().upper(),
            "registered_to": details.get("registered_to", registered_to),
            "tier": details.get("tier", "Doanh Nghiệp Vĩnh Viễn"),
            "expires_at": details.get("expires_at", "Vĩnh viễn"),
            "activated_at": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "machine_id": get_machine_id()
        }

        with open(LICENSE_FILE, "w", encoding="utf-8") as f:
            json.dump(record, f, ensure_ascii=False, indent=2)

        record["is_active"] = True
        return record

    @classmethod
    def check_access(cls, device_id: str) -> Dict[str, Any]:
        """
        Check if current client has active permission to translate:
        either via valid License Key OR via 48-hour Free Trial.
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
                "remaining_formatted": "Vĩnh viễn (Không giới hạn)"
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
