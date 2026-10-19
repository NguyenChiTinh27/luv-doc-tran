"""
src/engine/phase3_bilingual/leak_hunter.py
──────────────────────────────────────────
Post-translation scanner and repairer for Phase 3 (R6).

Detects and repairs leaked English in translated lines:
  - Ordinals: 1st -> thứ 1, 2nd -> thứ 2, 3rd -> thứ 3, 4th..16th -> thứ N
  - Section references: section 6.2 / section6.2 -> mục 6.2
  - Untranslated standard headings and keywords: Goods, Measures, Emergency, etc.
"""

import re
from typing import List, Tuple


class LeakHunter:
    """
    Scans translated Vietnamese lines and automatically repairs leaked English
    ordinals, section references, and untranslated standard terms.
    """

    REPAIRS: List[Tuple[re.Pattern, str]] = [
        # Ordinals
        (re.compile(r"\b1st\b", re.IGNORECASE), "thứ 1"),
        (re.compile(r"\b2nd\b", re.IGNORECASE), "thứ 2"),
        (re.compile(r"\b3rd\b", re.IGNORECASE), "thứ 3"),
        (re.compile(r"\b([4-9]|1[0-6])th\b", re.IGNORECASE), r"thứ \1"),
        # Section references (capitalized first)
        (re.compile(r"\bSECTION\s*(\d+(\.\d+)?)\b"), r"MỤC \1"),
        (re.compile(r"\bSection\s*(\d+(\.\d+)?)\b"), r"Mục \1"),
        (re.compile(r"\bsection\s*(\d+(\.\d+)?)\b", re.IGNORECASE), r"mục \1"),
        # Leaked standard headings/words
        (re.compile(r"\bGoods\b", re.IGNORECASE), "Hàng hóa"),
        (re.compile(r"\bMeasures\b", re.IGNORECASE), "Biện pháp"),
        (re.compile(r"\bEmergency\b", re.IGNORECASE), "Khẩn cấp"),
        (re.compile(r"\bPrecautionary statements\b", re.IGNORECASE), "Các cảnh báo phòng ngừa"),
        (re.compile(r"\bFirst aid\b", re.IGNORECASE), "Sơ cứu"),
        (re.compile(r"\bHandling and storage\b", re.IGNORECASE), "Sử dụng và lưu trữ"),
        (re.compile(r"\bExposure controls\b", re.IGNORECASE), "Kiểm soát phơi nhiễm"),
        (re.compile(r"\bPhysical and chemical\b", re.IGNORECASE), "Đặc tính lý hóa"),
        (re.compile(r"\bStability and reactivity\b", re.IGNORECASE), "Độ ổn định và khả năng phản ứng"),
        (re.compile(r"\bToxicological\b", re.IGNORECASE), "Độc tính"),
        (re.compile(r"\bEcological\b", re.IGNORECASE), "Sinh thái"),
        (re.compile(r"\bDisposal considerations\b", re.IGNORECASE), "Xem xét tiêu hủy"),
        (re.compile(r"\bTransport information\b", re.IGNORECASE), "Thông tin vận chuyển"),
        (re.compile(r"\bRegulatory information\b", re.IGNORECASE), "Thông tin pháp lý"),
        (re.compile(r"\bNot available\b", re.IGNORECASE), "Chưa có thông tin"),
        (re.compile(r"\bNot applicable\b", re.IGNORECASE), "Không áp dụng"),
        (re.compile(r"\bUnknown\b", re.IGNORECASE), "Không rõ"),
        (re.compile(r"\bNone\b", re.IGNORECASE), "Không có"),
        (re.compile(r"\bSee section\b", re.IGNORECASE), "Xem mục"),
        # Regulatory & Chemical Abbreviation Repairs (R1/R4)
        (re.compile(r"\bUN\s*không\s*có\b", re.IGNORECASE), "Số UN"),
        (re.compile(r"\bEMS\s*không\s*có\b", re.IGNORECASE), "Số EMS"),
        (re.compile(r"\bCAS\s*không\s*có\b", re.IGNORECASE), "Số CAS"),
        (re.compile(r"\bcan\s*Dạng\b", re.IGNORECASE), "có thể tạo thành dạng"),
        (re.compile(r"\b(can form explosive mixture)\b", re.IGNORECASE), "có thể tạo thành hỗn hợp nổ"),
    ]

    LEAK_DETECTION_PATTERNS: List[Tuple[re.Pattern, str]] = [
        (re.compile(r"\b(?:1st|2nd|3rd|[4-9]th|1[0-6]th)\b", re.IGNORECASE), "ordinal"),
        (re.compile(r"\bsection\s*\d+(\.\d+)?\b", re.IGNORECASE), "section_ref"),
        (re.compile(r"\b(?:Goods|Measures|Emergency|Precautionary statements)\b", re.IGNORECASE), "heading"),
    ]

    @classmethod
    def hunt_and_repair(cls, text: str) -> str:
        """
        Scans text and applies all repair regexes. Returns repaired text.
        """
        if not text:
            return ""
        result = text
        for pattern, replacement in cls.REPAIRS:
            result = pattern.sub(replacement, result)
        return result

    @classmethod
    def detect_leaks(cls, text: str) -> List[str]:
        """
        Detects any un-repaired English leaks or ordinals in text.
        Returns a list of leak descriptions if found, or empty list.
        """
        if not text:
            return []
        leaks = []
        for pattern, category in cls.LEAK_DETECTION_PATTERNS:
            matches = pattern.findall(text)
            if matches:
                leaks.extend([f"{category}:{m}" if isinstance(m, str) else f"{category}:{m[0]}" for m in matches])
        return leaks

    @classmethod
    def shield_abbreviations(cls, text: str) -> str:
        """
        Pre-translation token shielding: protects abbreviations like 'UN No.', 'EMS No.', 'CAS No.'
        before passing to machine translation engines.
        """
        if not text:
            return ""
        s = text
        replacements = [
            (r"(?<!\w)-\s*UN\s+No\.?", "- Số UN"),
            (r"(?<!\w)UN\s+No\.?", "Số UN"),
            (r"(?<!\w)-\s*EMS\s+No\.?", "- Số EMS"),
            (r"(?<!\w)EMS\s+No\.?", "Số EMS"),
            (r"(?<!\w)-\s*CAS\s+No\.?", "- Số CAS"),
            (r"(?<!\w)CAS\s+No\.?", "Số CAS"),
            (r"(?<!\w)-\s*IMO\s+Class\b", "- Phân loại IMO"),
            (r"(?<!\w)IMO\s+Class\b", "Phân loại IMO"),
            (r"(?<!\w)-\s*Package\s+group\b", "- Nhóm đóng gói"),
            (r"(?<!\w)Package\s+group\b", "Nhóm đóng gói"),
        ]
        for pattern, repl in replacements:
            s = re.sub(pattern, repl, s, flags=re.IGNORECASE)
        return s
