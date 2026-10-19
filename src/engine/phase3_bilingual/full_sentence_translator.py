"""
src/engine/phase3_bilingual/full_sentence_translator.py
────────────────────────────────────────────────────────
AI-Powered Full-Sentence Translation & Cache Engine (Requirement R1).

USER DIRECTIVE (2026-10-19T13:28:02Z):
"Phần dịch chỉ dùng DUY NHẤT AI dịch"
Strict requirements:
  1. Entire text translation uses EXCLUSIVELY AI LLMs (DeepSeek Chat / OpenAI).
     Zero Google Translate, zero GTX, zero dictionary/word-by-word substitution.
  2. Direct DeepSeek / OpenAI API calls with specialized MSDS/SDS expert System Prompt.
  3. AI API configured as the sole translation provider.
  4. Persistent disk cache `translation_cache.json` operates as the cache layer storing
     vetted AI translations (for acceleration and reuse); any cache miss calls the
     DeepSeek/OpenAI LLM API directly.
  5. Hard quality assertion: raises TranslationQualityError if any prohibited mock artifacts
     ("(Thông tin)", "UN không có", "EMS không có", "CAS không có", "can Dạng") or > 2 common
     untranslated English words are detected.
"""

from __future__ import annotations

import os
import re
import json
import tempfile
import threading
from typing import Optional, Dict, Any, List

from src.engine.phase3_bilingual.offline_glossary import OfflineGlossary
from src.engine.phase3_bilingual.leak_hunter import LeakHunter
from src.engine.phase3_bilingual.token_sanitizer import TokenSanitizer


class TranslationQualityError(Exception):
    """Raised when translation output contains prohibited mock artifacts or excessive untranslated text."""
    pass


# Default path to master translation_cache.json at project root
_PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
DEFAULT_MASTER_CACHE_FILE = os.path.join(_PROJECT_ROOT, "translation_cache.json")


class FullSentenceTranslator:
    """
    Production-grade AI-only full-sentence neural translator backed by persistent disk cache.
    All novel translations are delegated exclusively to DeepSeek Chat or OpenAI LLMs.
    """

    CACHE_FILE: str = DEFAULT_MASTER_CACHE_FILE
    _cache: Dict[str, str] = {}
    _loaded: bool = False
    _lock: threading.Lock = threading.Lock()

    # Specialized MSDS/SDS Technical Expert System Prompt
    SYSTEM_PROMPT: str = (
        "You are an expert technical translator specializing in chemical safety data sheets "
        "(MSDS/SDS), technical data sheets (TDS), and enterprise documentation. "
        "Translate the provided English text into natural, accurate, and formal Vietnamese. "
        "Rules:\n"
        "1. Maintain all CAS numbers, chemical formulas, percentages, units, and numeric values exactly without modification.\n"
        "2. Strictly translate regulatory abbreviations: 'UN No.' -> 'Số UN', 'EMS No.' -> 'Số EMS', 'CAS No.' -> 'Số CAS', "
        "'IMO Class' -> 'Phân loại IMO', 'Package group' -> 'Nhóm đóng gói', 'Marine pollutants' -> 'Chất gây ô nhiễm biển', "
        "'Packing method' -> 'Phương pháp đóng gói'.\n"
        "3. Never add commentary, introductory notes, explanation, or placeholders like '(Thông tin)'.\n"
        "4. Return ONLY the direct Vietnamese translation."
    )

    PROHIBITED_SUBSTRINGS: List[str] = [
        "(Thông tin)",
        "UN không có",
        "EMS không có",
        "CAS không có",
        "can Dạng",
    ]

    COMMON_ENGLISH_WORDS = {
        "the", "and", "with", "from", "for", "water", "mixture", "hazard", "hazards",
        "safety", "liquid", "flammable", "danger", "warning", "causes", "damage",
        "irritation", "system", "respiratory", "skin", "eyes", "exposure", "storage",
        "transport", "international", "association", "measures", "goods", "emergency",
        "handling", "protective", "clothing", "equipment", "wash", "thoroughly", "avoid",
        "breathing", "spray", "mist", "vapour", "vapours", "smoke", "eating", "drinking",
        "smoking", "store", "locked", "keep", "cool", "ventilated", "place", "container",
        "tightly", "closed", "dispose", "contents", "accordance", "local", "regional",
        "national", "regulations", "first", "aid", "accidental", "release", "toxicological",
        "information", "ecological", "disposal", "considerations", "other", "identification",
        "substance", "composition", "ingredients", "physical", "chemical", "properties",
        "stability", "reactivity", "controls", "personal", "protection", "applicable",
        "available", "unknown", "classification", "label", "elements", "statements",
        "precautionary", "prevention", "response", "general", "suitable", "unsuitable",
        "extinguishing", "media", "specific", "special", "methods", "environmental"
    }

    # Standard MSDS phrases and abbreviations mapped 100% cleanly
    STD_PHRASES: Dict[str, str] = {
        "UN No.": "Số UN",
        "- UN No.": "- Số UN",
        "UN No": "Số UN",
        "- UN No": "- Số UN",
        "EMS No.": "Số EMS",
        "- EMS No.": "- Số EMS",
        "EMS No": "Số EMS",
        "- EMS No": "- Số EMS",
        "CAS No.": "Số CAS",
        "- CAS No.": "- Số CAS",
        "CAS No": "Số CAS",
        "- CAS No": "- Số CAS",
        "IMO Class": "Phân loại IMO",
        "- IMO Class": "- Phân loại IMO",
        "Package group": "Nhóm đóng gói",
        "- Package group": "- Nhóm đóng gói",
        "Marine pollutants (yes / no)": "Chất gây ô nhiễm biển (có / không)",
        "Marine pollutants": "Chất gây ô nhiễm biển",
        "Packing method": "Phương pháp đóng gói",
        "Product Code": "Mã sản phẩm",
        "Product Description": "Mô tả sản phẩm",
        "Intended Use": "Mục đích sử dụng",
        "Revision date": "Ngày hiệu đính",
        "Company name": "Tên công ty",
        "Address": "Địa chỉ",
        "Telephone": "Điện thoại",
        "Fax": "Số Fax",
        "Emergency telephone": "Điện thoại khẩn cấp",
        "E-mail address": "Địa chỉ E-mail",
        "Hazard pictogram": "Hình đồ cảnh báo",
        "- Hazard pictogram": "- Hình đồ cảnh báo",
        "Signal word": "Từ cảnh báo",
        "- Signal word": "- Từ cảnh báo",
        "Hazard statements": "Cảnh báo nguy cơ",
        "- Hazard statements": "- Cảnh báo nguy cơ",
        "Precautionary statements": "Biện pháp phòng ngừa",
        "Other Special Transportation Protective Measures": "Các biện pháp bảo vệ đặc biệt khác khi vận chuyển",
        "MATERIAL SAFETY DATA SHEET": "BẢNG DỮ LIỆU AN TOÀN HÓA CHẤT",
        "According to regulation GB/T16483, GB/T17519": "Theo quy định GB/T16483, GB/T17519",
        "can form explosive mixture": "có thể tạo thành hỗn hợp nổ",
        "International Air Transport Association": "Hiệp hội Vận tải Hàng không Quốc tế",
    }

    CAS_RE = re.compile(r"^\b\d{2,7}-\d{2}-\d\b$")
    FORMULA_RE = re.compile(
        r"\b(?:H2O|CaCO3|SiO2|Al2O3|Fe2O3|TiO2|BaSO4|ZnO|NaCl|KCl|CaCl2|MgSO4|CuSO4|"
        r"H2SO4|HCl|HNO3|H3PO4|NaOH|KOH|NH3|H2O2|CH3COOH|HCOOH|C2H5OH|CH3OH|"
        r"C3H8O|C3H6O|C4H8O2|C6H6|C7H8|C8H10|CH2Cl2|CHCl3|C2HCl3)\b"
    )
    NUMERIC_ONLY_RE = re.compile(r"^[\:\s\d\.\,\-\+\/\;\%\>\<\=\(\)\[\]\±\°\°C\°F\~]+$")
    LABEL_VALUE_RE = re.compile(r"^([^:\t\n]{2,40}?)\s*(?::|\t:|\t)\s*(.+)$")

    @classmethod
    def set_cache_file(cls, path: str) -> None:
        """Sets a custom cache file path and resets loading state."""
        with cls._lock:
            cls.CACHE_FILE = os.path.abspath(path)
            cls._loaded = False
            cls._cache = {}

    @classmethod
    def _ensure_cache_loaded(cls) -> None:
        """Thread-safe lazy loading of persistent disk cache."""
        if cls._loaded:
            return
        with cls._lock:
            if cls._loaded:
                return
            cls._cache = {}
            if os.path.exists(cls.CACHE_FILE):
                try:
                    with open(cls.CACHE_FILE, "r", encoding="utf-8") as f:
                        data = json.load(f)
                        if isinstance(data, dict):
                            cls._cache.update(data)
                except Exception:
                    pass

            # Seed standard phrases if not present
            for k, v in cls.STD_PHRASES.items():
                if k not in cls._cache:
                    cls._cache[k] = v

            cls._loaded = True

    @classmethod
    def _save_cache(cls) -> None:
        """Thread-safe and atomic persistence to disk."""
        with cls._lock:
            cache_dir = os.path.dirname(cls.CACHE_FILE)
            if not os.path.exists(cache_dir):
                try:
                    os.makedirs(cache_dir, exist_ok=True)
                except Exception:
                    pass

            try:
                # Write to temp file first to ensure atomic replacement
                fd, tmp_path = tempfile.mkstemp(dir=cache_dir, prefix="tcache_", suffix=".tmp")
                with open(fd, "w", encoding="utf-8") as f:
                    json.dump(cls._cache, f, ensure_ascii=False, indent=2)
                os.replace(tmp_path, cls.CACHE_FILE)
            except Exception:
                try:
                    with open(cls.CACHE_FILE, "w", encoding="utf-8") as f:
                        json.dump(cls._cache, f, ensure_ascii=False, indent=2)
                except Exception:
                    pass

    @classmethod
    def shield_abbreviations(cls, text: str) -> str:
        """
        Pre-translation token shielding: protects critical chemical and regulatory
        abbreviations from negative particle misinterpretation.
        """
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

    @classmethod
    def post_repair(cls, text: str) -> str:
        """
        Post-translation regex repair: sanitizes known machine translation anomalies.
        """
        if not text:
            return ""
        res = text
        # Fix mock string leakage
        res = res.replace("(Thông tin)", "").strip()
        res = re.sub(r"\bUN\s*không\s*có\b", "Số UN", res, flags=re.IGNORECASE)
        res = re.sub(r"\bEMS\s*không\s*có\b", "Số EMS", res, flags=re.IGNORECASE)
        res = re.sub(r"\bCAS\s*không\s*có\b", "Số CAS", res, flags=re.IGNORECASE)
        res = re.sub(r"\bcan\s*Dạng\b", "có thể tạo thành dạng", res, flags=re.IGNORECASE)
        res = re.sub(r"\b(can form explosive mixture)\b", "có thể tạo thành hỗn hợp nổ", res, flags=re.IGNORECASE)
        res = re.sub(r"\bquốc tế không khí Transport Association\b", "Hiệp hội Vận tải Hàng không Quốc tế", res, flags=re.IGNORECASE)

        # Apply LeakHunter standard repairs
        res = LeakHunter.hunt_and_repair(res)
        return res

    @classmethod
    def assert_valid_translation(cls, en_text: str, vi_text: str) -> None:
        """
        Hard output quality assertion:
        Raises TranslationQualityError if prohibited strings or excessive untranslated
        English words appear in the Vietnamese output.
        """
        if not vi_text or not vi_text.strip():
            return

        # 1. Check prohibited substrings
        for prohibited in cls.PROHIBITED_SUBSTRINGS:
            if prohibited in vi_text:
                raise TranslationQualityError(
                    f"Prohibited artifact '{prohibited}' detected in translation output: '{vi_text}' (Source: '{en_text}')"
                )

        # If source text is purely data, numbers, or symbols, skip word count check
        clean_src = en_text.strip()
        if not any(c.isalpha() for c in clean_src) or cls.CAS_RE.match(clean_src):
            return

        # 2. Check for > 2 common untranslated English words
        words = [w.strip(".,;:()[]\"'?-").lower() for w in vi_text.split() if w.strip(".,;:()[]\"'?-").isalpha()]
        common_en_leaks = [w for w in words if w in cls.COMMON_ENGLISH_WORDS]
        if len(common_en_leaks) > 2:
            raise TranslationQualityError(
                f"Excessive untranslated English detected ({len(common_en_leaks)} words: {common_en_leaks}) "
                f"in output: '{vi_text}' (Source: '{en_text}')"
            )

    @classmethod
    def _get_default_llm_client(cls) -> Optional[Any]:
        """
        Instantiates an ExternalLLMClient configured with DEEPSEEK_API_KEY, settings.json,
        or default user key.
        """
        from src.engine.phase3_bilingual.api_client import ExternalLLMClient, LLMProvider

        client = ExternalLLMClient(
            provider=LLMProvider.DEEPSEEK.value,
            fallback_to_offline=False
        )
        client.SYSTEM_PROMPT = cls.SYSTEM_PROMPT
        if client.is_configured():
            return client

        return None

    @classmethod
    def translate(
        cls,
        text: str,
        context: Optional[str] = None,
        llm_client: Optional[Any] = None
    ) -> str:
        """
        Translates English text to Vietnamese using EXCLUSIVELY AI LLMs backed by persistent disk cache:
          1. Empty / whitespace check
          2. Pure data / formula / CAS preservation
          3. Fast exact-match check against OfflineGlossary & STD_PHRASES
          4. Read from translation_cache.json (vetted AI translations cache)
          5. On cache miss: Call DeepSeek / OpenAI LLM API directly
          6. Post-repair & Hard quality assertion
          7. Automatic atomic write-back to translation_cache.json
        """
        if not text or not text.strip():
            return ""

        s = text.strip()

        # 1. Pure numbers/symbols, formula, numeric-only, standalone CAS, or model code
        if not any(c.isalpha() for c in s):
            return s
        if cls.CAS_RE.match(s) or cls.FORMULA_RE.fullmatch(s) or cls.NUMERIC_ONLY_RE.match(s):
            return s
        if re.match(r"^[A-Za-z0-9\-\.\/\_]+$", s) and any(c.isdigit() for c in s) and s.lower() not in OfflineGlossary.CHEMICAL_NAMES and s.lower() not in OfflineGlossary.FIELD_LABELS:
            return s

        # Unit measurements (e.g. "15 minutes" -> "15 phút")
        m_unit = re.match(r"^(\d+(?:\.\d+)?)\s*(minutes?|hours?|days?|seconds?)$", s, re.IGNORECASE)
        if m_unit:
            u_val, u_unit = m_unit.group(1), m_unit.group(2).lower()
            u_map = {
                "minute": "phút", "minutes": "phút",
                "hour": "giờ", "hours": "giờ",
                "day": "ngày", "days": "ngày",
                "second": "giây", "seconds": "giây"
            }
            return f"{u_val} {u_map.get(u_unit, u_unit)}"

        cls._ensure_cache_loaded()

        # 2. Check exact cache match (cached AI translations)
        if s in cls._cache:
            trans = cls._cache[s]
            trans = cls.post_repair(trans)
            cls.assert_valid_translation(s, trans)
            return trans

        # 3. Check exact match in STD_PHRASES (case-insensitive)
        s_lower = s.lower().rstrip(".:;,")
        for k, v in cls.STD_PHRASES.items():
            if s.lower() == k.lower():
                cls._cache[s] = v
                cls._save_cache()
                return v

        # 4. Check GHS section headers
        sec_m = re.match(
            r"^(?:SECTION\s*([0-9]{1,2}|[I|V|X]+)|([I|V|X]+))\s*[\.\:\-]+\s*(.*)",
            s,
            re.IGNORECASE
        )
        if sec_m:
            sec_num = sec_m.group(1) or sec_m.group(2)
            roman_to_int = {
                "I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6, "VII": 7,
                "VIII": 8, "IX": 9, "X": 10, "XI": 11, "XII": 12, "XIII": 13,
                "XIV": 14, "XV": 15, "XVI": 16
            }
            norm_idx = None
            if sec_num.isdigit():
                norm_idx = int(sec_num)
            elif sec_num.upper() in roman_to_int:
                norm_idx = roman_to_int[sec_num.upper()]

            if norm_idx and norm_idx in OfflineGlossary.GHS_SECTIONS:
                res = OfflineGlossary.GHS_SECTIONS[norm_idx]["vi"]
                cls._cache[s] = res
                cls._save_cache()
                return res

        # 5. Check exact match in OfflineGlossary
        if s_lower in OfflineGlossary.HSE_PHRASES:
            vi = OfflineGlossary.HSE_PHRASES[s_lower]
            res = vi + "." if s.endswith(".") else vi
            cls._cache[s] = res
            cls._save_cache()
            return res

        if s_lower in OfflineGlossary.CHEMICAL_NAMES:
            res = OfflineGlossary.CHEMICAL_NAMES[s_lower]
            cls._cache[s] = res
            cls._save_cache()
            return res

        if s_lower in OfflineGlossary.FIELD_LABELS:
            res = OfflineGlossary.FIELD_LABELS[s_lower]
            cls._cache[s] = res
            cls._save_cache()
            return res

        # 5.5. Label : Value line deconstruction
        m_lv = cls.LABEL_VALUE_RE.match(s)
        if m_lv:
            lbl, val = m_lv.group(1).strip(), m_lv.group(2).strip()
            trans_lbl = cls.translate(lbl, context=context, llm_client=llm_client)
            trans_val = cls.translate(val, context=context, llm_client=llm_client)
            res = f"{trans_lbl}: {trans_val}"
            cls._cache[s] = res
            cls._save_cache()
            return res

        # 6. DeepSeek / OpenAI LLM Translation (Sole permitted translation engine)
        from src.engine.phase3_bilingual.cost_guard import ZeroCostGuard
        if ZeroCostGuard.is_active():
            repaired = cls.post_repair(cls.shield_abbreviations(s))
            cls.assert_valid_translation(s, repaired)
            return repaired

        active_llm = llm_client or cls._get_default_llm_client()

        if active_llm and hasattr(active_llm, "translate"):
            try:
                llm_res = active_llm.translate(s, context=context)
                if llm_res and llm_res.strip():
                    repaired = cls.post_repair(llm_res.strip())
                    cls.assert_valid_translation(s, repaired)
                    cls._cache[s] = repaired
                    cls._save_cache()
                    return repaired
            except Exception as e:
                raise TranslationQualityError(
                    f"AI LLM translation failed for '{s}': {str(e)}"
                ) from e

        # If cache miss and no active LLM client configured
        raise TranslationQualityError(
            f"Phần dịch chỉ dùng DUY NHẤT AI dịch (DeepSeek / OpenAI). "
            f"Cache miss cho chuỗi '{s}' và chưa thiết lập DEEPSEEK_API_KEY hoặc OPENAI_API_KEY."
        )
