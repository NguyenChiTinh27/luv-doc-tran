"""
src/engine/phase3_bilingual/rule_synthesizer.py
────────────────────────────────────────────────
100% offline zero-cost domain synthesizer for Phase 3 (R4, R7, R9).

Translates technical documents (MSDS/SDS, TDS) into Vietnamese without any external API calls:
  - Preserves CAS registry numbers (e.g. 7631-86-9, 471-34-1)
  - Preserves chemical formulas (e.g. CaCO3, SiO2, Al2O3, H2SO4)
  - Preserves numeric ranges, tolerances, and units (e.g. 15 minutes, 95-100%, ±0.05, > 1700 °C)
  - Enforces Roman numeral normalization for GHS section headers (I..XVI)
  - De-interleaves Label : Value structures and translates them independently
  - Invokes LeakHunter for post-synthesis quality repair
"""

import re
from typing import Optional, Tuple, List, Dict
from src.engine.phase3_bilingual.offline_glossary import OfflineGlossary
from src.engine.phase3_bilingual.token_sanitizer import TokenSanitizer
from src.engine.phase3_bilingual.leak_hunter import LeakHunter
from src.engine.phase3_bilingual.typography import ROMAN_NUMERALS


class RuleBasedSynthesizer:
    """
    Deterministic, offline local synthesis engine for zero-cost translation (R9).
    """

    SECTION_HEADER_RE = re.compile(
        r"^(?:"
        r"(?:SECTION|Section|MỤC|Mục|Phần)\s*([0-9]{1,2}|I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII|XIII|XIV|XV|XVI)[\.\:\s\-]+(.*)"
        r"|"
        r"(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII|XIII|XIV|XV|XVI)[\.\:\s]+(.*)"
        r"|"
        r"([1-9]|1[0-6])[\.\:\s\-]+(Identification|Hazards?|Composition|First|Fire|Accidental|Handling|Exposure|Physical|Stability|Toxicological|Ecological|Disposal|Transport|Regulatory|Other|Nhận|Thành|Biện|Kiểm|Đặc|Tính|Thông).*"
        r")$",
        re.IGNORECASE
    )

    LABEL_VALUE_RE = re.compile(r"^([^:\t\n]{2,40}?)\s*(?::|\t:|\t)\s*(.+)$")
    CAS_RE = re.compile(r"\b\d{2,7}-\d{2}-\d\b")
    NUMERIC_ONLY_RE = re.compile(r"^[\d\s\.\,\-\+\/\:\;\%\>\<\=\(\)\[\]\±\°\°C\°F\~]+$")

    # Common chemical formulas to preserve exactly
    FORMULA_RE = re.compile(
        r"\b(?:H2O|CaCO3|SiO2|Al2O3|Fe2O3|TiO2|BaSO4|ZnO|NaCl|KCl|CaCl2|MgSO4|CuSO4|"
        r"H2SO4|HCl|HNO3|H3PO4|NaOH|KOH|NH3|H2O2|CH3COOH|HCOOH|C2H5OH|CH3OH|"
        r"C3H8O|C3H6O|C4H8O2|C6H6|C7H8|C8H10|CH2Cl2|CHCl3|C2HCl3)\b"
    )

    @classmethod
    def roman_to_int(cls, roman_str: str) -> Optional[int]:
        """Converts Roman numeral or string digit to 1-based integer."""
        if not roman_str:
            return None
        s = roman_str.strip().upper()
        if s.isdigit():
            val = int(s)
            return val if 1 <= val <= 16 else None
        if s in ROMAN_NUMERALS:
            return ROMAN_NUMERALS.index(s) + 1
        return None

    @classmethod
    def int_to_roman(cls, val: int) -> str:
        """Converts 1-based integer (1..16) to Roman numeral string."""
        if 1 <= val <= 16:
            return ROMAN_NUMERALS[val - 1]
        return str(val)

    @classmethod
    def match_section_header(cls, text: str) -> Optional[Tuple[int, str]]:
        """
        Parses text for standard MSDS Section Header patterns (1..16 or I..XVI).
        Returns (section_idx, title) if matched, else None.
        """
        if not text or not text.strip():
            return None
        m = cls.SECTION_HEADER_RE.match(text.strip())
        if not m:
            return None
        num_str = m.group(1) or m.group(3) or m.group(5)
        title_str = m.group(2) or m.group(4) or m.group(6) or ""
        sec_idx = cls.roman_to_int(num_str) if num_str else None
        if sec_idx and 1 <= sec_idx <= 16:
            return sec_idx, title_str.strip()
        return None

    @classmethod
    def is_pure_data(cls, text: str) -> bool:
        """
        Determines whether text is purely non-translatable numeric data,
        CAS number, or chemical formula.
        """
        if not text:
            return True
        s = text.strip()
        if not any(c.isalpha() for c in s):
            return True
        if cls.CAS_RE.fullmatch(s):
            return True
        if cls.FORMULA_RE.fullmatch(s):
            return True
        return False

    @classmethod
    def synthesize(cls, text: str) -> str:
        """
        Synthesizes Vietnamese translation for the input English text.
        Guarantees 100% offline zero API cost.
        """
        if not text or not text.strip():
            return ""

        clean_text = TokenSanitizer.sanitize(text.strip())

        # If text is purely numbers, symbols, formula, or CAS, preserve as-is
        if cls.is_pure_data(clean_text) or cls.NUMERIC_ONLY_RE.match(clean_text):
            return clean_text

        # Unit measurements (e.g. "15 minutes" -> "15 phút")
        m_unit = re.match(r"^(\d+(?:\.\d+)?)\s*(minutes?|hours?|days?|seconds?)$", clean_text, re.IGNORECASE)
        if m_unit:
            u_val, u_unit = m_unit.group(1), m_unit.group(2).lower()
            u_map = {
                "minute": "phút", "minutes": "phút",
                "hour": "giờ", "hours": "giờ",
                "day": "ngày", "days": "ngày",
                "second": "giây", "seconds": "giây"
            }
            return f"{u_val} {u_map.get(u_unit, u_unit)}"

        # 1. Section Header detection & Roman Numeral normalization (R7)
        sec_match = cls.match_section_header(clean_text)
        if sec_match:
            sec_idx, _ = sec_match
            if sec_idx in OfflineGlossary.GHS_SECTIONS:
                sec_info = OfflineGlossary.GHS_SECTIONS[sec_idx]
                return sec_info["vi"]

        # 2. Exact match against HSE Phrases (R7)
        lower_clean = clean_text.lower().rstrip(".:;,")
        if lower_clean in OfflineGlossary.HSE_PHRASES:
            vi_phrase = OfflineGlossary.HSE_PHRASES[lower_clean]
            if clean_text.endswith("."):
                return vi_phrase + "."
            return vi_phrase

        # 3. Exact match against Chemical Names (R4)
        if lower_clean in OfflineGlossary.CHEMICAL_NAMES:
            return OfflineGlossary.CHEMICAL_NAMES[lower_clean]

        # 4. Exact match against Field Labels
        if lower_clean in OfflineGlossary.FIELD_LABELS:
            return OfflineGlossary.FIELD_LABELS[lower_clean]

        # 5. Label : Value line deconstruction
        m_lv = cls.LABEL_VALUE_RE.match(clean_text)
        if m_lv:
            label, val = m_lv.group(1).strip(), m_lv.group(2).strip()
            trans_label = cls.synthesize(label)
            trans_val = cls.synthesize(val)
            return f"{trans_label}: {trans_val}"

        # 6. Delegate non-exact matches to FullSentenceTranslator (R1: Zero Mock, Zero (Thông tin))
        from src.engine.phase3_bilingual.full_sentence_translator import FullSentenceTranslator
        return FullSentenceTranslator.translate(clean_text)
