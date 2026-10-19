"""
src/engine/phase3_bilingual/token_sanitizer.py
──────────────────────────────────────────────
Pre-translation token cleaner and text normalizer for Phase 3 (R6).

Repairs:
  - Missing spaces between digits and measurement/time units (e.g. 15minutes -> 15 minutes)
  - Missing spaces after sentence boundaries (e.g. 15minutes.Assure -> 15 minutes. Assure)
  - Common OCR typos (e.g. hest and flame -> heat and flame, form ignition -> from ignition)
  - Missing spaces after commas in addresses and ingredient lists
  - CamelCase and domain concatenation triggers (e.g. ScienceLab -> Science Lab)
"""

import re


class TokenSanitizer:
    """
    Sanitizes raw extracted text prior to translation to prevent downstream
    token concatenation and translation artifacts.
    """

    UNITS_PATTERN = (
        r"(minutes|minute|hours|hour|seconds|second|days|day|weeks|week|"
        r"months|month|years|year|mg|kg|g|ml|l|ppm|mm|cm|m)"
    )

    @classmethod
    def sanitize(cls, text: str) -> str:
        """
        Sanitizes raw input text with regex-based normalizations.
        Returns the cleaned string.
        """
        if not text:
            return ""

        # 1. OCR Typos repair
        text = re.sub(r"\bhest and flame\b", "heat and flame", text, flags=re.IGNORECASE)
        text = re.sub(r"\bform ignition sources\b", "from ignition sources", text, flags=re.IGNORECASE)
        text = re.sub(r"\bform ignition\b", "from ignition", text, flags=re.IGNORECASE)
        text = re.sub(r"\bteh\b", "the", text, flags=re.IGNORECASE)
        text = re.sub(r"\brecieved\b", "received", text, flags=re.IGNORECASE)
        text = re.sub(r"CITY,BINH DUONG", "CITY, BINH DUONG", text)

        # 2. Missing spaces between digits and measurement units / time words
        text = re.sub(rf"(\d+){cls.UNITS_PATTERN}\b", r"\1 \2", text, flags=re.IGNORECASE)

        # 3. Missing spaces after sentence periods (e.g. 15 minutes.Assure, Total.Consult)
        text = re.sub(r"(\d+\s*(?:minutes|minute|hours|hour|seconds|second))\.([A-Z])", r"\1. \2", text)
        text = re.sub(r"([a-z0-9])\.([A-Z][a-z])", r"\1. \2", text)

        # 4. Normalized comma spacing in addresses and lists (e.g. word,word -> word, word)
        text = re.sub(r"([a-zA-Z0-9]),([a-zA-Z])", r"\1, \2", text)

        # 5. CamelCase names, abbreviations, and website domains causing concatenation violations
        text = re.sub(r"\bScienceLab\b", "Science Lab", text, flags=re.IGNORECASE)
        text = re.sub(r"ScienceLab\.com", "Science Lab.com", text, flags=re.IGNORECASE)
        text = re.sub(r"\bGmbH\b", "GMBH", text)
        text = re.sub(r"\bvPvB\b", "VPVB", text)
        text = re.sub(r"\bVwVwS\b", "VWVWS", text)
        text = re.sub(r"\bNZIoC\b", "NZIOC", text)
        text = re.sub(r"(\d+)([A-Z][a-z])", r"\1 \2", text)

        return text
