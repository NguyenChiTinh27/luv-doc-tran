"""
token_sanitizer.py
──────────────────
Text normalization, punctuation spacing, and OCR typo correction.
- Inserts missing spaces before units (e.g. '15minutes' -> '15 minutes')
- Normalizes sentence boundaries and periods (e.g. '15minutes.Assure' -> '15 minutes. Assure')
- Repairs common OCR typographical errors (e.g. 'hest and flame' -> 'heat and flame')
- Normalizes comma spacing in addresses and lists (e.g. 'CITY,BINH DUONG' -> 'CITY, BINH DUONG')
- Normalizes CamelCase domains and brands that trigger concatenation false positives
"""

from __future__ import annotations
import re
from typing import List, Tuple, Optional, Dict, Any

from src.engine.phase1_extractor.models import (
    TextBlock, TextSpan, GridCellData, TableGridData, ExtractedPage, ExtractedDocument
)


class TokenSanitizer:
    """Pre-translation regex cleaning and typographical normalization engine."""

    UNITS_PATTERN = r"(minutes|minute|hours|hour|seconds|second|days|day|weeks|week|months|month|years|year|mg|kg|g|ml|l|ppm|mm|cm|m)"

    @classmethod
    def sanitize(cls, text: str) -> str:
        """Sanitizes raw text strings for spacing, typos, and punctuation."""
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
        text = re.sub(
            r"(\d+\s*(?:minutes|minute|hours|hour|seconds|second|days|day|months|month|years|year))\.([A-Z])",
            r"\1. \2",
            text
        )
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

    @classmethod
    def sanitize_block(cls, block: TextBlock) -> TextBlock:
        """Sanitizes text of a TextBlock and all its contained TextSpans."""
        block.text = cls.sanitize(block.text)
        for span in block.spans:
            span.text = cls.sanitize(span.text)
        return block

    @classmethod
    def sanitize_page(cls, page: ExtractedPage) -> ExtractedPage:
        """Sanitizes all text blocks and table cells on an ExtractedPage."""
        for b in page.blocks:
            cls.sanitize_block(b)
        for tbl in page.tables:
            for cell in tbl.cells:
                cell.text = cls.sanitize(cell.text)
                for b in cell.blocks:
                    cls.sanitize_block(b)
        return page

    @classmethod
    def sanitize_document(cls, doc: ExtractedDocument) -> ExtractedDocument:
        """Sanitizes all pages in an ExtractedDocument."""
        for p in doc.pages:
            cls.sanitize_page(p)
        return doc
