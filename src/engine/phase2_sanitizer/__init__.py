"""
src.engine.phase2_sanitizer
───────────────────────────
Phase 2: Sanitizer Engine
Spatial de-interleaving, paragraph defragmentation, token spacing & punctuation, and structure normalization.
Outputs SanitizedDocument with clean tokens, 2-column reconstructed grids, and defragmented paragraphs.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import List, Tuple, Optional, Dict, Any

from src.engine.phase1_extractor.models import (
    TextBlock, TextSpan, GridCellData, TableGridData,
    DrawingData, ImageData, ExtractedPage, ExtractedDocument
)
from .spatial_deinterleaver import SpatialDeinterleaver
from .paragraph_defragmenter import ParagraphDefragmenter
from .token_sanitizer import TokenSanitizer
from .structure_normalizer import StructureNormalizer, ROMAN_NUMERALS, GHS_SECTIONS


@dataclass
class SanitizedPage:
    """Sanitized page representation with defragmented blocks and structured tables."""
    page_num: int
    width: float
    height: float
    margins: Tuple[float, float, float, float] = (36.0, 36.0, 36.0, 36.0)
    blocks: List[TextBlock] = field(default_factory=list)
    tables: List[TableGridData] = field(default_factory=list)
    drawings: List[DrawingData] = field(default_factory=list)
    images: List[ImageData] = field(default_factory=list)
    rotation: float = 0.0

    @property
    def text(self) -> str:
        return "\n".join(b.text for b in self.blocks if b.text)


@dataclass
class SanitizedDocument:
    """Sanitized multi-page document ready for Phase 3 bilingual translation passes."""
    pages: List[SanitizedPage] = field(default_factory=list)
    pdf_path: str = ""
    metadata: Dict[str, Any] = field(default_factory=dict)

    @property
    def page_count(self) -> int:
        return len(self.pages)

    def get_page(self, page_num: int) -> Optional[SanitizedPage]:
        for p in self.pages:
            if p.page_num == page_num:
                return p
        return None

    @classmethod
    def sanitize_extracted_document(cls, doc: ExtractedDocument) -> "SanitizedDocument":
        """
        Executes full Phase 2 Sanitization pipeline:
        1. Token Sanitizer (typos, unit spacing, sentence boundary punctuation)
        2. Structure Normalizer (Roman numeral section normalization & inverted header repair)
        3. Spatial Deinterleaver (multi-field splitting, double-tab stripping, borderless 2-column tables)
        4. Paragraph Defragmenter (broken wrapped line consolidation)
        """
        # Step 1: Pre-sanitize tokens
        TokenSanitizer.sanitize_document(doc)

        # Step 2: Normalize structural section headers
        StructureNormalizer.normalize_document(doc)

        sanitized_pages: List[SanitizedPage] = []
        for p in doc.pages:
            # Step 3: Deinterleave multi-field rows and build 2-column grids
            SpatialDeinterleaver.deinterleave_page(p)

            # Step 4: Defragment body paragraphs
            ParagraphDefragmenter.defragment_page(p)

            sanitized_pages.append(SanitizedPage(
                page_num=p.page_num,
                width=p.width,
                height=p.height,
                margins=p.margins,
                blocks=p.blocks,
                tables=p.tables,
                drawings=p.drawings,
                images=p.images,
                rotation=p.rotation
            ))

        return cls(
            pages=sanitized_pages,
            pdf_path=doc.pdf_path,
            metadata=doc.metadata
        )


__all__ = [
    "SpatialDeinterleaver",
    "ParagraphDefragmenter",
    "TokenSanitizer",
    "StructureNormalizer",
    "SanitizedPage",
    "SanitizedDocument",
    "ROMAN_NUMERALS",
    "GHS_SECTIONS",
]
