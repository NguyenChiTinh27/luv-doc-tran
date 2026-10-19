"""
src/engine/phase3_bilingual/bilingual_translator.py
────────────────────────────────────────────────────
Unified Bilingual Translation Coordinator for Phase 3 (R1..R9).

Integrates:
  - Offline zero-cost synthesis via RuleBasedSynthesizer (default)
  - Runtime cost protection via ZeroCostGuard (active by default)
  - Optional external LLM translation via ExternalLLMClient
  - Visual pairing via PairingEngine (Pattern A dual-paragraph & Pattern B single-cell)
"""

import os
from pathlib import Path
from typing import Optional, Union, Any, List, Dict, Tuple
from dataclasses import dataclass, field
import docx

from src.engine.phase3_bilingual.cost_guard import ZeroCostGuard, ApiCostViolationError
from src.engine.phase3_bilingual.rule_synthesizer import RuleBasedSynthesizer
from src.engine.phase3_bilingual.offline_glossary import OfflineGlossary
from src.engine.phase3_bilingual.token_sanitizer import TokenSanitizer
from src.engine.phase3_bilingual.leak_hunter import LeakHunter
from src.engine.phase3_bilingual.pairing_engine import (
    PairingEngine,
    PairingPattern,
    BodyPair,
    CellPair,
    PairedDocument
)
from src.engine.phase3_bilingual.typography import (
    DEFAULT_FONT_NAME,
    DEFAULT_BODY_FONT_SIZE_PT,
    DEFAULT_TABLE_FONT_SIZE_PT
)
from src.engine.phase3_bilingual.api_client import ExternalLLMClient


@dataclass
class TranslationConfig:
    """Configuration options for BilingualTranslator."""
    mode: str = "offline"  # "offline", "deepseek", "openai"
    api_key: Optional[str] = None
    base_url: Optional[str] = None
    model: Optional[str] = None
    zero_cost: bool = True
    font_name: str = DEFAULT_FONT_NAME
    body_font_size: float = DEFAULT_BODY_FONT_SIZE_PT
    table_font_size: float = DEFAULT_TABLE_FONT_SIZE_PT


class BilingualTranslator:
    """
    Main entry point for Phase 3 Bilingual Translation & Pairing.
    Default execution enforces zero paid API cost and full offline dictionary pairing.
    """

    def __init__(self, config: Optional[TranslationConfig] = None):
        self.config = config or TranslationConfig()
        self.mode = self.config.mode.lower()
        self.zero_cost = self.config.zero_cost

        if self.mode != "offline" and self.config.api_key:
            self.llm_client: Optional[ExternalLLMClient] = ExternalLLMClient(
                api_key=self.config.api_key,
                provider=self.mode,
                base_url=self.config.base_url,
                model=self.config.model,
                fallback_to_offline=True
            )
        else:
            self.llm_client = None

        # Activate ZeroCostGuard if offline and zero_cost is True
        if self.zero_cost and self.mode == "offline":
            ZeroCostGuard.enable()

    def translate_text(self, text: str, context: Optional[str] = None) -> str:
        """
        Translates a single English text string into Vietnamese.
        """
        if not text or not text.strip():
            return ""

        from src.engine.phase3_bilingual.full_sentence_translator import FullSentenceTranslator
        return FullSentenceTranslator.translate(text, context=context, llm_client=self.llm_client)

    # Alias for synthesize
    def synthesize(self, text: str) -> str:
        return self.translate_text(text)

    def pair_paragraph(
        self,
        parent_container: Any,
        en_text: str,
        vi_text: Optional[str] = None,
        is_bold: bool = False,
        insert_after_p: Any = None
    ) -> Tuple[Any, Any]:
        """Creates a Pattern A dual-paragraph pair."""
        if vi_text is None:
            vi_text = self.translate_text(en_text)
        return PairingEngine.apply_pattern_a(
            parent_container,
            en_text=en_text,
            vi_text=vi_text,
            font_name=self.config.font_name,
            font_size_pt=self.config.body_font_size,
            is_bold=is_bold,
            insert_after_p=insert_after_p
        )

    def pair_cell(
        self,
        cell: Any,
        en_text: str,
        vi_text: Optional[str] = None,
        is_bold: bool = False
    ) -> Any:
        """Creates a Pattern B single-cell line break pair."""
        if vi_text is None:
            vi_text = self.translate_text(en_text)
        return PairingEngine.apply_pattern_b(
            cell,
            en_text=en_text,
            vi_text=vi_text,
            font_name=self.config.font_name,
            font_size_pt=self.config.table_font_size,
            is_bold=is_bold
        )

    def translate_document(
        self,
        doc: Any,
        config: Optional[TranslationConfig] = None
    ) -> PairedDocument:
        """
        Translates and pairs all translatable elements in a document.
        Supports python-docx Document instances or SanitizedDocument structures.
        """
        active_config = config or self.config

        if isinstance(doc, (str, Path)):
            doc = docx.Document(str(doc))

        # Check if doc is a python-docx Document
        if hasattr(doc, "paragraphs") and hasattr(doc, "tables"):
            return PairingEngine.pair_document(doc, synthesizer=self)

        # For abstract data structures (e.g. SanitizedDocument)
        paired_doc = PairedDocument(doc)
        if hasattr(doc, "blocks"):
            for block in doc.blocks:
                text = getattr(block, "text", "")
                if PairingEngine.is_translatable(text):
                    vi = self.translate_text(text)
                    paired_doc.body_pairs.append(BodyPair(block, None, text, vi))

        if hasattr(doc, "tables"):
            for tbl in doc.tables:
                cells = getattr(tbl, "cells", [])
                for cell in cells:
                    text = getattr(cell, "text", "")
                    if PairingEngine.is_translatable(text):
                        vi = self.translate_text(text)
                        paired_doc.cell_pairs.append(CellPair(cell, text, vi))

        return paired_doc

    def process_docx(
        self,
        input_path: Union[str, Path],
        output_path: Union[str, Path]
    ) -> str:
        """
        Reads input DOCX, pairs all translatable content, and writes output DOCX.
        """
        doc = docx.Document(str(input_path))
        self.translate_document(doc)
        doc.save(str(output_path))
        return str(output_path)
