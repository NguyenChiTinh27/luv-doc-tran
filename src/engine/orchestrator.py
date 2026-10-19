"""
orchestrator.py
───────────────
Top-level Pipeline Orchestrator coordinating all 5 phases:
- Phase 1: Deep PDF Extractor (geometry, text spans, vector grid, raster images)
- Phase 2: Layout Sanitizer (spatial deinterleaver, paragraph defragmenter, token cleaner)
- Phase 3: Bilingual Translator & Pairing Engine (offline zero-cost dictionary / LLM adapter)
- Phase 4: OpenXML Surgeon (dynamic reflow row unlocker, native header promoter, tab normalizer)
- Phase 5: QC/QA Gate & Self-Healing Loop (authoritative 7-point audit, SSIM fidelity, parity)
"""

from __future__ import annotations
import os
import sys
import logging
from dataclasses import dataclass, field
from typing import Optional, Tuple, Dict, Any, List

from src.engine.phase1_extractor import PDFInspector, ExtractedDocument
from src.engine.phase2_sanitizer import SanitizedDocument
from src.engine.phase3_bilingual.cost_guard import ZeroCostGuard, ApiCostViolationError
from src.engine.phase4_surgeon import OpenXmlSurgeon
from src.engine.phase5_qc_qa import MasterQualityGate, VisualQA, SelfHealingLoop, GateResult

from core.bilingual_qc_v2 import reconstruct_bilingual_document

logger = logging.getLogger("PipelineOrchestrator")


@dataclass
class EngineConfig:
    """Configuration settings for the End-to-End Pipeline."""
    engine: str = "offline"  # 'offline', 'deepseek', 'openai'
    api_key: Optional[str] = None
    enable_visual_qa: bool = True
    compaction_factor: float = 1.0
    remove_body_pagination: bool = True
    title_text: Optional[str] = None
    target_dpi: int = 150


class PipelineOrchestrator:
    """Master controller orchestrating end-to-end PDF-to-Bilingual-Word processing."""

    def __init__(self, config: Optional[EngineConfig] = None):
        self.config = config or EngineConfig()
        self.surgeon = OpenXmlSurgeon(target_dpi=self.config.target_dpi * 2)
        self.healing_loop = SelfHealingLoop(target_dpi=self.config.target_dpi)

    def run(
        self,
        input_path: str,
        output_path: Optional[str] = None
    ) -> Tuple[str, GateResult]:
        """
        Executes the 5-phase pipeline on the input document:
        1. Validates input path and resolves output path.
        2. Applies ZeroCostGuard unless external API is explicitly configured with a key.
        3. Converts and translates into line-by-line bilingual Word document.
        4. Applies OpenXML surgery: unlocks all table row heights, promotes headers/footers,
           strips double tabs, normalizes indents.
        5. Evaluates 7-Point OpenXML Audit gate and runs self-healing compaction loop if needed.
        """
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Input file not found: {input_path}")

        abs_input = os.path.abspath(input_path)
        base_name, ext = os.path.splitext(abs_input)
        ext_lower = ext.lower()

        if output_path is None:
            resolved_output = f"{base_name}_bilingual.docx"
        else:
            resolved_output = os.path.abspath(output_path)

        logger.info(f"Starting 5-Phase Pipeline: {abs_input} -> {resolved_output} (Engine: {self.config.engine})")

        # Guard cost: enforce ZeroCostGuard by default
        guard = ZeroCostGuard() if self.config.engine == "offline" else None

        def _execute_pipeline():
            # Phase 1, 2, 3: Deep Extraction, Sanitization, Bilingual Translation & Pairing
            # reconstruct_bilingual_document implements the full pipeline for PDF, DOC, DOCX
            temp_out = reconstruct_bilingual_document(
                input_path=abs_input,
                output_path=resolved_output
            )

            # Phase 4: OpenXML Surgeon
            self.surgeon.process_document(
                docx_path=resolved_output,
                output_path=resolved_output,
                compaction_factor=self.config.compaction_factor,
                title_text=self.config.title_text,
                remove_body_pagination=self.config.remove_body_pagination
            )

            # Phase 5: QC/QA Gate & Self-Healing Loop
            is_pdf = ext_lower == ".pdf"
            gate_result = self.healing_loop.run_healing_loop(
                docx_path=resolved_output,
                pdf_path=abs_input if is_pdf else None,
                enable_visual_qa=self.config.enable_visual_qa and is_pdf
            )

            return resolved_output, gate_result

        if guard:
            with guard:
                return _execute_pipeline()
        else:
            return _execute_pipeline()
