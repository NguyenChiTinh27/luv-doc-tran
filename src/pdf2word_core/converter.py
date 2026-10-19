"""
converter.py
────────────
Master Pipeline Orchestrator for the Hybrid Geometry-to-OpenXML PDF Conversion Engine.
Executes Base Structural Pass (Engine B / Engine A) -> Deep OpenXML Surgical Pass
-> Automated Visual QA & Self-Healing Loop.
"""

from __future__ import annotations
import os
import time
import logging
from typing import Optional

from src.pdf2word_core.models.schema import ConversionConfig, ConversionReport
from src.pdf2word_core.extractor.pdf_inspector import PDFInspector
from src.pdf2word_core.engines.engine_b_word_com import WordCOMEngine
from src.pdf2word_core.engines.engine_a_pymupdf import PyMuPDFLayoutEngine
from src.pdf2word_core.renderer.oxml_surgeon import OxmlSurgeon
from src.pdf2word_core.verifier.visual_qa import VisualQALoop

logger = logging.getLogger("PDF2WordConverter")


class PDF2WordConverter:
    """Master conversion engine orchestrating dual-engine baseline and deep OpenXML surgery."""

    def __init__(self, config: Optional[ConversionConfig] = None):
        self.config = config or ConversionConfig()
        self.surgeon = OxmlSurgeon(target_dpi=self.config.vector_raster_dpi)
        self.visual_qa = VisualQALoop(target_dpi=self.config.target_dpi)

    def convert(self, pdf_path: str, output_docx_path: str) -> ConversionReport:
        """
        Converts PDF to high-fidelity, fully editable DOCX with visual self-healing.
        """
        start_time = time.time()
        abs_pdf = os.path.abspath(pdf_path)
        abs_docx = os.path.abspath(output_docx_path)
        os.makedirs(os.path.dirname(abs_docx), exist_ok=True)

        report = ConversionReport(
            input_file=abs_pdf,
            output_file=abs_docx
        )

        # 1. Step 1: Deep PDF Inspection (Ground-truth geometry & typography)
        logger.info(f"Extracting ground-truth geometry from {abs_pdf}...")
        with PDFInspector(abs_pdf) as inspector:
            pages_data = inspector.extract_all_pages()
            report.pdf_pages = len(pages_data)

        # 2. Step 2: Base Structural Pass (Engine B with fallback to Engine A)
        engine_used = "none"
        use_engine_b = (
            self.config.engine_preference == "engine_b" or
            (self.config.engine_preference == "auto" and WordCOMEngine.is_available())
        )

        if use_engine_b:
            logger.info("Executing Base Pass with Engine B (MS Word COM native reflow)...")
            success = WordCOMEngine.convert_to_docx(abs_pdf, abs_docx)
            if success:
                engine_used = "Engine B (Word COM)"
            else:
                logger.warning("Engine B failed, falling back to Engine A (PyMuPDF)...")
                success = PyMuPDFLayoutEngine.convert_to_docx(abs_pdf, abs_docx)
                if success:
                    engine_used = "Engine A (PyMuPDF Fallback)"
        else:
            logger.info("Executing Base Pass with Engine A (PyMuPDF Layout Engine)...")
            success = PyMuPDFLayoutEngine.convert_to_docx(abs_pdf, abs_docx)
            if success:
                engine_used = "Engine A (PyMuPDF)"

        if not success or not os.path.exists(abs_docx):
            report.success = False
            report.warnings.append("Base conversion failed to generate output DOCX.")
            report.execution_time_seconds = time.time() - start_time
            return report

        report.engine_used = engine_used

        # 3. Step 3: Deep OpenXML Surgical Pass & Visual QA Self-Healing Loop
        if self.config.enable_visual_qa and self.config.enable_self_heal:
            logger.info("Running Deep OpenXML Surgical Pass with Visual QA Self-Healing Loop...")
            parity_match, ssim, passes = self.visual_qa.run_self_healing_loop(
                original_pdf_path=abs_pdf,
                docx_path=abs_docx,
                pages_data=pages_data,
                surgeon=self.surgeon,
                max_passes=self.config.max_heal_passes
            )
            report.page_count_match = parity_match
            report.ssim_score = ssim
            report.passes_executed = passes
        elif self.config.enable_surgical_pass:
            logger.info("Applying single-pass OpenXML Deep Surgery...")
            self.surgeon.perform_surgical_pass(
                docx_path=abs_docx,
                pages_data=pages_data,
                original_pdf_path=abs_pdf,
                compaction_factor=1.0
            )
            report.passes_executed = 1

        report.success = True
        report.execution_time_seconds = time.time() - start_time
        logger.info(
            f"Conversion completed successfully in {report.execution_time_seconds:.2f}s! "
            f"Engine: {report.engine_used}, SSIM: {report.ssim_score:.4f}"
        )
        return report
