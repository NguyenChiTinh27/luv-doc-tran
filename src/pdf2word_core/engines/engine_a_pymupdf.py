"""
engine_a_pymupdf.py
───────────────────
Engine A: Cross-platform Base Structural Engine.
Combines PyMuPDF geometric inspection with pdf2docx layout DOM baseline
or direct DocxBuilder fallback to construct the initial editable Word document.
"""

from __future__ import annotations
import os
import logging
from typing import Optional, List

from src.pdf2word_core.extractor.pdf_inspector import PDFInspector
from src.pdf2word_core.analyzer.table_reconstructor import TableReconstructor
from src.pdf2word_core.analyzer.style_inferencer import StyleInferencer
from src.pdf2word_core.renderer.docx_builder import DocxBuilder

logger = logging.getLogger("EngineA_PyMuPDF")


class PyMuPDFLayoutEngine:
    """Cross-platform layout engine based on PyMuPDF and pdf2docx."""

    @classmethod
    def convert_to_docx(cls, pdf_path: str, output_docx_path: str) -> bool:
        """Converts PDF to DOCX using pdf2docx or fallback builder."""
        abs_pdf = os.path.abspath(pdf_path)
        abs_docx = os.path.abspath(output_docx_path)
        os.makedirs(os.path.dirname(abs_docx), exist_ok=True)

        # 1. Try pdf2docx first for rich baseline DOM
        try:
            from pdf2docx import Converter
            cv = Converter(abs_pdf)
            cv.convert(abs_docx)
            cv.close()
            if os.path.exists(abs_docx) and os.path.getsize(abs_docx) > 500:
                logger.info(f"Engine A (pdf2docx) successfully produced baseline: {abs_docx}")
                return True
        except Exception as e:
            logger.warning(f"Engine A pdf2docx baseline failed, falling back to DocxBuilder: {e}")

        # 2. Pure PyMuPDF + DocxBuilder Fallback
        try:
            with PDFInspector(abs_pdf) as inspector:
                pages_data = inspector.extract_all_pages()
                table_recon = TableReconstructor()

                for p_data in pages_data:
                    tables, remaining_lines = table_recon.reconstruct_tables(
                        p_data.lines, p_data.vector_paths, p_data.width, p_data.height
                    )
                    p_data.tables = tables
                    p_data.paragraphs = StyleInferencer.consolidate_paragraphs(
                        remaining_lines, p_data.width
                    )

                builder = DocxBuilder()
                builder.build(pages_data, abs_docx)

                if os.path.exists(abs_docx) and os.path.getsize(abs_docx) > 500:
                    logger.info(f"Engine A (DocxBuilder) produced baseline: {abs_docx}")
                    return True
        except Exception as e:
            logger.error(f"Engine A pure PyMuPDF fallback failed: {e}")
            return False

        return False
