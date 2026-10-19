"""
engine_b_word_com.py
────────────────────
Engine B: Windows MS Word Native PDF Reflow Engine.
Uses Microsoft Word's built-in PDF importer via COM automation to generate
the initial high-fidelity baseline document.
Includes automatic DRM permission stripping via PyMuPDF to prevent COM error -2147352567.
"""

from __future__ import annotations
import os
import sys
import time
import tempfile
import logging
from typing import Optional

from src.pdf2word_core.extractor.pdf_inspector import PDFInspector

logger = logging.getLogger("EngineB_WordCOM")


class WordCOMEngine:
    """Automates Microsoft Word COM native PDF reflow import."""

    @staticmethod
    def is_available() -> bool:
        """Checks if running on Windows and Word COM is accessible."""
        if sys.platform != "win32":
            return False
        try:
            import win32com.client
            # Test dispatch
            word = win32com.client.DispatchEx("Word.Application")
            word.Quit()
            return True
        except Exception as e:
            logger.debug(f"Word COM check failed: {e}")
            return False

    @classmethod
    def convert_to_docx(cls, pdf_path: str, output_docx_path: str) -> bool:
        """
        Converts PDF to DOCX using Word COM native PDF reflow.
        Strips DRM restrictions before calling Word to prevent permissions errors.
        """
        if not cls.is_available():
            logger.warning("Word COM is not available on this platform.")
            return False

        import win32com.client

        abs_pdf = os.path.abspath(pdf_path)
        abs_docx = os.path.abspath(output_docx_path)
        os.makedirs(os.path.dirname(abs_docx), exist_ok=True)

        # Create temporary unrestricted PDF
        with tempfile.NamedTemporaryFile(suffix="_unrestricted.pdf", delete=False) as tmp:
            unrestricted_pdf = tmp.name

        try:
            # Strip restrictions
            PDFInspector.strip_pdf_restrictions(abs_pdf, unrestricted_pdf)

            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0  # wdAlertsNone

            try:
                # Open PDF with native reflow
                doc = word.Documents.Open(
                    unrestricted_pdf,
                    ConfirmConversions=False,
                    ReadOnly=True,
                    AddToRecentFiles=False
                )
                # wdFormatXMLDocument = 16
                doc.SaveAs2(abs_docx, FileFormat=16)
                doc.Close(SaveChanges=False)
                logger.info(f"Engine B successfully reflowed PDF to DOCX: {abs_docx}")
                return True
            finally:
                try:
                    word.Quit()
                except Exception:
                    pass

        except Exception as e:
            logger.error(f"Engine B Word COM conversion failed: {e}")
            return False
        finally:
            if os.path.exists(unrestricted_pdf):
                try:
                    os.remove(unrestricted_pdf)
                except Exception:
                    pass
