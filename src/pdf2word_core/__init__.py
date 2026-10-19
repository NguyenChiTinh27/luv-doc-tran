"""
pdf2word_core
─────────────
Production-grade Hybrid Geometry-to-OpenXML PDF-to-DOCX Core Conversion Engine.
"""

from src.pdf2word_core.models.schema import ConversionConfig, ConversionReport
from src.pdf2word_core.converter import PDF2WordConverter
from src.pdf2word_core.extractor.pdf_inspector import PDFInspector
from src.pdf2word_core.renderer.oxml_surgeon import OxmlSurgeon
from src.pdf2word_core.verifier.visual_qa import VisualQALoop

__all__ = [
    "ConversionConfig",
    "ConversionReport",
    "PDF2WordConverter",
    "PDFInspector",
    "OxmlSurgeon",
    "VisualQALoop"
]
