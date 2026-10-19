"""
src.engine.phase1_extractor
───────────────────────────
Phase 1: PDF Extractor Engine
Extracts ground-truth geometry, typography, vector drawings, tables, and raster images from PDF.
"""

from .models import (
    BBox,
    TextSpan,
    TextBlock,
    GridCellData,
    TableGridData,
    DrawingData,
    ImageData,
    ExtractedPage,
    ExtractedDocument,
)
from .pdf_inspector import PDFInspector, int_to_hex_color
from .vector_extractor import VectorExtractor, rgb_tuple_to_hex
from .table_grid_graph import TableGridGraph
from .image_clipper import ImageClipper

__all__ = [
    "BBox",
    "TextSpan",
    "TextBlock",
    "GridCellData",
    "TableGridData",
    "DrawingData",
    "ImageData",
    "ExtractedPage",
    "ExtractedDocument",
    "PDFInspector",
    "VectorExtractor",
    "TableGridGraph",
    "ImageClipper",
    "int_to_hex_color",
    "rgb_tuple_to_hex",
]
