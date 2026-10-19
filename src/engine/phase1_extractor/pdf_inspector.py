"""
pdf_inspector.py
────────────────
Deep inspection and extraction of PDF ground-truth primitives using PyMuPDF rawdict.
Extracts:
- Text spans, bounding boxes, font attributes (name, size, weight, italic, color, rotation)
- Groups spans into coherent text lines/blocks with natural reading order
- Vector drawings, line grids, and stroke classifications
- Tables with snapped grid coordinates and cell text associations
- Embedded raster images with DPI and bounding boxes
- DRM restriction stripping to guarantee unrestricted downstream processing
"""

from __future__ import annotations
import os
import re
import math
import logging
from typing import List, Tuple, Optional, Dict, Any

import fitz  # PyMuPDF

from .models import (
    BBox, TextSpan, TextBlock, GridCellData, TableGridData,
    DrawingData, ImageData, ExtractedPage, ExtractedDocument
)
from .vector_extractor import VectorExtractor
from .table_grid_graph import TableGridGraph
from .image_clipper import ImageClipper

logger = logging.getLogger("PDFInspector")


def int_to_hex_color(color_int: Optional[int]) -> str:
    """Convert an RGB integer color to a 6-character hex string RRGGBB."""
    if color_int is None or color_int < 0:
        return "000000"
    r = (color_int >> 16) & 255
    g = (color_int >> 8) & 255
    b = color_int & 255
    return f"{r:02X}{g:02X}{b:02X}"


class PDFInspector:
    """Extracts raw ground-truth geometry, typography, and vectors from a PDF."""

    def __init__(
        self,
        pdf_path: Optional[str] = None,
        doc: Optional[fitz.Document] = None,
        snap_tolerance: float = 2.0,
        raster_dpi: int = 300
    ):
        self.pdf_path = os.path.abspath(pdf_path) if pdf_path else ""
        self._external_doc = doc is not None
        if doc is not None:
            self.doc = doc
        elif pdf_path:
            if not os.path.exists(self.pdf_path):
                raise FileNotFoundError(f"PDF file not found: {self.pdf_path}")
            self.doc = fitz.open(self.pdf_path)
        else:
            raise ValueError("Either pdf_path or doc must be provided to PDFInspector.")

        self.num_pages = len(self.doc)
        self.vector_extractor = VectorExtractor(snap_tolerance=snap_tolerance)
        self.table_grid_graph = TableGridGraph(snap_tolerance=snap_tolerance)
        self.image_clipper = ImageClipper(high_dpi=raster_dpi)

    def close(self):
        if not self._external_doc and self.doc:
            self.doc.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    @staticmethod
    def strip_pdf_restrictions(input_path: str, output_path: str) -> bool:
        """
        Creates an unrestricted copy of the PDF with all DRM permissions allowed.
        Prevents Word COM and reader processes from throwing permission exceptions.
        """
        try:
            with fitz.open(input_path) as d:
                d.save(output_path, encryption=fitz.PDF_ENCRYPT_NONE)
            return True
        except Exception as e:
            logger.warning(f"Could not strip restrictions on {input_path}: {e}")
            return False

    def inspect_page(self, page_index: int) -> ExtractedPage:
        """Extract all geometric, textual, vector, and image primitives for a page."""
        page = self.doc[page_index]
        rect = page.rect
        page_width = float(rect.width)
        page_height = float(rect.height)
        rotation = float(page.rotation)

        # 1. Text extraction via rawdict / dict
        raw_blocks = self._extract_text_blocks(page)

        # 2. Vector extraction & stroke classification
        drawings, dense_clusters = self.vector_extractor.extract_drawings(page)

        # 3. Table grid reconstruction & line assignment
        tables, remaining_blocks = self.table_grid_graph.reconstruct_tables(
            raw_blocks, drawings, page_width, page_height
        )

        # 4. Raster images & vector cluster clippings
        images = self.image_clipper.extract_page_images(page, self.doc)
        if dense_clusters:
            cluster_images = self.image_clipper.rasterize_vector_clusters(page, dense_clusters)
            images.extend(cluster_images)

        # 5. Margins estimation
        margins = self._calculate_page_margins(
            remaining_blocks, drawings, page_width, page_height
        )

        return ExtractedPage(
            page_num=page_index + 1,
            width=page_width,
            height=page_height,
            margins=margins,
            blocks=remaining_blocks,
            tables=tables,
            drawings=drawings,
            images=images,
            rotation=rotation
        )

    def extract_document(self) -> ExtractedDocument:
        """Extract all pages into a consolidated ExtractedDocument."""
        pages = [self.inspect_page(i) for i in range(self.num_pages)]
        metadata = dict(self.doc.metadata) if self.doc.metadata else {}
        return ExtractedDocument(
            pages=pages,
            pdf_path=self.pdf_path,
            metadata=metadata
        )

    def _extract_text_blocks(self, page: fitz.Page) -> List[TextBlock]:
        """Extract structured text lines with span typography from PyMuPDF dict."""
        blocks_out: List[TextBlock] = []
        try:
            # Use rawdict when available for granular sub-span detail, fallback to dict
            page_dict = page.get_text("rawdict")
        except Exception:
            page_dict = page.get_text("dict")

        line_counter = 0
        for block in page_dict.get("blocks", []):
            if block.get("type") != 0:  # 0 is text block
                continue
            for line in block.get("lines", []):
                spans: List[TextSpan] = []
                for s in line.get("spans", []):
                    text = s.get("text", "")
                    if not text and "chars" in s:
                        text = "".join(ch.get("c", "") for ch in s.get("chars", []))
                    if not text:
                        continue
                    b = s.get("bbox", (0, 0, 0, 0))
                    bbox = (float(b[0]), float(b[1]), float(b[2]), float(b[3]))
                    font_name = s.get("font", "Arial")
                    font_size = float(s.get("size", 10.0))
                    flags = int(s.get("flags", 0))
                    color_hex = int_to_hex_color(s.get("color"))
                    raw_origin = s.get("origin", (b[0], b[3]))
                    origin = (float(raw_origin[0]), float(raw_origin[1]))
                    rot = float(line.get("dir", (1.0, 0.0))[1])  # angle or direction

                    spans.append(TextSpan(
                        text=text,
                        bbox=bbox,
                        font_name=font_name,
                        font_size=font_size,
                        flags=flags,
                        color_hex=color_hex,
                        origin=origin,
                        rotation=rot
                    ))

                if spans:
                    tb = TextBlock(spans=spans, line_index=line_counter)
                    tb.update_from_spans()
                    blocks_out.append(tb)
                    line_counter += 1

        # Sort blocks in reading order: top-to-bottom, left-to-right (3 pt vertical bucket)
        blocks_out.sort(key=lambda b: (round(b.bbox[1] / 3.0) * 3.0, b.bbox[0]))
        return blocks_out

    def _calculate_page_margins(
        self,
        blocks: List[TextBlock],
        drawings: List[DrawingData],
        page_width: float,
        page_height: float
    ) -> Tuple[float, float, float, float]:
        """Estimate printable page margins (top, bottom, left, right in pt)."""
        if not blocks and not drawings:
            return (36.0, 36.0, 36.0, 36.0)

        min_x = page_width
        min_y = page_height
        max_x = 0.0
        max_y = 0.0

        for b in blocks:
            min_x = min(min_x, b.bbox[0])
            min_y = min(min_y, b.bbox[1])
            max_x = max(max_x, b.bbox[2])
            max_y = max(max_y, b.bbox[3])

        for d in drawings:
            if d.is_horizontal or d.is_vertical:
                min_x = min(min_x, d.bbox[0])
                min_y = min(min_y, d.bbox[1])
                max_x = max(max_x, d.bbox[2])
                max_y = max(max_y, d.bbox[3])

        top_margin = max(18.0, min(72.0, min_y))
        bottom_margin = max(18.0, min(72.0, page_height - max_y))
        left_margin = max(18.0, min(72.0, min_x))
        right_margin = max(18.0, min(72.0, page_width - max_x))

        return (top_margin, bottom_margin, left_margin, right_margin)
