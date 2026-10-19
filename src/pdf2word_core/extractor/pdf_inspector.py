"""
pdf_inspector.py
────────────────
Deep inspection and extraction of PDF primitives using PyMuPDF (fitz).
Extracts text spans, fonts, colors, geometric vector paths, raster images,
and detects dense vector clusters (logos, GHS pictograms, stamps).
"""

from __future__ import annotations
import os
import re
import math
import logging
from typing import List, Tuple, Optional, Dict, Any

import fitz  # PyMuPDF
from src.pdf2word_core.models.schema import (
    BBox, SpanData, LineData, VectorPathData, VectorCluster,
    ImageData, PageData
)

logger = logging.getLogger("PDFInspector")


def int_to_hex_color(color_int: Optional[int]) -> str:
    """Convert an RGB integer color to a 6-character hex string."""
    if color_int is None or color_int < 0:
        return "000000"
    r = (color_int >> 16) & 255
    g = (color_int >> 8) & 255
    b = color_int & 255
    return f"{r:02X}{g:02X}{b:02X}"


def rgb_tuple_to_hex(rgb: Optional[Tuple[float, ...]]) -> Optional[str]:
    """Convert float RGB tuple (0.0 - 1.0) to hex string."""
    if not rgb:
        return None
    if len(rgb) >= 3:
        r = int(min(255, max(0, round(rgb[0] * 255))))
        g = int(min(255, max(0, round(rgb[1] * 255))))
        b = int(min(255, max(0, round(rgb[2] * 255))))
        return f"{r:02X}{g:02X}{b:02X}"
    elif len(rgb) == 1:
        # Grayscale
        val = int(min(255, max(0, round(rgb[0] * 255))))
        return f"{val:02X}{val:02X}{val:02X}"
    return None


class PDFInspector:
    """Extracts raw ground-truth geometry, typography, and vectors from PDF."""

    def __init__(self, pdf_path: str):
        self.pdf_path = os.path.abspath(pdf_path)
        if not os.path.exists(self.pdf_path):
            raise FileNotFoundError(f"PDF file not found: {self.pdf_path}")
        self.doc = fitz.open(self.pdf_path)
        self.num_pages = len(self.doc)

    def close(self):
        if self.doc:
            self.doc.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    @staticmethod
    def strip_pdf_restrictions(input_path: str, output_path: str) -> bool:
        """
        Creates an unrestricted copy of the PDF with all DRM permissions allowed.
        Prevents Word COM from throwing permission exceptions (-2147352567).
        """
        try:
            with fitz.open(input_path) as doc:
                doc.save(output_path, encryption=fitz.PDF_ENCRYPT_NONE)
            return True
        except Exception as e:
            logger.warning(f"Could not strip restrictions on {input_path}: {e}")
            return False

    def inspect_page(self, page_index: int) -> PageData:
        """Extract all primitives for a single page."""
        page = self.doc[page_index]
        rect = page.rect
        page_width = rect.width
        page_height = rect.height

        lines = self._extract_text_lines(page)
        vector_paths, vector_clusters = self._extract_vectors_and_clusters(page)
        images = self._extract_images(page)

        # Estimate margins based on bounding box of elements
        margins = self._calculate_page_margins(lines, vector_paths, page_width, page_height)

        return PageData(
            page_num=page_index + 1,
            width=page_width,
            height=page_height,
            margins=margins,
            lines=lines,
            vector_paths=vector_paths,
            vector_clusters=vector_clusters,
            images=images
        )

    def _extract_text_lines(self, page: fitz.Page) -> List[LineData]:
        """Extract structured text lines with span typography."""
        lines_out: List[LineData] = []
        page_dict = page.get_text("dict")

        for block in page_dict.get("blocks", []):
            if block.get("type") != 0:  # 0 is text block
                continue
            for line in block.get("lines", []):
                spans: List[SpanData] = []
                for s in line.get("spans", []):
                    text = s.get("text", "")
                    if not text:
                        continue
                    b = s.get("bbox", (0, 0, 0, 0))
                    bbox = BBox(b[0], b[1], b[2], b[3])
                    font_name = s.get("font", "Arial")
                    font_size = float(s.get("size", 10.0))
                    flags = int(s.get("flags", 0))
                    color_hex = int_to_hex_color(s.get("color"))
                    origin = s.get("origin", (b[0], b[3]))

                    spans.append(SpanData(
                        text=text,
                        bbox=bbox,
                        font_name=font_name,
                        font_size=font_size,
                        flags=flags,
                        color_hex=color_hex,
                        origin=origin
                    ))

                if spans:
                    ld = LineData(spans=spans)
                    ld.update_text_and_bbox()
                    # Check for list bullet
                    stripped = ld.text.strip()
                    if stripped and stripped[0] in {'•', '●', '▪', '▫', '-', '+', '*', '–', '—', '➢', '❖', '✓', '✔'}:
                        ld.is_bullet = True
                        ld.bullet_char = stripped[0]
                    lines_out.append(ld)

        # Sort lines in top-to-bottom reading order
        lines_out.sort(key=lambda l: (round(l.bbox.y0 / 3.0) * 3.0, l.bbox.x0))
        return lines_out

    def _extract_vectors_and_clusters(
        self, page: fitz.Page
    ) -> Tuple[List[VectorPathData], List[VectorCluster]]:
        """
        Classifies vector drawings into lines/rectangles and detects dense vector clusters
        (>30 path strokes in a compact region, e.g. logos, stamps, GHS pictograms).
        """
        drawings = page.get_drawings()
        vector_paths: List[VectorPathData] = []
        complex_strokes_bbox: List[BBox] = []

        for d in drawings:
            rect = d.get("rect")
            if not rect:
                continue
            bbox = BBox(rect[0], rect[1], rect[2], rect[3])
            stroke_hex = rgb_tuple_to_hex(d.get("color"))
            fill_hex = rgb_tuple_to_hex(d.get("fill"))
            raw_w = d.get("width")
            width = float(raw_w) if raw_w is not None else 1.0
            items = d.get("items", [])

            # Check if this drawing represents a horizontal or vertical line
            is_horiz = (bbox.height <= 2.5 and bbox.width >= 3.0)
            is_vert = (bbox.width <= 2.5 and bbox.height >= 3.0)
            is_rect = bool(d.get("fill")) and not (is_horiz or is_vert)
            is_curve = any(it[0] in ('c', 'v', 'y') for it in items)

            if is_curve or len(items) > 5:
                complex_strokes_bbox.append(bbox)

            vector_paths.append(VectorPathData(
                bbox=bbox,
                stroke_color=stroke_hex,
                fill_color=fill_hex,
                line_width=width,
                is_horizontal=is_horiz,
                is_vertical=is_vert,
                is_rect=is_rect,
                is_curve=is_curve
            ))

        # Detect vector clusters (>30 complex stroke components clustered together)
        clusters: List[VectorCluster] = []
        if complex_strokes_bbox:
            merged_boxes = self._cluster_bboxes(complex_strokes_bbox, distance_threshold=20.0)
            for m_bbox, count in merged_boxes:
                # Dense cluster threshold: >= 15 curves/strokes within a logo/stamp sized region (< 200 pt)
                if count >= 15 and m_bbox.width < 250 and m_bbox.height < 250:
                    clusters.append(VectorCluster(
                        bbox=m_bbox,
                        path_count=count,
                        is_dense=True
                    ))

        return vector_paths, clusters

    def _cluster_bboxes(
        self, bboxes: List[BBox], distance_threshold: float = 20.0
    ) -> List[Tuple[BBox, int]]:
        """Greedy agglomerative clustering for bounding boxes."""
        if not bboxes:
            return []
        clusters: List[List[BBox]] = []
        for box in bboxes:
            matched = False
            for c in clusters:
                # Check proximity to any box in cluster
                c_union = c[0]
                for cb in c[1:]:
                    c_union = c_union.union(cb)
                if c_union.intersects(box, tolerance=distance_threshold):
                    c.append(box)
                    matched = True
                    break
            if not matched:
                clusters.append([box])

        result: List[Tuple[BBox, int]] = []
        for c in clusters:
            u = c[0]
            for cb in c[1:]:
                u = u.union(cb)
            result.append((u, len(c)))
        return result

    def _extract_images(self, page: fitz.Page) -> List[ImageData]:
        """Extract raster images with exact bounding boxes."""
        images_out: List[ImageData] = []
        image_list = page.get_images(full=True)
        for img_info in image_list:
            xref = img_info[0]
            # Try to get image bounding box on the page
            rects = page.get_image_rects(xref)
            if not rects:
                continue
            base_image = self.doc.extract_image(xref)
            image_bytes = base_image.get("image", b"")
            ext = base_image.get("ext", "png")
            width = base_image.get("width", 0)
            height = base_image.get("height", 0)

            for r in rects:
                images_out.append(ImageData(
                    bbox=BBox(r.x0, r.y0, r.x1, r.y1),
                    image_bytes=image_bytes,
                    ext=ext,
                    width=width,
                    height=height
                ))
        return images_out

    def _calculate_page_margins(
        self,
        lines: List[LineData],
        vectors: List[VectorPathData],
        page_width: float,
        page_height: float
    ) -> Tuple[float, float, float, float]:
        """Calculate printable page margins (top, bottom, left, right in pt)."""
        if not lines and not vectors:
            return (36.0, 36.0, 36.0, 36.0)

        min_x = page_width
        min_y = page_height
        max_x = 0.0
        max_y = 0.0

        for line in lines:
            min_x = min(min_x, line.bbox.x0)
            min_y = min(min_y, line.bbox.y0)
            max_x = max(max_x, line.bbox.x1)
            max_y = max(max_y, line.bbox.y1)

        for v in vectors:
            if v.is_horizontal or v.is_vertical:
                min_x = min(min_x, v.bbox.x0)
                min_y = min(min_y, v.bbox.y0)
                max_x = max(max_x, v.bbox.x1)
                max_y = max(max_y, v.bbox.y1)

        top_margin = max(18.0, min(72.0, min_y))
        bottom_margin = max(18.0, min(72.0, page_height - max_y))
        left_margin = max(18.0, min(72.0, min_x))
        right_margin = max(18.0, min(72.0, page_width - max_x))

        return (top_margin, bottom_margin, left_margin, right_margin)

    def extract_all_pages(self) -> List[PageData]:
        """Extract data across all pages."""
        return [self.inspect_page(i) for i in range(self.num_pages)]
