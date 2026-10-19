"""
vector_rasterizer.py
────────────────────
High-DPI transparent rasterizer for dense vector clusters (logos, stamps,
GHS pictograms, chemical structure diagrams) that cannot be translated into Word tables.
"""

from __future__ import annotations
import fitz  # PyMuPDF
from typing import Optional
from src.pdf2word_core.models.schema import BBox, VectorCluster


class VectorRasterizer:
    """Rasterizes dense vector regions into transparent high-resolution PNGs."""

    def __init__(self, target_dpi: int = 300):
        self.target_dpi = target_dpi

    def rasterize_cluster(self, page: fitz.Page, cluster: VectorCluster) -> bytes:
        """
        Renders the bounding box region of the page at 300 DPI with alpha transparency.
        """
        bbox = cluster.bbox
        clip_rect = fitz.Rect(bbox.x0 - 1.0, bbox.y0 - 1.0, bbox.x1 + 1.0, bbox.y1 + 1.0)
        # Render with alpha channel enabled
        pix = page.get_pixmap(clip=clip_rect, dpi=self.target_dpi, alpha=True)
        img_bytes = pix.tobytes("png")
        cluster.rendered_png = img_bytes
        return img_bytes

    def rasterize_bbox(self, page: fitz.Page, bbox: BBox) -> bytes:
        """Rasterize arbitrary bbox with alpha channel."""
        clip_rect = fitz.Rect(bbox.x0, bbox.y0, bbox.x1, bbox.y1)
        pix = page.get_pixmap(clip=clip_rect, dpi=self.target_dpi, alpha=True)
        return pix.tobytes("png")
