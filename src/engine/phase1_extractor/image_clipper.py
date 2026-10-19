"""
image_clipper.py
────────────────
Raster image extraction, clipping, and high-DPI preservation.
- Extracts embedded raster images with exact page bounding boxes and dimensions
- Rasterizes dense vector clusters or complex graphics at 300 DPI with alpha transparency
- Preserves DPI metadata and format attributes (PNG, JPEG, etc.)
"""

from __future__ import annotations
import logging
from typing import List, Tuple, Optional, Dict, Any
import fitz  # PyMuPDF

from .models import BBox, ImageData, DrawingData

logger = logging.getLogger("ImageClipper")


class ImageClipper:
    """Extracts raster images and rasterizes high-DPI vector clusters."""

    def __init__(self, target_dpi: int = 150, high_dpi: int = 300):
        self.target_dpi = target_dpi
        self.high_dpi = high_dpi

    def extract_page_images(
        self, page: fitz.Page, doc: fitz.Document
    ) -> List[ImageData]:
        """
        Extracts embedded raster images on the page along with their bounding boxes.
        """
        images_out: List[ImageData] = []
        try:
            image_list = page.get_images(full=True)
        except Exception as e:
            logger.warning(f"Failed to get_images on page: {e}")
            return []

        for img_info in image_list:
            xref = img_info[0]
            try:
                rects = page.get_image_rects(xref)
            except Exception:
                rects = []
            if not rects:
                continue

            try:
                base_image = doc.extract_image(xref)
            except Exception as e:
                logger.warning(f"Failed to extract image xref={xref}: {e}")
                continue

            image_bytes = base_image.get("image", b"")
            ext = base_image.get("ext", "png")
            width = base_image.get("width", 0)
            height = base_image.get("height", 0)
            dpi = base_image.get("xres", self.target_dpi) or self.target_dpi

            for r in rects:
                images_out.append(ImageData(
                    bbox=(r.x0, r.y0, r.x1, r.y1),
                    image_bytes=image_bytes,
                    ext=ext,
                    width=width,
                    height=height,
                    dpi=int(dpi)
                ))

        return images_out

    def clip_region_raster(
        self,
        page: fitz.Page,
        bbox: Tuple[float, float, float, float],
        dpi: Optional[int] = None
    ) -> bytes:
        """
        Renders a bounding box region of a PDF page at specified DPI with alpha transparency.
        """
        use_dpi = dpi if dpi is not None else self.high_dpi
        clip_rect = fitz.Rect(bbox[0], bbox[1], bbox[2], bbox[3])
        # Render with alpha channel enabled
        pix = page.get_pixmap(clip=clip_rect, dpi=use_dpi, alpha=True)
        return pix.tobytes("png")

    def rasterize_vector_clusters(
        self,
        page: fitz.Page,
        clusters: List[DrawingData],
        dpi: Optional[int] = None
    ) -> List[ImageData]:
        """
        Converts dense vector clusters (logos, stamps, hazard diamonds) into high-resolution PNG ImageData.
        """
        use_dpi = dpi if dpi is not None else self.high_dpi
        out_images: List[ImageData] = []

        for cluster in clusters:
            bbox = cluster.bbox
            b = BBox.from_tuple(bbox)
            # Add 1 pt margin for anti-aliasing bounds
            clip_rect = fitz.Rect(b.x0 - 1.0, b.y0 - 1.0, b.x1 + 1.0, b.y1 + 1.0)
            try:
                pix = page.get_pixmap(clip=clip_rect, dpi=use_dpi, alpha=True)
                png_bytes = pix.tobytes("png")
                out_images.append(ImageData(
                    bbox=(b.x0, b.y0, b.x1, b.y1),
                    image_bytes=png_bytes,
                    ext="png",
                    width=pix.width,
                    height=pix.height,
                    dpi=use_dpi
                ))
            except Exception as e:
                logger.warning(f"Failed to rasterize cluster {bbox}: {e}")

        return out_images
