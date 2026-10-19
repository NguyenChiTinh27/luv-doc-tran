"""
vector_extractor.py
───────────────────
Vector graphics, drawings, lines, strokes, and dense cluster detection.
Extracts vector paths from PyMuPDF, classifies strokes (horizontal/vertical lines,
fills, curves), groups lines for table grid topology, and detects dense vector
clusters (logos, stamps, GHS hazard pictograms).
"""

from __future__ import annotations
import math
import logging
from typing import List, Tuple, Optional, Dict, Any, Set
import fitz  # PyMuPDF

from .models import BBox, DrawingData

logger = logging.getLogger("VectorExtractor")


def rgb_tuple_to_hex(rgb: Optional[Tuple[float, ...]]) -> Optional[str]:
    """Convert float RGB tuple (0.0 - 1.0) to hex string RRGGBB."""
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


class VectorExtractor:
    """Extracts, classifies, and clusters vector graphics from a PDF page."""

    def __init__(self, snap_tolerance: float = 2.0, cluster_distance_threshold: float = 20.0):
        self.snap_tolerance = snap_tolerance
        self.cluster_distance_threshold = cluster_distance_threshold

    def extract_drawings(
        self, page: fitz.Page
    ) -> Tuple[List[DrawingData], List[DrawingData]]:
        """
        Extracts all vector paths from the page.
        Returns:
            all_drawings: List of all classified DrawingData
            dense_clusters: List of aggregate DrawingData representing dense clusters (logos, GHS pictograms)
        """
        raw_drawings = page.get_drawings()
        all_drawings: List[DrawingData] = []
        complex_boxes: List[BBox] = []

        for d in raw_drawings:
            rect = d.get("rect")
            if not rect:
                continue
            x0, y0, x1, y1 = float(rect[0]), float(rect[1]), float(rect[2]), float(rect[3])
            bbox = BBox(x0, y0, x1, y1)

            stroke_hex = rgb_tuple_to_hex(d.get("color"))
            fill_hex = rgb_tuple_to_hex(d.get("fill"))
            raw_w = d.get("width")
            line_w = float(raw_w) if raw_w is not None else 1.0
            items = d.get("items", [])

            # Geometric line classification
            is_horiz = (bbox.height <= 2.5 and bbox.width >= 3.0)
            is_vert = (bbox.width <= 2.5 and bbox.height >= 3.0)
            is_rect = bool(d.get("fill")) and not (is_horiz or is_vert)
            is_curve = any(it[0] in ('c', 'v', 'y', 're', 'qu') for it in items)

            if is_curve or len(items) > 5:
                complex_boxes.append(bbox)

            all_drawings.append(DrawingData(
                bbox=(bbox.x0, bbox.y0, bbox.x1, bbox.y1),
                stroke_color=stroke_hex,
                fill_color=fill_hex,
                line_width=line_w,
                is_horizontal=is_horiz,
                is_vertical=is_vert,
                is_rect=is_rect,
                is_curve=is_curve,
                items=items
            ))

        # Cluster complex curves (logos, GHS hazard diamonds, stamps)
        dense_clusters: List[DrawingData] = []
        if complex_boxes:
            clustered = self._cluster_bboxes(complex_boxes, self.cluster_distance_threshold)
            for m_box, count in clustered:
                if count >= 15 and m_box.width < 250.0 and m_box.height < 250.0:
                    dense_clusters.append(DrawingData(
                        bbox=(m_box.x0, m_box.y0, m_box.x1, m_box.y1),
                        stroke_color="000000",
                        fill_color=None,
                        line_width=1.0,
                        is_curve=True,
                        items=[("cluster", count)]
                    ))

        return all_drawings, dense_clusters

    def _cluster_bboxes(
        self, bboxes: List[BBox], distance_threshold: float
    ) -> List[Tuple[BBox, int]]:
        """Greedy agglomerative clustering for bounding boxes."""
        if not bboxes:
            return []
        clusters: List[List[BBox]] = []
        for box in bboxes:
            matched = False
            for c in clusters:
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

    def build_stroke_graph(self, drawings: List[DrawingData]) -> Dict[str, Any]:
        """
        Constructs a topology graph from stroke lines:
        returns horizontal lines, vertical lines, fill rects, and intersection points.
        """
        h_lines = [d for d in drawings if d.is_horizontal]
        v_lines = [d for d in drawings if d.is_vertical]
        fills = [d for d in drawings if d.is_rect and d.fill_color]

        intersections: List[Tuple[float, float]] = []
        tol = self.snap_tolerance
        for h in h_lines:
            hy = (h.bbox[1] + h.bbox[3]) / 2.0
            hx0, hx1 = h.bbox[0], h.bbox[2]
            for v in v_lines:
                vx = (v.bbox[0] + v.bbox[2]) / 2.0
                vy0, vy1 = v.bbox[1], v.bbox[3]
                if (hx0 - tol <= vx <= hx1 + tol) and (vy0 - tol <= hy <= vy1 + tol):
                    intersections.append((vx, hy))

        return {
            "h_lines": h_lines,
            "v_lines": v_lines,
            "fill_rects": fills,
            "intersections": intersections
        }
