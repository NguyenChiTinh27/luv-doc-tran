"""
xy_cut.py
─────────
Recursive XY-Cut algorithm for spatial reading order detection and multi-column decomposition.
Extracts natural reading flow across multi-column, side-by-side, and complex page layouts.
"""

from __future__ import annotations
from typing import List, Tuple, Any
from src.pdf2word_core.models.schema import BBox


class XYCut:
    """Recursive XY-Cut spatial partitioning algorithm."""

    @classmethod
    def sort_reading_order(cls, items: List[Any], get_bbox_fn) -> List[Any]:
        """Sort items into natural reading order using recursive XY-Cut."""
        if len(items) <= 1:
            return items
        indexed_items = [(i, get_bbox_fn(items[i])) for i in range(len(items))]
        sorted_indices = cls._recursive_xy_cut(indexed_items)
        return [items[i] for i in sorted_indices]

    @classmethod
    def _recursive_xy_cut(cls, indexed_boxes: List[Tuple[int, BBox]]) -> List[int]:
        if len(indexed_boxes) <= 1:
            return [b[0] for b in indexed_boxes]

        # 1. Try horizontal cut first (splitting into vertical bands)
        h_cut = cls._find_horizontal_cut(indexed_boxes)
        if h_cut is not None:
            top_group = [b for b in indexed_boxes if b[1].center_y < h_cut]
            bottom_group = [b for b in indexed_boxes if b[1].center_y >= h_cut]
            if top_group and bottom_group:
                return cls._recursive_xy_cut(top_group) + cls._recursive_xy_cut(bottom_group)

        # 2. Try vertical cut (splitting multi-columns side-by-side)
        v_cut = cls._find_vertical_cut(indexed_boxes)
        if v_cut is not None:
            left_group = [b for b in indexed_boxes if b[1].center_x < v_cut]
            right_group = [b for b in indexed_boxes if b[1].center_x >= v_cut]
            if left_group and right_group:
                return cls._recursive_xy_cut(left_group) + cls._recursive_xy_cut(right_group)

        # Fallback: standard top-to-bottom, left-to-right sort
        sorted_fallback = sorted(
            indexed_boxes,
            key=lambda b: (round(b[1].y0 / 5.0) * 5.0, b[1].x0)
        )
        return [b[0] for b in sorted_fallback]

    @classmethod
    def _find_horizontal_cut(cls, indexed_boxes: List[Tuple[int, BBox]]) -> float | None:
        """Finds horizontal whitespace valley that cleanly divides items."""
        y_events: List[Tuple[float, int]] = []
        for _, box in indexed_boxes:
            y_events.append((box.y0, 1))   # Start of element
            y_events.append((box.y1, -1))  # End of element
        y_events.sort(key=lambda e: e[0])

        active_count = 0
        best_gap = 0.0
        best_cut = None

        for i in range(len(y_events) - 1):
            active_count += y_events[i][1]
            if active_count == 0:
                gap = y_events[i + 1][0] - y_events[i][0]
                if gap > 8.0 and gap > best_gap:
                    best_gap = gap
                    best_cut = (y_events[i][0] + y_events[i + 1][0]) / 2.0

        return best_cut

    @classmethod
    def _find_vertical_cut(cls, indexed_boxes: List[Tuple[int, BBox]]) -> float | None:
        """Finds vertical whitespace valley that cleanly divides multi-column items."""
        x_events: List[Tuple[float, int]] = []
        for _, box in indexed_boxes:
            x_events.append((box.x0, 1))
            x_events.append((box.x1, -1))
        x_events.sort(key=lambda e: e[0])

        active_count = 0
        best_gap = 0.0
        best_cut = None

        for i in range(len(x_events) - 1):
            active_count += x_events[i][1]
            if active_count == 0:
                gap = x_events[i + 1][0] - x_events[i][0]
                if gap > 15.0 and gap > best_gap:
                    best_gap = gap
                    best_cut = (x_events[i][0] + x_events[i + 1][0]) / 2.0

        return best_cut
