"""
table_grid_graph.py
───────────────────
Table grid topology, coordinate snapping, and cell structure reconstruction.
Reconstructs table grids from explicit vector lines and stroke graphs:
- Groups horizontal and vertical lines by spatial adjacency
- Snaps coordinates within tolerance (e.g. 2.0 pt) to form canonical rows and columns
- Maps background shading rects to individual cells
- Associates text blocks contained inside cell bounding boxes
- Detects borders (top, bottom, left, right) for high-fidelity OpenXML rendering
"""

from __future__ import annotations
import math
from typing import List, Tuple, Optional, Dict, Any, Set

from .models import (
    BBox, TextBlock, DrawingData, GridCellData, TableGridData
)


class TableGridGraph:
    """Table grid topology reconstruction and coordinate divider snapping."""

    def __init__(self, snap_tolerance: float = 2.0):
        self.snap_tolerance = snap_tolerance

    def snap_coordinates(self, coords: List[float], tolerance: Optional[float] = None) -> List[float]:
        """Snaps 1D coordinates within tolerance into canonical grid dividers."""
        if not coords:
            return []
        tol = tolerance if tolerance is not None else self.snap_tolerance
        sorted_coords = sorted(coords)
        snapped: List[float] = [sorted_coords[0]]
        for c in sorted_coords[1:]:
            if abs(c - snapped[-1]) > tol:
                snapped.append(c)
            else:
                # Weighted average
                snapped[-1] = (snapped[-1] + c) / 2.0
        return snapped

    def reconstruct_tables(
        self,
        blocks: List[TextBlock],
        drawings: List[DrawingData],
        page_width: float,
        page_height: float
    ) -> Tuple[List[TableGridData], List[TextBlock]]:
        """
        Reconstructs table grid structures from vector strokes and text blocks.
        Returns:
            tables: List of TableGridData
            remaining_blocks: Text blocks that do not belong to any reconstructed table
        """
        h_lines: List[DrawingData] = [d for d in drawings if d.is_horizontal]
        v_lines: List[DrawingData] = [d for d in drawings if d.is_vertical]
        fill_rects: List[DrawingData] = [d for d in drawings if d.is_rect and d.fill_color]

        if len(h_lines) < 2:
            return [], blocks

        # Group horizontal lines by proximity to identify table regions
        h_groups = self._cluster_lines_by_table(h_lines, v_lines)
        tables: List[TableGridData] = []
        assigned_block_indices: Set[int] = set()

        for group_h, group_v in h_groups:
            if len(group_h) < 2:
                continue

            # Extract unique snapped Y coordinates (row dividers)
            raw_ys: List[float] = []
            for p in group_h:
                raw_ys.append(p.bbox[1])
                raw_ys.append(p.bbox[3])
            snapped_ys = self.snap_coordinates(raw_ys, self.snap_tolerance)
            if len(snapped_ys) < 2:
                continue

            # Extract unique snapped X coordinates (column dividers)
            raw_xs: List[float] = []
            for p in group_v:
                raw_xs.append(p.bbox[0])
                raw_xs.append(p.bbox[2])
            for p in group_h:
                raw_xs.append(p.bbox[0])
                raw_xs.append(p.bbox[2])
            snapped_xs = self.snap_coordinates(raw_xs, self.snap_tolerance)
            if len(snapped_xs) < 2:
                continue

            num_rows = len(snapped_ys) - 1
            num_cols = len(snapped_xs) - 1
            if num_rows < 1 or num_cols < 1:
                continue

            col_widths = [snapped_xs[c + 1] - snapped_xs[c] for c in range(num_cols)]
            t_bbox = (snapped_xs[0], snapped_ys[0], snapped_xs[-1], snapped_ys[-1])

            cells: List[GridCellData] = []
            for r in range(num_rows):
                r_top = snapped_ys[r]
                r_bottom = snapped_ys[r + 1]
                for c in range(num_cols):
                    c_left = snapped_xs[c]
                    c_right = snapped_xs[c + 1]
                    cell_bbox = BBox(c_left, r_top, c_right, r_bottom)

                    # Match cell background shading
                    cell_shd = None
                    for fr in fill_rects:
                        fr_box = BBox.from_tuple(fr.bbox)
                        if cell_bbox.intersects(fr_box, tolerance=1.0) and fr_box.area > 10.0:
                            cell_shd = fr.fill_color
                            break

                    # Match text blocks inside cell
                    cell_blocks: List[TextBlock] = []
                    for idx, blk in enumerate(blocks):
                        if idx in assigned_block_indices:
                            continue
                        blk_box = BBox.from_tuple(blk.bbox)
                        if cell_bbox.contains(blk_box, tolerance=2.0) or (
                            cell_bbox.intersects(blk_box) and
                            (c_left <= blk_box.center_x <= c_right and
                             r_top <= blk_box.center_y <= r_bottom)
                        ):
                            cell_blocks.append(blk)
                            assigned_block_indices.add(idx)

                    # Detect cell borders from group_h and group_v
                    borders = self._detect_cell_borders(cell_bbox, group_h, group_v)

                    cell_text = "\n".join(b.text for b in cell_blocks)
                    cells.append(GridCellData(
                        row_idx=r,
                        col_idx=c,
                        row_span=1,
                        col_span=1,
                        bbox=cell_bbox.to_tuple(),
                        text=cell_text,
                        blocks=cell_blocks,
                        bg_color=cell_shd,
                        borders=borders
                    ))

            # Only add table if it contains text or has multiple rows
            if any(cell.blocks for cell in cells) or num_rows >= 2:
                tables.append(TableGridData(
                    rows=num_rows,
                    cols=num_cols,
                    col_widths=col_widths,
                    cells=cells,
                    bbox=t_bbox,
                    is_bordered=True
                ))

        remaining_blocks = [
            b for idx, b in enumerate(blocks)
            if idx not in assigned_block_indices
        ]
        return tables, remaining_blocks

    def _cluster_lines_by_table(
        self,
        h_lines: List[DrawingData],
        v_lines: List[DrawingData]
    ) -> List[Tuple[List[DrawingData], List[DrawingData]]]:
        """Groups horizontal and vertical lines that form contiguous table grids."""
        sorted_h = sorted(h_lines, key=lambda l: l.bbox[1])
        groups: List[Tuple[List[DrawingData], List[DrawingData]]] = []

        current_h: List[DrawingData] = []
        for line in sorted_h:
            if not current_h:
                current_h.append(line)
            else:
                prev_y = current_h[-1].bbox[1]
                # If distance to previous horizontal line is reasonable (< 120 pt), group together
                if (line.bbox[1] - prev_y) <= 120.0:
                    current_h.append(line)
                else:
                    min_y = min(h.bbox[1] for h in current_h) - self.snap_tolerance
                    max_y = max(h.bbox[3] for h in current_h) + self.snap_tolerance
                    matched_v = [
                        v for v in v_lines
                        if not (v.bbox[3] < min_y or v.bbox[1] > max_y)
                    ]
                    groups.append((current_h, matched_v))
                    current_h = [line]

        if current_h:
            min_y = min(h.bbox[1] for h in current_h) - self.snap_tolerance
            max_y = max(h.bbox[3] for h in current_h) + self.snap_tolerance
            matched_v = [
                v for v in v_lines
                if not (v.bbox[3] < min_y or v.bbox[1] > max_y)
            ]
            groups.append((current_h, matched_v))

        return groups

    def _detect_cell_borders(
        self,
        cell_bbox: BBox,
        h_lines: List[DrawingData],
        v_lines: List[DrawingData]
    ) -> Dict[str, Dict[str, Any]]:
        """Determines border styles for each edge of a cell."""
        tol = self.snap_tolerance
        borders: Dict[str, Dict[str, Any]] = {
            "top": {"val": "none", "sz": 0, "color": "auto"},
            "bottom": {"val": "none", "sz": 0, "color": "auto"},
            "left": {"val": "none", "sz": 0, "color": "auto"},
            "right": {"val": "none", "sz": 0, "color": "auto"},
        }

        for h in h_lines:
            hy = (h.bbox[1] + h.bbox[3]) / 2.0
            hx0, hx1 = h.bbox[0], h.bbox[2]
            # Check overlap along X
            if max(hx0, cell_bbox.x0) < min(hx1, cell_bbox.x1) - tol:
                if abs(hy - cell_bbox.y0) <= tol:
                    borders["top"] = {"val": "single", "sz": 4, "color": h.stroke_color or "000000"}
                if abs(hy - cell_bbox.y1) <= tol:
                    borders["bottom"] = {"val": "single", "sz": 4, "color": h.stroke_color or "000000"}

        for v in v_lines:
            vx = (v.bbox[0] + v.bbox[2]) / 2.0
            vy0, vy1 = v.bbox[1], v.bbox[3]
            # Check overlap along Y
            if max(vy0, cell_bbox.y0) < min(vy1, cell_bbox.y1) - tol:
                if abs(vx - cell_bbox.x0) <= tol:
                    borders["left"] = {"val": "single", "sz": 4, "color": v.stroke_color or "000000"}
                if abs(vx - cell_bbox.x1) <= tol:
                    borders["right"] = {"val": "single", "sz": 4, "color": v.stroke_color or "000000"}

        return borders
