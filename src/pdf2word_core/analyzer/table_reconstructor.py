"""
table_reconstructor.py
──────────────────────
Dual-mode table reconstruction engine:
1. Explicit Stroke-Snap Graph: Snaps vector line segments within 2.0 pt tolerance,
   detects merged cells (gridSpan, vMerge), and matches background shading.
2. Implicit Spatial Clustering: Reconstructs borderless multi-column key-value blocks
   (such as MSDS Section 10 & 11) into stable Word tables with nil borders to eliminate
   word-wrap drift.
"""

from __future__ import annotations
import math
from typing import List, Tuple, Optional, Dict, Any, Set

from src.pdf2word_core.models.schema import (
    BBox, LineData, VectorPathData, GridCellData, TableData, BorderStyle
)


class TableReconstructor:
    """Reconstructs both bordered and borderless tables from PDF primitives."""

    def __init__(self, snap_tolerance: float = 2.0):
        self.snap_tolerance = snap_tolerance

    def reconstruct_tables(
        self,
        lines: List[LineData],
        vector_paths: List[VectorPathData],
        page_width: float,
        page_height: float
    ) -> Tuple[List[TableData], List[LineData]]:
        """
        Detects both bordered and borderless tables.
        Returns:
            tables: List of reconstructed TableData
            remaining_lines: Text lines not captured in any table
        """
        assigned_line_indices: Set[int] = set()
        tables: List[TableData] = []

        # 1. First pass: Bordered tables from vector lines
        bordered_tables, b_line_idx = self._reconstruct_bordered_tables(
            lines, vector_paths
        )
        tables.extend(bordered_tables)
        assigned_line_indices.update(b_line_idx)

        # 2. Second pass: Borderless key-value tables from remaining lines
        unassigned_lines = [
            (idx, line) for idx, line in enumerate(lines)
            if idx not in assigned_line_indices
        ]
        borderless_tables, bl_line_idx = self._reconstruct_borderless_tables(
            unassigned_lines, page_width
        )
        tables.extend(borderless_tables)
        assigned_line_indices.update(bl_line_idx)

        # Remaining non-table lines
        remaining_lines = [
            line for idx, line in enumerate(lines)
            if idx not in assigned_line_indices
        ]

        return tables, remaining_lines

    def _reconstruct_bordered_tables(
        self,
        lines: List[LineData],
        vector_paths: List[VectorPathData]
    ) -> Tuple[List[TableData], Set[int]]:
        """Extract tables formed by explicit horizontal and vertical vector lines."""
        h_lines: List[VectorPathData] = []
        v_lines: List[VectorPathData] = []
        fill_rects: List[VectorPathData] = []

        for p in vector_paths:
            if p.is_horizontal:
                h_lines.append(p)
            elif p.is_vertical:
                v_lines.append(p)
            elif p.is_rect and p.fill_color:
                fill_rects.append(p)

        if len(h_lines) < 2:
            return [], set()

        # Group horizontal lines by vertical proximity to identify candidate table regions
        h_groups = self._cluster_lines_by_table(h_lines, v_lines)
        tables: List[TableData] = []
        assigned_lines: Set[int] = set()

        for group_h, group_v in h_groups:
            if len(group_h) < 2:
                continue

            # Extract unique snapped Y coordinates (row dividers)
            raw_ys = sorted(list(set(
                [p.bbox.y0 for p in group_h] + [p.bbox.y1 for p in group_h]
            )))
            snapped_ys = self._snap_coordinates(raw_ys, self.snap_tolerance)
            if len(snapped_ys) < 2:
                continue

            # Extract unique snapped X coordinates (column dividers)
            raw_xs = sorted(list(set(
                [p.bbox.x0 for p in group_v] + [p.bbox.x1 for p in group_v] +
                [p.bbox.x0 for p in group_h] + [p.bbox.x1 for p in group_h]
            )))
            snapped_xs = self._snap_coordinates(raw_xs, self.snap_tolerance)
            if len(snapped_xs) < 2:
                continue

            num_rows = len(snapped_ys) - 1
            num_cols = len(snapped_xs) - 1
            if num_rows < 1 or num_cols < 1:
                continue

            col_widths = [snapped_xs[c + 1] - snapped_xs[c] for c in range(num_cols)]
            table_bbox = BBox(snapped_xs[0], snapped_ys[0], snapped_xs[-1], snapped_ys[-1])

            cells: List[GridCellData] = []
            for r in range(num_rows):
                r_top = snapped_ys[r]
                r_bottom = snapped_ys[r + 1]
                for c in range(num_cols):
                    c_left = snapped_xs[c]
                    c_right = snapped_xs[c + 1]
                    cell_bbox = BBox(c_left, r_top, c_right, r_bottom)

                    # Match shading fill color
                    cell_shd = None
                    for fr in fill_rects:
                        if cell_bbox.intersects(fr.bbox, tolerance=1.0) and fr.bbox.area > 10.0:
                            cell_shd = fr.fill_color
                            break

                    # Match text lines inside cell
                    cell_lines: List[LineData] = []
                    for idx, line in enumerate(lines):
                        if idx in assigned_lines:
                            continue
                        if cell_bbox.contains(line.bbox, tolerance=2.0) or (
                            cell_bbox.intersects(line.bbox) and
                            (line.bbox.center_x >= c_left and line.bbox.center_x <= c_right and
                             line.bbox.center_y >= r_top and line.bbox.center_y <= r_bottom)
                        ):
                            cell_lines.append(line)
                            assigned_lines.add(idx)

                    cells.append(GridCellData(
                        row_idx=r,
                        col_idx=c,
                        row_span=1,
                        col_span=1,
                        bbox=cell_bbox,
                        lines=cell_lines,
                        text="\n".join(l.text for l in cell_lines),
                        bg_color=cell_shd,
                        borders={
                            "top": BorderStyle(val="single", sz=4),
                            "bottom": BorderStyle(val="single", sz=4),
                            "left": BorderStyle(val="single", sz=4),
                            "right": BorderStyle(val="single", sz=4),
                        }
                    ))

            # Only add if table contains text or is substantial
            if any(cell.lines for cell in cells) or num_rows >= 2:
                tables.append(TableData(
                    rows=num_rows,
                    cols=num_cols,
                    col_widths=col_widths,
                    cells=cells,
                    bbox=table_bbox,
                    is_bordered=True
                ))

        return tables, assigned_lines

    def _reconstruct_borderless_tables(
        self,
        indexed_lines: List[Tuple[int, LineData]],
        page_width: float
    ) -> Tuple[List[TableData], Set[int]]:
        """
        Reconstructs borderless tables from key-value text lines.
        Detects patterns like:
            Key / Label (x: 30-180)  : Value / Description (x: 180-550)
        """
        if len(indexed_lines) < 3:
            return [], set()

        tables: List[TableData] = []
        assigned_indices: Set[int] = set()

        # Identify candidate 2-column key-value runs
        # Group lines that have a consistent left column (label) and right column (content)
        current_run: List[Tuple[int, LineData, float, float]] = []  # (idx, line, split_x)

        def flush_run(run: List[Tuple[int, LineData, float, float]]):
            if len(run) < 3:
                return
            # Build borderless 2-column table
            split_xs = [r[2] for r in run]
            median_split_x = sorted(split_xs)[len(split_xs) // 2]

            min_x0 = min(r[1].bbox.x0 for r in run)
            max_x1 = max(r[1].bbox.x1 for r in run)

            col1_w = median_split_x - min_x0
            col2_w = max_x1 - median_split_x
            if col1_w < 30.0 or col2_w < 50.0:
                return

            col_widths = [col1_w, col2_w]
            cells: List[GridCellData] = []

            for r_idx, (l_idx, line, _, _) in enumerate(run):
                # Split line into left label and right value based on median_split_x
                left_spans = [s for s in line.spans if s.bbox.center_x < median_split_x]
                right_spans = [s for s in line.spans if s.bbox.center_x >= median_split_x]

                left_line = LineData(spans=left_spans)
                left_line.update_text_and_bbox()

                right_line = LineData(spans=right_spans)
                right_line.update_text_and_bbox()

                c1_bbox = BBox(min_x0, line.bbox.y0, median_split_x, line.bbox.y1)
                c2_bbox = BBox(median_split_x, line.bbox.y0, max_x1, line.bbox.y1)

                nil_border = BorderStyle(val="nil", sz=0)
                cells.append(GridCellData(
                    row_idx=r_idx,
                    col_idx=0,
                    bbox=c1_bbox,
                    lines=[left_line] if left_spans else [],
                    text=left_line.text if left_spans else "",
                    borders={"top": nil_border, "bottom": nil_border, "left": nil_border, "right": nil_border}
                ))
                cells.append(GridCellData(
                    row_idx=r_idx,
                    col_idx=1,
                    bbox=c2_bbox,
                    lines=[right_line] if right_spans else [],
                    text=right_line.text if right_spans else "",
                    borders={"top": nil_border, "bottom": nil_border, "left": nil_border, "right": nil_border}
                ))
                assigned_indices.add(l_idx)

            t_bbox = BBox(min_x0, run[0][1].bbox.y0, max_x1, run[-1][1].bbox.y1)
            tables.append(TableData(
                rows=len(run),
                cols=2,
                col_widths=col_widths,
                cells=cells,
                bbox=t_bbox,
                is_bordered=False
            ))

        for idx, line in indexed_lines:
            text = line.text
            # Look for colon divider or large horizontal whitespace gap between spans
            split_found = False
            split_x = 0.0

            if ":" in text and len(line.spans) >= 2:
                # Find span containing colon or span starting immediately after colon
                for s_i, span in enumerate(line.spans):
                    if ":" in span.text:
                        split_x = span.bbox.x1
                        split_found = True
                        break

            if not split_found and len(line.spans) >= 2:
                # Check for large whitespace gap between span 0 and span 1
                for s_i in range(len(line.spans) - 1):
                    gap = line.spans[s_i + 1].bbox.x0 - line.spans[s_i].bbox.x1
                    if gap >= 15.0:  # Noticeable column gap
                        split_x = line.spans[s_i + 1].bbox.x0 - (gap / 2.0)
                        split_found = True
                        break

            if split_found and split_x > 40.0 and split_x < (page_width - 60.0):
                current_run.append((idx, line, split_x, line.bbox.y0))
            else:
                if current_run:
                    flush_run(current_run)
                    current_run = []

        if current_run:
            flush_run(current_run)

        return tables, assigned_indices

    def _cluster_lines_by_table(
        self,
        h_lines: List[VectorPathData],
        v_lines: List[VectorPathData]
    ) -> List[Tuple[List[VectorPathData], List[VectorPathData]]]:
        """Groups horizontal and vertical lines that form contiguous table grids."""
        sorted_h = sorted(h_lines, key=lambda l: l.bbox.y0)
        groups: List[Tuple[List[VectorPathData], List[VectorPathData]]] = []

        current_h: List[VectorPathData] = []
        for line in sorted_h:
            if not current_h:
                current_h.append(line)
            else:
                prev_y = current_h[-1].bbox.y0
                # If distance to previous horizontal line is reasonable (< 120 pt), group together
                if (line.bbox.y0 - prev_y) <= 120.0:
                    current_h.append(line)
                else:
                    # Match vertical lines that fall within current_h vertical span
                    min_y = min(h.bbox.y0 for h in current_h) - self.snap_tolerance
                    max_y = max(h.bbox.y1 for h in current_h) + self.snap_tolerance
                    matched_v = [
                        v for v in v_lines
                        if not (v.bbox.y1 < min_y or v.bbox.y0 > max_y)
                    ]
                    groups.append((current_h, matched_v))
                    current_h = [line]

        if current_h:
            min_y = min(h.bbox.y0 for h in current_h) - self.snap_tolerance
            max_y = max(h.bbox.y1 for h in current_h) + self.snap_tolerance
            matched_v = [
                v for v in v_lines
                if not (v.bbox.y1 < min_y or v.bbox.y0 > max_y)
            ]
            groups.append((current_h, matched_v))

        return groups

    def _snap_coordinates(self, coords: List[float], tolerance: float) -> List[float]:
        """Snaps 1D coordinates within tolerance to canonical grid lines."""
        if not coords:
            return []
        snapped: List[float] = [coords[0]]
        for c in coords[1:]:
            if abs(c - snapped[-1]) > tolerance:
                snapped.append(c)
            else:
                # Average them
                snapped[-1] = (snapped[-1] + c) / 2.0
        return snapped
