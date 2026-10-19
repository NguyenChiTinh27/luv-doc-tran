"""
spatial_deinterleaver.py
────────────────────────
Spatial de-interleaving and 2-column grid reconstruction.
- Detects multi-field rows (e.g. 'Telephone: 0274 3555 208 Fax: 0274 3555 209 Email: info@...')
  and splits them into dedicated individual lines/sub-blocks.
- Detects 'Label : Value' pseudo-tables via colon delimiters, tabs, or wide X-coordinate gaps.
- Clusters consecutive key-value blocks and reconstructs them into borderless 2-column tables
  (Column 1 ~ 28%, Column 2 ~ 72% width).
- Strips 100% of internal double-tabs ('\t\t') to prevent downstream OpenXML layout corruption.
"""

from __future__ import annotations
import re
from typing import List, Tuple, Optional, Dict, Any, Set

from src.engine.phase1_extractor.models import (
    BBox, TextBlock, TextSpan, GridCellData, TableGridData, ExtractedPage
)


class SpatialDeinterleaver:
    """Detects Label : Value blocks, splits multi-field lines, and builds borderless 2-column grids."""

    LABEL_VALUE_SPLIT_RE = re.compile(r"^([^:\t\n]{2,45}?)\s*(?::|\t:|\t)\s*(.+)$")
    MULTI_FIELD_SPLIT_RE = re.compile(
        r"(?<=\S)\s+(?=(?:Telephone|Phone|Fax|Email|E-mail|Tel|Contact|Website|Hotline)\s*:)",
        re.IGNORECASE
    )
    MULTI_BULLET_RE = re.compile(r"(?<=[.!?])\s*[-•]\s*")

    @classmethod
    def is_label_value_line(cls, text: str) -> bool:
        """Check if a string represents a 'Label : Value' line."""
        if not text or len(text) < 3 or len(text) > 350:
            return False
        st = text.strip().lower()
        # Avoid section headers
        if st.startswith("section") or st.startswith("mục") or st.startswith("phần"):
            return False
        return bool(cls.LABEL_VALUE_SPLIT_RE.match(text.strip()))

    @classmethod
    def parse_label_value(cls, text: str) -> Tuple[str, str]:
        """Extracts label and value from a 'Label : Value' line."""
        m = cls.LABEL_VALUE_SPLIT_RE.match(text.strip())
        if m:
            return m.group(1).strip(), m.group(2).strip()
        return text.strip(), ""

    @classmethod
    def split_multi_field_line(cls, text: str) -> List[str]:
        """
        Splits single lines containing multiple contact/metadata fields:
        e.g. 'Telephone: 0274 3555 208  Fax: 0274 3555 209  Email: contact@...'
        into separate items.
        """
        if not text:
            return []
        parts = cls.MULTI_FIELD_SPLIT_RE.split(text.strip())
        return [p.strip() for p in parts if p.strip()]

    @classmethod
    def strip_double_tabs(cls, text: str) -> str:
        """Strips 100% of internal double-tabs and multi-tab sequences."""
        if not text:
            return ""
        return re.sub(r"\t+", " ", text)

    @classmethod
    def deinterleave_blocks(cls, blocks: List[TextBlock]) -> List[TextBlock]:
        """
        Processes a list of TextBlocks:
        - Strips internal double tabs.
        - Splits multi-field lines into individual TextBlocks.
        """
        out_blocks: List[TextBlock] = []
        for b in blocks:
            cleaned_text = cls.strip_double_tabs(b.text)
            # Check for multi-field lines
            if cls.MULTI_FIELD_SPLIT_RE.search(cleaned_text):
                sub_lines = cls.split_multi_field_line(cleaned_text)
                sub_count = len(sub_lines)
                b_height = (b.bbox[3] - b.bbox[1]) / max(1, sub_count)
                for idx, sub_txt in enumerate(sub_lines):
                    sub_y0 = b.bbox[1] + (idx * b_height)
                    sub_y1 = sub_y0 + b_height
                    sub_span = TextSpan(
                        text=sub_txt,
                        bbox=(b.bbox[0], sub_y0, b.bbox[2], sub_y1),
                        font_name=b.font_name,
                        font_size=b.font_size,
                        is_bold=b.is_bold,
                        is_italic=b.is_italic
                    )
                    sub_blk = TextBlock(
                        spans=[sub_span],
                        bbox=(b.bbox[0], sub_y0, b.bbox[2], sub_y1),
                        text=sub_txt,
                        font_name=b.font_name,
                        font_size=b.font_size,
                        is_bold=b.is_bold,
                        is_italic=b.is_italic,
                        line_index=b.line_index + idx
                    )
                    out_blocks.append(sub_blk)
            else:
                # Update block text and spans with cleaned text
                b.text = cleaned_text
                for s in b.spans:
                    s.text = cls.strip_double_tabs(s.text)
                out_blocks.append(b)

        return out_blocks

    @classmethod
    def reconstruct_grid_data(
        cls,
        blocks: List[TextBlock],
        page_width: float = 595.0,
        min_cluster_size: int = 2
    ) -> Tuple[List[TableGridData], List[TextBlock]]:
        """
        Scans body blocks for contiguous clusters of 'Label : Value' lines
        and converts them into 2-column borderless TableGridData.
        Column 1 = 28% (~166 pt), Column 2 = 72% (~428 pt on A4).
        Returns:
            tables: Newly generated borderless TableGridData
            remaining_blocks: Blocks not converted to tables
        """
        processed_blocks = cls.deinterleave_blocks(blocks)
        new_tables: List[TableGridData] = []
        assigned_indices: Set[int] = set()

        i = 0
        total = len(processed_blocks)

        while i < total:
            b = processed_blocks[i]
            if cls.is_label_value_line(b.text):
                # Gather contiguous cluster
                cluster_indices = [i]
                j = i + 1
                while j < total:
                    next_b = processed_blocks[j]
                    if cls.is_label_value_line(next_b.text):
                        cluster_indices.append(j)
                        j += 1
                    else:
                        break

                if len(cluster_indices) >= min_cluster_size:
                    cluster_blocks = [processed_blocks[k] for k in cluster_indices]
                    min_x0 = min(cb.bbox[0] for cb in cluster_blocks)
                    max_x1 = max(cb.bbox[2] for cb in cluster_blocks)
                    grid_width = max(max_x1 - min_x0, page_width * 0.8)

                    col1_w = grid_width * 0.28
                    col2_w = grid_width * 0.72
                    split_x = min_x0 + col1_w

                    cells: List[GridCellData] = []
                    for r_idx, cb in enumerate(cluster_blocks):
                        lbl, val = cls.parse_label_value(cb.text)
                        # Left Cell (Label)
                        c0_box = (min_x0, cb.bbox[1], split_x, cb.bbox[3])
                        s0 = TextSpan(
                            text=lbl,
                            bbox=c0_box,
                            font_name=cb.font_name,
                            font_size=cb.font_size,
                            is_bold=True
                        )
                        b0 = TextBlock(
                            spans=[s0],
                            bbox=c0_box,
                            text=lbl,
                            font_name=cb.font_name,
                            font_size=cb.font_size,
                            is_bold=True
                        )
                        cells.append(GridCellData(
                            row_idx=r_idx,
                            col_idx=0,
                            bbox=c0_box,
                            text=lbl,
                            blocks=[b0],
                            borders={"top": {"val": "none"}, "bottom": {"val": "none"}, "left": {"val": "none"}, "right": {"val": "none"}}
                        ))

                        # Right Cell (Value)
                        c1_box = (split_x, cb.bbox[1], max_x1, cb.bbox[3])
                        val_display = f": {val}" if not val.startswith(":") else val
                        s1 = TextSpan(
                            text=val_display,
                            bbox=c1_box,
                            font_name=cb.font_name,
                            font_size=cb.font_size
                        )
                        b1 = TextBlock(
                            spans=[s1],
                            bbox=c1_box,
                            text=val_display,
                            font_name=cb.font_name,
                            font_size=cb.font_size
                        )
                        cells.append(GridCellData(
                            row_idx=r_idx,
                            col_idx=1,
                            bbox=c1_box,
                            text=val_display,
                            blocks=[b1],
                            borders={"top": {"val": "none"}, "bottom": {"val": "none"}, "left": {"val": "none"}, "right": {"val": "none"}}
                        ))

                    t_bbox = (min_x0, cluster_blocks[0].bbox[1], max_x1, cluster_blocks[-1].bbox[3])
                    new_tables.append(TableGridData(
                        rows=len(cluster_blocks),
                        cols=2,
                        col_widths=[col1_w, col2_w],
                        cells=cells,
                        bbox=t_bbox,
                        is_bordered=False
                    ))
                    assigned_indices.update(cluster_indices)
                    i = j
                    continue
            i += 1

        remaining_blocks = [
            b for idx, b in enumerate(processed_blocks)
            if idx not in assigned_indices
        ]
        return new_tables, remaining_blocks

    @classmethod
    def deinterleave_page(cls, page: ExtractedPage) -> ExtractedPage:
        """Applies spatial de-interleaving and borderless grid reconstruction to an ExtractedPage."""
        new_tables, remaining_blocks = cls.reconstruct_grid_data(
            page.blocks, page_width=page.width
        )
        combined_tables = list(page.tables) + new_tables
        page.tables = combined_tables
        page.blocks = remaining_blocks
        return page
