"""
style_inferencer.py
───────────────────
Analyzes typography, maps PDF font names to Word-compatible families,
and groups individual text lines into semantic paragraphs while detecting
bullet lists and precise line-spacing.
"""

from __future__ import annotations
import re
from typing import List, Tuple, Optional, Dict, Any

from src.pdf2word_core.models.schema import (
    BBox, SpanData, LineData, ParagraphData
)

FONT_MAP = {
    "times": "Times New Roman",
    "arial": "Arial",
    "helvetica": "Arial",
    "calibri": "Calibri",
    "cambria": "Cambria",
    "courier": "Courier New",
    "georgia": "Georgia",
    "tahoma": "Tahoma",
    "verdana": "Verdana",
    "songti": "SimSun",
    "simsun": "SimSun",
    "heiti": "SimHei",
    "kaiti": "KaiTi",
    "msyh": "Microsoft YaHei",
    "yahei": "Microsoft YaHei",
}

TERMINAL_PUNCTUATION = {'.', '!', '?', ':', ';', '。', '！', '？', '；', '…', '"', '”', '’'}


class StyleInferencer:
    """Infers typography styles and consolidates lines into paragraphs."""

    @staticmethod
    def map_font_name(pdf_font: str) -> str:
        """Map PDF internal font name (e.g., ABCDEE+Arial-BoldMT) to standard font family."""
        if not pdf_font:
            return "Arial"
        # Strip subset prefix (e.g., 'ABCDEF+')
        clean_name = re.sub(r'^[A-Z]{6}\+', '', pdf_font).lower()
        for key, target in FONT_MAP.items():
            if key in clean_name:
                return target
        return "Arial"

    @classmethod
    def consolidate_paragraphs(
        cls, lines: List[LineData], page_width: float
    ) -> List[ParagraphData]:
        """
        Merges consecutive lines into paragraphs when they share left alignment,
        font family, font size, and consistent line-spacing.
        """
        if not lines:
            return []

        # Sort lines top-to-bottom
        sorted_lines = sorted(lines, key=lambda l: (round(l.bbox.y0 / 2.0) * 2.0, l.bbox.x0))
        paragraphs: List[ParagraphData] = []
        current_lines: List[LineData] = [sorted_lines[0]]

        for i in range(1, len(sorted_lines)):
            prev = sorted_lines[i - 1]
            curr = sorted_lines[i]

            should_merge = cls._should_merge_lines(prev, curr, page_width)
            if should_merge:
                current_lines.append(curr)
            else:
                p = cls._build_paragraph(current_lines)
                paragraphs.append(p)
                current_lines = [curr]

        if current_lines:
            p = cls._build_paragraph(current_lines)
            paragraphs.append(p)

        return paragraphs

    @classmethod
    def _should_merge_lines(
        cls, prev: LineData, curr: LineData, page_width: float
    ) -> bool:
        """Checks if two consecutive lines belong to the same paragraph."""
        # If current line starts with a list bullet, do not merge
        if curr.is_bullet:
            return False

        # If previous line ends with terminal punctuation and next line has different left x
        prev_text = prev.text.strip()
        if prev_text and prev_text[-1] in TERMINAL_PUNCTUATION:
            # Significant vertical gap (> 1.8x line height) means new paragraph
            dy = curr.bbox.y0 - prev.bbox.y1
            avg_font = (prev.spans[0].font_size + curr.spans[0].font_size) / 2.0 if (prev.spans and curr.spans) else 10.0
            if dy > avg_font * 0.8:
                return False

        # Check font consistency
        if prev.spans and curr.spans:
            p_span = prev.spans[0]
            c_span = curr.spans[0]
            if abs(p_span.font_size - c_span.font_size) > 1.5:
                return False

        # Check vertical line-spacing
        dy = curr.bbox.y0 - prev.bbox.y1
        avg_font = prev.spans[0].font_size if prev.spans else 10.0
        # Line spacing should be within normal bounds (< 1.6x font size)
        if dy < -2.0 or dy > (avg_font * 1.5):
            return False

        # Check left margin alignment (within 10 pt tolerance)
        dx_left = abs(curr.bbox.x0 - prev.bbox.x0)
        if dx_left > 12.0:
            return False

        return True

    @classmethod
    def _build_paragraph(cls, lines: List[LineData]) -> ParagraphData:
        """Constructs a ParagraphData object from a group of lines."""
        text = " ".join(l.text.strip() for l in lines)
        min_x0 = min(l.bbox.x0 for l in lines)
        min_y0 = min(l.bbox.y0 for l in lines)
        max_x1 = max(l.bbox.x1 for l in lines)
        max_y1 = max(l.bbox.y1 for l in lines)
        bbox = BBox(min_x0, min_y0, max_x1, max_y1)

        # Detect alignment
        align = "left"
        # If line starts with bullet
        indent_left = 0.0
        if lines and lines[0].is_bullet:
            indent_left = 18.0  # ~0.25 inch indent

        # Calculate line spacing multiplier
        line_spacing = 1.15
        if len(lines) > 1:
            dy = (lines[-1].bbox.y0 - lines[0].bbox.y0) / (len(lines) - 1)
            font_size = lines[0].spans[0].font_size if lines[0].spans else 10.0
            if font_size > 0:
                line_spacing = max(1.0, min(1.5, dy / font_size))

        return ParagraphData(
            lines=lines,
            bbox=bbox,
            text=text,
            align=align,
            line_spacing=line_spacing,
            space_before=0.0,
            space_after=2.0,
            indent_left=indent_left
        )
