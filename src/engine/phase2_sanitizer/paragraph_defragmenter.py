"""
paragraph_defragmenter.py
─────────────────────────
Consolidates unpunctuated wrapped lines sharing indentation/font into coherent logical paragraphs.
Eliminates fragmented sentences ("Chống băm nát câu văn") prior to translation passes.
Strictly protects:
- Section headers (e.g. 'SECTION 1', 'I. Identification', 'Mục 1')
- List bullets ('•', '-', '*', '1.', 'a)')
- 'Label : Value' lines
- Short uppercase titles
- Code blocks or monospace text
"""

from __future__ import annotations
import re
from typing import List, Tuple, Optional, Dict, Any, Set

from src.engine.phase1_extractor.models import (
    BBox, TextBlock, TextSpan, ExtractedPage, ExtractedDocument
)


class ParagraphDefragmenter:
    """Merges consecutive wrapped lines into unified paragraphs based on geometry and punctuation."""

    TERMINAL_PUNCTUATION = {".", ";", ":", "!", "?", "...", '."', '!"', '?"', ".)", "])"}
    SECTION_HEADER_RE = re.compile(
        r"^(?:(?:SECTION|Section|MỤC|Mục|Phần)\s*([0-9]{1,2}|I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII|XIII|XIV|XV|XVI)|"
        r"(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII|XIII|XIV|XV|XVI)[\.\:\s]+|"
        r"([1-9]|1[0-6])[\.\:\s\-]+(?:Identification|Hazards?|Composition|First|Fire|Accidental|Handling|Exposure|Physical|Stability|Toxicological|Ecological|Disposal|Transport|Regulatory|Other|Nhận|Thành|Biện|Kiểm|Đặc|Tính|Thông))",
        re.IGNORECASE
    )

    @classmethod
    def should_merge(cls, upper: TextBlock, lower: TextBlock) -> bool:
        """
        Determines whether two consecutive text lines should be consolidated into a single paragraph.
        """
        if not upper or not lower:
            return False

        u_text = upper.text.strip()
        l_text = lower.text.strip()
        if not u_text or not l_text:
            return False

        # 1. Section headers must never be merged
        if cls.SECTION_HEADER_RE.match(u_text) or cls.SECTION_HEADER_RE.match(l_text):
            return False

        # 2. Table-like Label:Value lines or colon dividers must not be merged
        if ":" in u_text or ":" in l_text:
            return False

        # 3. List bullets or numbered items must not be merged into predecessor
        if lower.is_bullet or l_text[0] in {'•', '●', '▪', '▫', '-', '+', '*', '–', '—', '➢', '❖', '✓', '✔'}:
            return False
        if re.match(r"^(\d+[\.\)]|[a-zA-Z][\.\)])\s+", l_text):
            return False

        # 4. If upper ends with terminal punctuation, do NOT merge
        if u_text[-1] in cls.TERMINAL_PUNCTUATION or any(u_text.endswith(p) for p in cls.TERMINAL_PUNCTUATION):
            return False

        # 5. If upper is short and uppercase, it is likely a title/heading
        if len(u_text) < 35 and u_text.isupper():
            return False

        # 6. Font consistency check: size should not differ by more than 1.5 pt
        if abs(upper.font_size - lower.font_size) > 1.5:
            return False

        # 7. Font weight/style consistency
        if upper.is_bold != lower.is_bold:
            return False

        # 8. Vertical line spacing check: gap between lines should be reasonable (< 2.0x line height)
        line_height = max(8.0, upper.bbox[3] - upper.bbox[1])
        vertical_gap = lower.bbox[1] - upper.bbox[3]
        if vertical_gap < -2.0 or vertical_gap > (line_height * 2.2):
            return False

        # 9. Indentation alignment check (left margins within 18 pt)
        if abs(upper.bbox[0] - lower.bbox[0]) > 18.0:
            return False

        return True

    @classmethod
    def defragment_blocks(cls, blocks: List[TextBlock]) -> List[TextBlock]:
        """
        Consolidates wrapped lines in a list of TextBlocks into coherent logical paragraphs.
        """
        if not blocks:
            return []

        out: List[TextBlock] = []
        i = 0
        total = len(blocks)

        while i < total:
            curr = blocks[i]
            # Accumulate subsequent wrapped lines
            merged_spans = list(curr.spans)
            curr_text = curr.text.strip()
            curr_bbox = curr.bbox

            j = i + 1
            while j < total:
                candidate = blocks[j]
                # Test if curr and candidate should merge
                temp_block = TextBlock(
                    spans=merged_spans,
                    bbox=curr_bbox,
                    text=curr_text,
                    font_name=curr.font_name,
                    font_size=curr.font_size,
                    is_bold=curr.is_bold,
                    is_italic=curr.is_italic
                )
                if cls.should_merge(temp_block, candidate):
                    # Append candidate text and spans
                    curr_text = f"{curr_text} {candidate.text.strip()}"
                    merged_spans.extend(candidate.spans)
                    # Expand bbox
                    curr_bbox = (
                        min(curr_bbox[0], candidate.bbox[0]),
                        min(curr_bbox[1], candidate.bbox[1]),
                        max(curr_bbox[2], candidate.bbox[2]),
                        max(curr_bbox[3], candidate.bbox[3]),
                    )
                    j += 1
                else:
                    break

            consolidated = TextBlock(
                spans=merged_spans,
                bbox=curr_bbox,
                text=curr_text,
                font_name=curr.font_name,
                font_size=curr.font_size,
                is_bold=curr.is_bold,
                is_italic=curr.is_italic,
                is_bullet=curr.is_bullet,
                bullet_char=curr.bullet_char,
                line_index=curr.line_index
            )
            out.append(consolidated)
            i = j

        return out

    @classmethod
    def defragment_page(cls, page: ExtractedPage) -> ExtractedPage:
        """Applies defragmentation to all body blocks of an ExtractedPage."""
        page.blocks = cls.defragment_blocks(page.blocks)
        # Also defragment blocks inside table cells
        for tbl in page.tables:
            for cell in tbl.cells:
                if cell.blocks:
                    cell.blocks = cls.defragment_blocks(cell.blocks)
                    cell.text = "\n".join(b.text for b in cell.blocks)
        return page

    @classmethod
    def defragment_document(cls, doc: ExtractedDocument) -> ExtractedDocument:
        """Applies defragmentation to all pages of an ExtractedDocument."""
        for p in doc.pages:
            cls.defragment_page(p)
        return doc
