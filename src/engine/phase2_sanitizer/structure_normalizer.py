"""
structure_normalizer.py
────────────────────────
Section header normalization and inverted flow repair.
- Normalizes Section Headers consistently to Roman numerals (I. through XVI.)
- Inverted Section Header Repair: Detects out-of-order section headers in body streams
  (e.g. Section 1 appearing after Section 4 in PDF streams) and partitions/reorders them
  in strictly ascending order (1 -> 2 -> ... -> 16).
- Standardizes standard GHS/MSDS 16-section titles and numbering.
"""

from __future__ import annotations
import re
from typing import List, Tuple, Optional, Dict, Any

from src.engine.phase1_extractor.models import (
    TextBlock, TextSpan, ExtractedPage, ExtractedDocument
)

ROMAN_NUMERALS = [
    "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X",
    "XI", "XII", "XIII", "XIV", "XV", "XVI"
]

GHS_SECTIONS: Dict[int, Dict[str, str]] = {
    1: {
        "roman": "I.",
        "en": "Identification of the substance/mixture and of the company/undertaking",
        "vi": "I. Nhận dạng hóa chất và thông tin về nhà cung cấp"
    },
    2: {
        "roman": "II.",
        "en": "Hazards Identification",
        "vi": "II. Nhận dạng đặc tính nguy hiểm"
    },
    3: {
        "roman": "III.",
        "en": "Composition / Information on Ingredients",
        "vi": "III. Thông tin về thành phần các chất"
    },
    4: {
        "roman": "IV.",
        "en": "First Aid Measures",
        "vi": "IV. Biện pháp sơ cứu về y tế"
    },
    5: {
        "roman": "V.",
        "en": "Fire Fighting Measures",
        "vi": "V. Biện pháp xử lý khi có hỏa hoạn"
    },
    6: {
        "roman": "VI.",
        "en": "Accidental Release Measures",
        "vi": "VI. Biện pháp phòng ngừa sự cố tràn đổ"
    },
    7: {
        "roman": "VII.",
        "en": "Handling and Storage",
        "vi": "VII. Sử dụng và bảo quản"
    },
    8: {
        "roman": "VIII.",
        "en": "Exposure Controls / Personal Protection",
        "vi": "VIII. Kiểm soát phơi nhiễm / Bảo vệ cá nhân"
    },
    9: {
        "roman": "IX.",
        "en": "Physical and Chemical Properties",
        "vi": "IX. Đặc tính lý, hóa học"
    },
    10: {
        "roman": "X.",
        "en": "Stability and Reactivity",
        "vi": "X. Tính ổn định và khả năng phản ứng"
    },
    11: {
        "roman": "XI.",
        "en": "Toxicological Information",
        "vi": "XI. Thông tin về độc tính"
    },
    12: {
        "roman": "XII.",
        "en": "Ecological Information",
        "vi": "XII. Thông tin về sinh thái"
    },
    13: {
        "roman": "XIII.",
        "en": "Disposal Considerations",
        "vi": "XIII. Biện pháp tiêu hủy"
    },
    14: {
        "roman": "XIV.",
        "en": "Transport Information",
        "vi": "XIV. Thông tin vận chuyển"
    },
    15: {
        "roman": "XV.",
        "en": "Regulatory Information",
        "vi": "XV. Thông tin pháp lý"
    },
    16: {
        "roman": "XVI.",
        "en": "Other Information",
        "vi": "XVI. Thông tin khác"
    }
}


class StructureNormalizer:
    """Repairs out-of-order section headers and normalizes section titles to canonical Roman numerals."""

    SECTION_HEADER_RE = re.compile(
        r"^(?:"
        r"(?:SECTION|Section|MỤC|Mục|Phần)\s*([0-9]{1,2}|I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII|XIII|XIV|XV|XVI)[\.\:\s\-]+(.*)"
        r"|"
        r"(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII|XIII|XIV|XV|XVI)[\.\:\s]+(.*)"
        r"|"
        r"([1-9]|1[0-6])[\.\:\s\-]+(Identification|Hazards?|Composition|First|Fire|Accidental|Handling|Exposure|Physical|Stability|Toxicological|Ecological|Disposal|Transport|Regulatory|Other|Nhận|Thành|Biện|Kiểm|Đặc|Tính|Thông).*"
        r")$",
        re.IGNORECASE
    )

    @classmethod
    def roman_to_int(cls, roman_str: str) -> Optional[int]:
        """Converts Roman numeral or string integer (1-16) to int."""
        if not roman_str:
            return None
        s = roman_str.strip().upper()
        if s.isdigit():
            val = int(s)
            return val if 1 <= val <= 16 else None
        if s in ROMAN_NUMERALS:
            return ROMAN_NUMERALS.index(s) + 1
        return None

    @classmethod
    def match_section_header(cls, text: str) -> Optional[Tuple[int, str]]:
        """
        Matches a string against MSDS/GHS section patterns.
        Returns:
            Tuple of (section_number: 1..16, title_text) if matched, None otherwise.
        """
        if not text or not text.strip():
            return None
        m = cls.SECTION_HEADER_RE.match(text.strip())
        if not m:
            return None
        num_str = m.group(1) or m.group(3) or m.group(5)
        title_str = m.group(2) or m.group(4) or m.group(6) or ""
        sec_idx = cls.roman_to_int(num_str) if num_str else None
        if sec_idx and 1 <= sec_idx <= 16:
            return sec_idx, title_str.strip()
        return None

    @classmethod
    def normalize_section_header(cls, text: str) -> str:
        """
        Converts any section header into canonical Roman numeral format:
        e.g. 'SECTION 1: Chemical Product...' -> 'I. Identification of the substance/mixture and of the company/undertaking'
        """
        sec_match = cls.match_section_header(text)
        if not sec_match:
            return text
        sec_idx, original_title = sec_match
        if sec_idx in GHS_SECTIONS:
            info = GHS_SECTIONS[sec_idx]
            # If original title is meaningful, use it or fallback to canonical English title
            title = original_title if len(original_title) > 3 else info["en"]
            return f"{info['roman']} {title}"
        return text

    @classmethod
    def repair_inverted_headers(cls, blocks: List[TextBlock]) -> List[TextBlock]:
        """
        Detects out-of-order section headers in blocks and reorders section partitions
        so sections 1 through 16 appear in strictly ascending order.
        """
        if not blocks:
            return []

        # Find section positions
        sec_positions: List[Tuple[int, int, TextBlock]] = []  # (sec_idx, block_index, block)
        for idx, blk in enumerate(blocks):
            m = cls.match_section_header(blk.text)
            if m:
                sec_positions.append((m[0], idx, blk))

        if len(sec_positions) < 2:
            return blocks

        indices = [s[0] for s in sec_positions]
        # Check if already strictly ascending
        if indices == sorted(indices):
            return blocks

        # Partition blocks into preamble and section blocks
        first_sec_pos = sec_positions[0][1]
        preamble = blocks[:first_sec_pos]
        section_blocks: Dict[int, List[TextBlock]] = {}

        for s_i in range(len(sec_positions)):
            curr_sec_num, curr_pos, _ = sec_positions[s_i]
            next_pos = sec_positions[s_i + 1][1] if s_i + 1 < len(sec_positions) else len(blocks)
            block_slice = blocks[curr_pos:next_pos]
            if curr_sec_num not in section_blocks:
                section_blocks[curr_sec_num] = []
            section_blocks[curr_sec_num].extend(block_slice)

        # Reassemble in sorted section order
        reordered: List[TextBlock] = list(preamble)
        for sorted_sec in sorted(section_blocks.keys()):
            reordered.extend(section_blocks[sorted_sec])

        return reordered

    @classmethod
    def normalize_page(cls, page: ExtractedPage) -> ExtractedPage:
        """
        Normalizes section headers and repairs inverted order across an ExtractedPage.
        """
        # 1. Normalize section headers in text
        for b in page.blocks:
            if cls.match_section_header(b.text):
                b.text = cls.normalize_section_header(b.text)
                for s in b.spans:
                    s.text = cls.normalize_section_header(s.text)

        # 2. Repair inverted section order
        page.blocks = cls.repair_inverted_headers(page.blocks)
        return page

    @classmethod
    def normalize_document(cls, doc: ExtractedDocument) -> ExtractedDocument:
        """
        Normalizes section headers and flow across all pages of an ExtractedDocument.
        """
        for p in doc.pages:
            cls.normalize_page(p)
        return doc
