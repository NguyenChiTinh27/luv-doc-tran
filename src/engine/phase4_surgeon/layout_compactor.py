"""
layout_compactor.py
───────────────────
Provides automated OpenXML layout compaction: injects character scaling (<w:w>)
and negative character spacing (<w:spacing>) on runs to prevent cell text wrapping
and enforce page count parity.
"""

from typing import Optional
import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def apply_line_compaction(
    doc: docx.Document,
    compaction_factor: float = 0.95,
    w_val: Optional[int] = None,
    spacing_val: Optional[int] = None
) -> int:
    """
    Applies font character scaling and negative tracking on text runs:
    - compaction_factor == 1.0: no compaction applied.
    - compaction_factor == 0.95: w:w="94", w:spacing="-4" (default mild compaction).
    - compaction_factor <= 0.90: w:w="90", w:spacing="-8" (aggressive compaction).

    Returns count of compacted runs.
    """
    if compaction_factor >= 1.0 and w_val is None and spacing_val is None:
        return 0

    if w_val is None:
        w_val = 90 if compaction_factor <= 0.90 else 94
    if spacing_val is None:
        spacing_val = -8 if compaction_factor <= 0.90 else -4

    compacted_count = 0

    # Compact runs in table cells (primary target for overflow prevention)
    for tbl_el in doc.element.xpath("//w:tbl"):
        for p in tbl_el.xpath(".//w:p"):
            for r in p.xpath(".//w:r"):
                rPr = r.find(qn("w:rPr"))
                if rPr is None:
                    rPr = OxmlElement("w:rPr")
                    r.insert(0, rPr)

                # Inject w:w
                w_el = rPr.find(qn("w:w"))
                if w_el is None:
                    w_el = OxmlElement("w:w")
                    rPr.append(w_el)
                w_el.set(qn("w:val"), str(w_val))

                # Inject w:spacing
                sp_el = rPr.find(qn("w:spacing"))
                if sp_el is None:
                    sp_el = OxmlElement("w:spacing")
                    rPr.append(sp_el)
                sp_el.set(qn("w:val"), str(spacing_val))

                compacted_count += 1

    return compacted_count
