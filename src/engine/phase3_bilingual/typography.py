"""
src/engine/phase3_bilingual/typography.py
──────────────────────────────────────────
Authoritative typography specifications and styling constants for Pattern A & Pattern B
bilingual pairing in DOCX documents (R1, R8).

Defines:
  - Colors: DARK_NAVY_HEX ("1B365D"), DARK_NAVY_RGB (27, 54, 93),
            DARK_BLUE_HEX ("002060"), DARK_BLUE_RGB (0, 32, 96),
            BLACK_HEX ("000000"), BLACK_RGB (0, 0, 0)
  - Font styling: Times New Roman, 10.5 pt body, 10.0 pt table
  - Paragraph spacing: 0 pt before, 0 pt after for source, 4.0 pt (80 dxa) after for translation
  - XML helpers: <w:keepNext/> injector, <w:vAlign w:val="center"/> injector,
                 spacing normalizer, and formatting cleaners.
"""

from typing import Any, Optional, Tuple, Union
import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS & METRICS
# ─────────────────────────────────────────────────────────────────────────────
DARK_NAVY_HEX: str = "1B365D"
DARK_NAVY_RGB: RGBColor = RGBColor(0x1B, 0x36, 0x5D)
DARK_NAVY_RGBCOLOR: RGBColor = DARK_NAVY_RGB

DARK_BLUE_HEX: str = "002060"
DARK_BLUE_RGB: RGBColor = RGBColor(0, 32, 96)
DARK_BLUE_RGBCOLOR: RGBColor = DARK_BLUE_RGB

BLACK_HEX: str = "000000"
BLACK_RGB: RGBColor = RGBColor(0, 0, 0)
BLACK_RGBCOLOR: RGBColor = BLACK_RGB

DEFAULT_FONT_NAME: str = "Times New Roman"
DEFAULT_BODY_FONT_SIZE_PT: float = 10.5
DEFAULT_TABLE_FONT_SIZE_PT: float = 10.0

SPACING_AFTER_DXA: int = 80    # 80 dxa = 4.0 pt (Pattern A translated paragraph)
SPACING_AFTER_PT: float = 4.0
SPACING_BEFORE_DXA: int = 0
LINE_SPACING_DXA: int = 240    # single line spacing (240 / 240 = 1.0)
LINE_RULE_AUTO: str = "auto"

W_NS: str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP: dict = {"w": W_NS}

ROMAN_NUMERALS = [
    "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X",
    "XI", "XII", "XIII", "XIV", "XV", "XVI"
]


# ─────────────────────────────────────────────────────────────────────────────
# XML & STYLING HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def add_keep_next(paragraph: Any) -> None:
    """
    Injects <w:keepNext/> into the paragraph's <w:pPr> element.
    Guarantees Word never places a page break between the source line
    and its immediately following translation (R1/R8).
    """
    p_el = getattr(paragraph, "_p", paragraph)
    pPr = p_el.get_or_add_pPr() if hasattr(p_el, "get_or_add_pPr") else p_el.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_el.insert(0, pPr)
    if pPr.find(qn("w:keepNext")) is None:
        pPr.append(OxmlElement("w:keepNext"))


def set_paragraph_spacing(
    paragraph: Any,
    before_dxa: int = 0,
    after_dxa: int = 0,
    line_dxa: int = 240,
    line_rule: str = "auto"
) -> None:
    """
    Sets precise spacing (<w:spacing>) on a paragraph's <w:pPr>.
    """
    p_el = getattr(paragraph, "_p", paragraph)
    pPr = p_el.get_or_add_pPr() if hasattr(p_el, "get_or_add_pPr") else p_el.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_el.insert(0, pPr)

    sp = pPr.find(qn("w:spacing"))
    if sp is None:
        sp = OxmlElement("w:spacing")
        pPr.append(sp)

    sp.set(qn("w:before"), str(before_dxa))
    sp.set(qn("w:after"), str(after_dxa))
    sp.set(qn("w:line"), str(line_dxa))
    sp.set(qn("w:lineRule"), line_rule)


def set_cell_valign_center(cell: Any) -> None:
    """
    Ensures a table cell's vertical alignment is explicitly center (<w:vAlign w:val="center"/>).
    Required for Pattern B table cells (R1/R4).
    """
    tc_el = getattr(cell, "_tc", cell)
    tcPr = tc_el.get_or_add_tcPr() if hasattr(tc_el, "get_or_add_tcPr") else tc_el.find(qn("w:tcPr"))
    if tcPr is None:
        tcPr = OxmlElement("w:tcPr")
        tc_el.insert(0, tcPr)

    vAlign = tcPr.find(qn("w:vAlign"))
    if vAlign is None:
        vAlign = OxmlElement("w:vAlign")
        tcPr.append(vAlign)
    vAlign.set(qn("w:val"), "center")


def clean_inherited_styles(paragraph: Any) -> None:
    """
    Removes any inherited <w:i> (italic) or <w:color> elements from a source paragraph.
    Prevents audit verifiers from misidentifying source lines as translated lines (Check 4).
    """
    p_el = getattr(paragraph, "_p", paragraph)
    # Remove from paragraph runs
    for i_elem in p_el.xpath(".//w:i"):
        parent = i_elem.getparent()
        if parent is not None:
            parent.remove(i_elem)
    for col_elem in p_el.xpath(".//w:color"):
        parent = col_elem.getparent()
        if parent is not None:
            parent.remove(col_elem)

    # Also clean pPr/rPr default run formatting if present
    pPr = p_el.find(qn("w:pPr"))
    if pPr is not None:
        rPr_p = pPr.find(qn("w:rPr"))
        if rPr_p is not None:
            for sub in list(rPr_p):
                if sub.tag in (qn("w:i"), qn("w:color")):
                    rPr_p.remove(sub)
