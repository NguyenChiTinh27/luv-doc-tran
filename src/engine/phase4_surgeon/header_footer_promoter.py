"""
header_footer_promoter.py
─────────────────────────
Promotes drifting body pagination and document titles into native Word Section
headers and footers (<w:hdr> / <w:ftr>) with dynamic PAGE / NUMPAGES XML fields,
and strips hardcoded pagination paragraphs from flowing body text.
"""

import re
from typing import Optional, Dict, Any, List
import docx
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
from docx.shared import Pt, RGBColor

PAGINATION_REGEX = re.compile(
    r"^(?:Page\s+\d+\s*(?:of|/)\s*\d+|p\.\s*\d+|Page\s+\d+|Trang\s+\d+(?:\s*(?:/|của)\s*\d+)?|\(Contd\.\s+on\s+page\s+\d+\))\s*$",
    re.IGNORECASE
)


def add_field(p: docx.text.paragraph.Paragraph, field_name: str) -> None:
    """Helper to inject dynamic Word XML field (e.g., PAGE or NUMPAGES)."""
    fldSimple = OxmlElement("w:fldSimple")
    fldSimple.set(qn("w:instr"), field_name)
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = "1"
    r.append(t)
    fldSimple.append(r)
    p._p.append(fldSimple)


def promote_headers_footers(
    doc: docx.Document,
    title_text: Optional[str] = None,
    remove_body_pagination: bool = True
) -> Dict[str, Any]:
    """
    Promotes pagination and running titles to native headers/footers:
    1. Sets up Section header: Document Title (if provided or detected).
    2. Sets up Section footer: Right-aligned 'Page [PAGE] / [NUMPAGES]'.
    3. Scans body paragraphs: purges unpromoted static pagination strings.

    Returns summary metrics dict.
    """
    removed_body_paragraphs = 0
    promoted_sections = 0

    # 1. Configure Header & Footer on Sections
    for s_idx, section in enumerate(doc.sections):
        promoted_sections += 1

        # Header
        header = section.header
        if title_text:
            if not header.paragraphs or not header.paragraphs[0].text:
                hp = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
                hp.text = title_text
                hp.paragraph_format.space_before = Pt(0)
                hp.paragraph_format.space_after = Pt(0)
                if hp.runs:
                    hp.runs[0].font.name = "Times New Roman"
                    hp.runs[0].font.size = Pt(8.5)
                    hp.runs[0].font.color.rgb = RGBColor(128, 128, 128)

        # Footer with dynamic PAGE / NUMPAGES
        footer = section.footer
        fp = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        # Only setup if not already containing dynamic field
        if not fp._p.xpath(".//w:fldSimple") and not fp._p.xpath(".//w:instrText"):
            fp.text = ""
            fp.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.RIGHT
            fp.paragraph_format.space_before = Pt(0)
            fp.paragraph_format.space_after = Pt(0)

            r1 = fp.add_run("Page ")
            r1.font.name = "Times New Roman"
            r1.font.size = Pt(9)
            r1.font.color.rgb = RGBColor(100, 100, 100)

            add_field(fp, "PAGE")

            r2 = fp.add_run(" / ")
            r2.font.name = "Times New Roman"
            r2.font.size = Pt(9)
            r2.font.color.rgb = RGBColor(100, 100, 100)

            add_field(fp, "NUMPAGES")

    # 2. Purge unpromoted static pagination from body paragraphs
    if remove_body_pagination:
        for p in list(doc.paragraphs):
            text = p.text.strip()
            if PAGINATION_REGEX.match(text):
                try:
                    p._p.getparent().remove(p._p)
                    removed_body_paragraphs += 1
                except Exception:
                    pass

    return {
        "promoted_sections": promoted_sections,
        "removed_body_paragraphs": removed_body_paragraphs
    }
