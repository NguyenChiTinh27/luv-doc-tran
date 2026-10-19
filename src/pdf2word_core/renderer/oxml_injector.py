"""
oxml_injector.py
────────────────
Low-level OpenXML (ECMA-376) injection and surgical manipulation using lxml.etree.
Bypasses python-docx high-level API limitations to enforce exact fixed table layouts,
zero-drift cell padding, anti-overflow font compensation, and cantSplit row locks.
"""

from __future__ import annotations
from typing import List, Dict, Optional, Tuple, Any
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.table import Table, _Cell, _Row
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from docx.section import Section

from src.pdf2word_core.models.schema import BorderStyle


class OxmlInjector:
    """Surgical low-level OpenXML injector for formatting locks and compensation."""

    @staticmethod
    def set_table_fixed_layout(table: Table) -> None:
        """Enforces <w:tblLayout w:type=\"fixed\"/> so Word never auto-resizes columns."""
        tblPr = table._tbl.tblPr
        tblLayout = tblPr.find(qn("w:tblLayout"))
        if tblLayout is None:
            tblLayout = OxmlElement("w:tblLayout")
            tblPr.append(tblLayout)
        tblLayout.set(qn("w:type"), "fixed")

    @staticmethod
    def set_table_grid_widths(table: Table, col_widths_dxa: List[int]) -> None:
        """Explicitly sets <w:tblGrid><w:gridCol w:w=\"...\"/></w:tblGrid> down to 1 twip."""
        tbl = table._tbl
        old_grid = tbl.find(qn("w:tblGrid"))
        if old_grid is not None:
            tbl.remove(old_grid)

        tblGrid = OxmlElement("w:tblGrid")
        total_width = 0
        for w in col_widths_dxa:
            col = OxmlElement("w:gridCol")
            col.set(qn("w:w"), str(w))
            tblGrid.append(col)
            total_width += w

        tblPr = tbl.tblPr
        # Insert tblGrid after tblPr
        tbl.insert(tbl.index(tblPr) + 1, tblGrid)

        # Set tblW to total width
        tblW = tblPr.find(qn("w:tblW"))
        if tblW is None:
            tblW = OxmlElement("w:tblW")
            tblPr.append(tblW)
        tblW.set(qn("w:w"), str(total_width))
        tblW.set(qn("w:type"), "dxa")

    @staticmethod
    def set_cell_width(cell: _Cell, width_dxa: int) -> None:
        """Sets <w:tcW w:w=\"...\" w:type=\"dxa\"/> on cell."""
        tcPr = cell._tc.get_or_add_tcPr()
        tcW = tcPr.find(qn("w:tcW"))
        if tcW is None:
            tcW = OxmlElement("w:tcW")
            tcPr.append(tcW)
        tcW.set(qn("w:w"), str(width_dxa))
        tcW.set(qn("w:type"), "dxa")

    @staticmethod
    def set_cell_margins(
        cell: _Cell,
        top_dxa: int = 40,
        bottom_dxa: int = 40,
        left_dxa: int = 40,
        right_dxa: int = 40
    ) -> None:
        """
        Explicitly injects <w:tcMar> on individual cells to eliminate Word's
        default 5.4pt (108 dxa) padding which triggers cascading text-wrapping.
        """
        tcPr = cell._tc.get_or_add_tcPr()
        old_mar = tcPr.find(qn("w:tcMar"))
        if old_mar is not None:
            tcPr.remove(old_mar)

        tcMar = OxmlElement("w:tcMar")
        for side, val in [("top", top_dxa), ("bottom", bottom_dxa), ("left", left_dxa), ("right", right_dxa)]:
            m = OxmlElement(f"w:{side}")
            m.set(qn("w:w"), str(val))
            m.set(qn("w:type"), "dxa")
            tcMar.append(m)
        tcPr.append(tcMar)

    @staticmethod
    def set_cell_shading(cell: _Cell, hex_color: str) -> None:
        """Sets cell background fill color <w:shd w:fill=\"...\">."""
        if not hex_color:
            return
        clean_hex = hex_color.replace("#", "").upper()
        tcPr = cell._tc.get_or_add_tcPr()
        shd = tcPr.find(qn("w:shd"))
        if shd is None:
            shd = OxmlElement("w:shd")
            tcPr.append(shd)
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), clean_hex)

    @staticmethod
    def set_cell_borders(cell: _Cell, borders: Dict[str, BorderStyle]) -> None:
        """Injects explicit <w:tcBorders> for top, bottom, left, right."""
        tcPr = cell._tc.get_or_add_tcPr()
        old_borders = tcPr.find(qn("w:tcBorders"))
        if old_borders is not None:
            tcPr.remove(old_borders)

        tcBorders = OxmlElement("w:tcBorders")
        for side in ["top", "left", "bottom", "right"]:
            b_style = borders.get(side)
            b_elem = OxmlElement(f"w:{side}")
            if b_style and b_style.val != "nil":
                b_elem.set(qn("w:val"), b_style.val)
                b_elem.set(qn("w:sz"), str(b_style.sz))
                b_elem.set(qn("w:space"), str(b_style.space))
                b_elem.set(qn("w:color"), b_style.color if b_style.color != "auto" else "auto")
            else:
                b_elem.set(qn("w:val"), "nil")
            tcBorders.append(b_elem)
        tcPr.append(tcBorders)

    @staticmethod
    def lock_row_anti_split(row: _Row) -> None:
        """Injects <w:cantSplit/> on table row to prevent cross-page fracturing."""
        trPr = row._tr.get_or_add_trPr()
        if trPr.find(qn("w:cantSplit")) is None:
            trPr.append(OxmlElement("w:cantSplit"))

    @staticmethod
    def set_row_header(row: _Row) -> None:
        """Injects <w:tblHeader/> to repeat header row across pages."""
        trPr = row._tr.get_or_add_trPr()
        if trPr.find(qn("w:tblHeader")) is None:
            trPr.append(OxmlElement("w:tblHeader"))

    @staticmethod
    def apply_anti_overflow_compensation(
        run: Run,
        scale_pct: int = 94,
        spacing_twips: int = -4
    ) -> None:
        """
        Anti-overflow character tracking and horizontal scaling:
        Injects <w:w w:val=\"94\"/> and <w:spacing w:val=\"-4\"/> to condense text
        and prevent a single word from spilling onto an extra line.
        """
        rPr = run._r.get_or_add_rPr()

        # Horizontal scaling (<w:w w:val="94"/>)
        w_elem = rPr.find(qn("w:w"))
        if w_elem is None:
            w_elem = OxmlElement("w:w")
            rPr.append(w_elem)
        w_elem.set(qn("w:val"), str(scale_pct))

        # Letter spacing (<w:spacing w:val="-4"/>)
        sp_elem = rPr.find(qn("w:spacing"))
        if sp_elem is None:
            sp_elem = OxmlElement("w:spacing")
            rPr.append(sp_elem)
        sp_elem.set(qn("w:val"), str(spacing_twips))

    @staticmethod
    def set_paragraph_spacing(
        p: Paragraph,
        before_dxa: int = 0,
        after_dxa: int = 0,
        line_dxa: Optional[int] = None,
        line_rule: str = "auto"
    ) -> None:
        """Sets exact spacing before, after, and line height."""
        pPr = p._p.get_or_add_pPr()
        sp = pPr.find(qn("w:spacing"))
        if sp is None:
            sp = OxmlElement("w:spacing")
            pPr.append(sp)
        sp.set(qn("w:before"), str(before_dxa))
        sp.set(qn("w:after"), str(after_dxa))
        if line_dxa is not None:
            sp.set(qn("w:line"), str(line_dxa))
            sp.set(qn("w:lineRule"), line_rule)

    @staticmethod
    def set_page_geometry(
        section: Section,
        width_pt: float,
        height_pt: float,
        margins_pt: Tuple[float, float, float, float]
    ) -> None:
        """Sets page dimensions and margins down to 1 twip."""
        sectPr = section._sectPr
        pgSz = sectPr.find(qn("w:pgSz"))
        if pgSz is None:
            pgSz = OxmlElement("w:pgSz")
            sectPr.append(pgSz)
        pgSz.set(qn("w:w"), str(round(width_pt * 20)))
        pgSz.set(qn("w:h"), str(round(height_pt * 20)))

        top_m, bot_m, left_m, right_m = margins_pt
        pgMar = sectPr.find(qn("w:pgMar"))
        if pgMar is None:
            pgMar = OxmlElement("w:pgMar")
            sectPr.append(pgMar)
        pgMar.set(qn("w:top"), str(round(top_m * 20)))
        pgMar.set(qn("w:bottom"), str(round(bot_m * 20)))
        pgMar.set(qn("w:left"), str(round(left_m * 20)))
        pgMar.set(qn("w:right"), str(round(right_m * 20)))
        pgMar.set(qn("w:header"), "720")
        pgMar.set(qn("w:footer"), "720")
