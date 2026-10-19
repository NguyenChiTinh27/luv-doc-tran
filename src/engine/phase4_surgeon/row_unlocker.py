"""
row_unlocker.py
───────────────
Unlocks table row heights from exact to atLeast across all tables (including nested tables),
enforces cantSplit row protection, sets fixed table layout, and normalizes cell padding.
"""

from typing import List, Optional
import docx
from docx.table import Table, _Cell, _Row
from docx.oxml import OxmlElement
from docx.oxml.ns import qn, nsdecls


def unlock_table_rows(
    doc: docx.Document,
    fixed_layout: bool = True,
    cant_split: bool = True,
    margin_dxa: int = 20
) -> int:
    """
    Scans all tables (including nested tables) in the document:
    1. Converts <w:trHeight w:hRule="exact"/> to w:hRule="atLeast" so rows expand organically.
    2. Injects <w:tblLayout w:type="fixed"/> on all tables to prevent border drift.
    3. Injects <w:cantSplit/> on all rows to prevent awkward page-boundary splits.
    4. Normalizes cell margins (<w:tcMar>) to prevent horizontal text clipping.
    
    Returns the total number of unlocked rows.
    """
    unlocked_count = 0
    # Use XPath to find all tables including nested tables
    all_tbl_elements = doc.element.xpath("//w:tbl")

    for tbl_el in all_tbl_elements:
        tbl = Table(tbl_el, doc)

        # 1. Fixed layout
        if fixed_layout:
            tblPr = tbl._tbl.tblPr
            layout_el = tblPr.find(qn("w:tblLayout"))
            if layout_el is None:
                layout_el = OxmlElement("w:tblLayout")
                tblPr.append(layout_el)
            layout_el.set(qn("w:type"), "fixed")

        # 2. Iterate rows
        for row in tbl.rows:
            trPr = row._tr.get_or_add_trPr()

            # Unlock trHeight exact -> atLeast
            trHeight = trPr.find(qn("w:trHeight"))
            if trHeight is not None:
                h_rule = trHeight.get(qn("w:hRule"))
                if h_rule == "exact":
                    trHeight.set(qn("w:hRule"), "atLeast")
                    unlocked_count += 1
            else:
                # If no trHeight, ensure default reflow at least
                pass

            # Inject cantSplit
            if cant_split:
                cantSplit = trPr.find(qn("w:cantSplit"))
                if cantSplit is None:
                    cantSplit = OxmlElement("w:cantSplit")
                    trPr.append(cantSplit)

            # 3. Cell margins
            if margin_dxa is not None:
                for cell in row.cells:
                    tcPr = cell._tc.get_or_add_tcPr()
                    tcMar = tcPr.find(qn("w:tcMar"))
                    if tcMar is None:
                        tcMar = OxmlElement("w:tcMar")
                        tcPr.append(tcMar)
                    for side in ("top", "bottom", "left", "right"):
                        m = tcMar.find(qn(f"w:{side}"))
                        if m is None:
                            m = OxmlElement(f"w:{side}")
                            tcMar.append(m)
                        m.set(qn("w:w"), str(margin_dxa))
                        m.set(qn("w:type"), "dxa")

    return unlocked_count
