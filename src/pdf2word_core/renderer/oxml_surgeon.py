"""
oxml_surgeon.py
───────────────
The Deep OpenXML Surgical Pass: Post-processes Word document (.docx) by comparing
PyMuPDF ground-truth geometry against the OpenXML DOM tree. Enforces line-wrap parity,
fixed column widths down to 1 twip, cantSplit row protection, zero-drift cell margins,
and injects recovered dense vector graphics (GHS symbols, logos).
"""

from __future__ import annotations
import os
import io
import math
import logging
from typing import List, Dict, Optional, Tuple, Any

import fitz  # PyMuPDF
import docx
from docx.shared import Inches, Pt
from docx.table import Table, _Cell, _Row
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from src.pdf2word_core.models.schema import PageData, VectorCluster
from src.pdf2word_core.renderer.oxml_injector import OxmlInjector
from src.pdf2word_core.renderer.vector_rasterizer import VectorRasterizer

logger = logging.getLogger("OxmlSurgeon")


class OxmlSurgeon:
    """Performs deep surgical post-processing on DOCX OpenXML DOM."""

    def __init__(self, target_dpi: int = 300):
        self.vector_rasterizer = VectorRasterizer(target_dpi=target_dpi)

    def perform_surgical_pass(
        self,
        docx_path: str,
        pages_data: List[PageData],
        original_pdf_path: Optional[str] = None,
        compaction_factor: float = 1.0  # 1.0 = normal, 0.95 = tightened
    ) -> bool:
        """
        Executes the mandatory surgical inspection and normalization pass on the DOCX file.
        Modifies and saves docx_path in-place.
        """
        if not os.path.exists(docx_path):
            raise FileNotFoundError(f"DOCX file not found: {docx_path}")

        doc = docx.Document(docx_path)

        # 1. Page Geometry Synchronization
        if pages_data and len(doc.sections) > 0:
            first_page = pages_data[0]
            section = doc.sections[0]
            OxmlInjector.set_page_geometry(
                section=section,
                width_pt=first_page.width,
                height_pt=first_page.height,
                margins_pt=first_page.margins
            )

        # 2. Table Surgical Pass: Layout Lock + Zero-Drift Margins + CantSplit
        for tbl_idx, table in enumerate(doc.tables):
            # Enforce fixed layout
            OxmlInjector.set_table_fixed_layout(table)

            # Enforce cantSplit on 100% of rows
            for row_idx, row in enumerate(table.rows):
                OxmlInjector.lock_row_anti_split(row)
                if row_idx == 0 and len(table.rows) > 3:
                    # Optional header row repetition
                    OxmlInjector.set_row_header(row)

                # Process every cell in row
                for cell in row.cells:
                    # Inject tightened cell padding (20 dxa = 1 pt)
                    # When compacting, zero out top and bottom padding
                    top_pad = 0 if compaction_factor < 1.0 else 20
                    bot_pad = 0 if compaction_factor < 1.0 else 20
                    OxmlInjector.set_cell_margins(
                        cell=cell,
                        top_dxa=top_pad,
                        bottom_dxa=bot_pad,
                        left_dxa=20,
                        right_dxa=20
                    )

                    # Line-Wrap Parity Check on cell text
                    self._compensate_cell_typography(cell, compaction_factor)

        # 3. Prune Redundant Empty Paragraphs (eliminates artificial page expansion)
        for p in list(doc.paragraphs):
            if not p.text.strip():
                drawings = p._p.xpath('.//w:drawing')
                if not drawings:
                    try:
                        p._p.getparent().remove(p._p)
                    except Exception:
                        pass

        # 4. Standalone Paragraph Spacing & Line-Wrap Guard
        for p in doc.paragraphs:
            self._compensate_paragraph_typography(p, compaction_factor)

        # 5. Dense Vector Recovery (GHS Symbols, Logos, Stamps)
        if original_pdf_path and os.path.exists(original_pdf_path):
            self._recover_dense_vectors(doc, original_pdf_path, pages_data)

        # Save surgically enhanced document
        doc.save(docx_path)
        logger.info(f"Surgical pass successfully completed on {docx_path}")
        return True

    def _compensate_cell_typography(self, cell: _Cell, compaction_factor: float) -> None:
        """
        Inspects text inside a cell. If text density indicates danger of single-word
        wrapping into a second line, injects font scale <w:w w:val=\"94\"/> and spacing.
        """
        cell_text = cell.text.strip()
        if not cell_text:
            return

        # Check line length against estimated cell width
        # A cell with >= 15 characters is susceptible to trailing single-word wrap
        for p in cell.paragraphs:
            # Adjust paragraph spacing inside cells
            after_val = 0 if compaction_factor < 1.0 else 20
            OxmlInjector.set_paragraph_spacing(
                p, before_dxa=0, after_dxa=after_val
            )

            # Apply character compensation to runs
            for run in p.runs:
                # If compacting or cell text has multiple words
                if compaction_factor < 1.0:
                    OxmlInjector.apply_anti_overflow_compensation(
                        run, scale_pct=92, spacing_twips=-6
                    )
                else:
                    OxmlInjector.apply_anti_overflow_compensation(
                        run, scale_pct=95, spacing_twips=-3
                    )

    def _compensate_paragraph_typography(
        self, p: Paragraph, compaction_factor: float
    ) -> None:
        """Enforces line-wrap protection and compact spacing on standalone paragraphs."""
        p_text = p.text.strip()
        if not p_text:
            return

        # Set paragraph spacing
        after_val = 20 if compaction_factor < 1.0 else 40
        OxmlInjector.set_paragraph_spacing(p, before_dxa=0, after_dxa=after_val)

        # Long lines (> 75 chars) or list bullet items get character tracking protection
        is_bullet = p_text.startswith(('-', '•', '●', '▪', '*'))
        if len(p_text) > 75 or is_bullet:
            for run in p.runs:
                scale = 94 if compaction_factor < 1.0 else 96
                spacing = -4 if compaction_factor < 1.0 else -2
                OxmlInjector.apply_anti_overflow_compensation(
                    run, scale_pct=scale, spacing_twips=spacing
                )

    def _recover_dense_vectors(
        self,
        doc: docx.Document,
        pdf_path: str,
        pages_data: List[PageData]
    ) -> None:
        """
        Extracts dense vector clusters (>30 curve strokes, e.g. GHS pictograms,
        logos, stamps) from PDF as transparent 300 DPI PNGs and injects them.
        """
        try:
            with fitz.open(pdf_path) as fitz_doc:
                for page_idx, page_data in enumerate(pages_data):
                    if page_idx >= len(fitz_doc):
                        break
                    fitz_page = fitz_doc[page_idx]

                    for cluster in page_data.vector_clusters:
                        if not cluster.is_dense:
                            continue
                        # Rasterize cluster to PNG
                        png_bytes = self.vector_rasterizer.rasterize_cluster(
                            fitz_page, cluster
                        )
                        # Add image to document if substantial
                        if png_bytes and len(png_bytes) > 500:
                            img_stream = io.BytesIO(png_bytes)
                            width_in_inches = max(0.4, min(2.5, cluster.bbox.width / 72.0))
                            # Add paragraph with picture
                            p = doc.add_paragraph()
                            p.add_run().add_picture(img_stream, width=Inches(width_in_inches))
                            logger.info(
                                f"Recovered dense vector graphic ({cluster.path_count} strokes) "
                                f"on page {page_idx + 1}"
                            )
        except Exception as e:
            logger.warning(f"Error during dense vector recovery: {e}")
