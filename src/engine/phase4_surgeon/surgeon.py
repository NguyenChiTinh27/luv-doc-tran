"""
surgeon.py
──────────
Phase 4 OpenXML Deep Surgeon: orchestrates row height unlocking, header/footer
promotion, tab/indent normalization, anti-overflow character compaction, and
OpenXML package sanitization.
"""

from __future__ import annotations
import os
import shutil
from typing import Optional, Dict, Any
import docx

from src.engine.phase4_surgeon.row_unlocker import unlock_table_rows
from src.engine.phase4_surgeon.header_footer_promoter import promote_headers_footers
from src.engine.phase4_surgeon.tab_normalizer import normalize_tabs_and_indents, sanitize_package_xml
from src.engine.phase4_surgeon.layout_compactor import apply_line_compaction


class OpenXmlSurgeon:
    """Master orchestrator for Phase 4 OpenXML surgery."""

    def __init__(self, target_dpi: int = 300):
        self.target_dpi = target_dpi

    def process_document(
        self,
        docx_path: str,
        output_path: Optional[str] = None,
        compaction_factor: float = 1.0,
        title_text: Optional[str] = None,
        remove_body_pagination: bool = True
    ) -> str:
        """
        Executes the deep surgical pass on the DOCX file:
        1. Opens the DOCX DOM.
        2. Unlocks table row heights from exact to atLeast across all nested tables.
        3. Promotes recurring headers & pagination to native Word Sections.
        4. Strips internal double-tabs and normalizes table indents.
        5. If compaction_factor < 1.0, applies line character compaction.
        6. Saves DOCX and runs package-level XML sanitization.
        """
        if not os.path.exists(docx_path):
            raise FileNotFoundError(f"Input DOCX not found: {docx_path}")

        target_path = output_path or docx_path
        if target_path != docx_path:
            shutil.copy2(docx_path, target_path)

        doc = docx.Document(target_path)

        # 1. Row Unlocker
        unlock_table_rows(doc, fixed_layout=True, cant_split=True, margin_dxa=20)

        # 2. Header & Footer Promotion
        promote_headers_footers(doc, title_text=title_text, remove_body_pagination=remove_body_pagination)

        # 3. Tab Normalization
        normalize_tabs_and_indents(doc)

        # 4. Anti-Overflow Compaction
        if compaction_factor < 1.0:
            apply_line_compaction(doc, compaction_factor=compaction_factor)

        # Save changes to target path
        doc.save(target_path)

        # 5. Package-level XML Sanitization
        sanitize_package_xml(target_path)

        return target_path
