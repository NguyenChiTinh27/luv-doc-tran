"""
Phase 4 Surgeon Module: OpenXML deep surgical modifications, row height unlocking,
native header/footer promotion, tab normalizer, and anti-overflow line compactor.
"""

from src.engine.phase4_surgeon.row_unlocker import unlock_table_rows
from src.engine.phase4_surgeon.header_footer_promoter import promote_headers_footers
from src.engine.phase4_surgeon.tab_normalizer import normalize_tabs_and_indents, sanitize_package_xml
from src.engine.phase4_surgeon.layout_compactor import apply_line_compaction
from src.engine.phase4_surgeon.surgeon import OpenXmlSurgeon

__all__ = [
    "unlock_table_rows",
    "promote_headers_footers",
    "normalize_tabs_and_indents",
    "sanitize_package_xml",
    "apply_line_compaction",
    "OpenXmlSurgeon",
]
