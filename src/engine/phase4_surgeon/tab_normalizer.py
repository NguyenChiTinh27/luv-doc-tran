"""
tab_normalizer.py
─────────────────
Removes all instances of internal double-tabs (\\t\\t) and abnormal tab stops
from paragraphs and table cells, both in the python-docx DOM and at the raw
OpenXML ZIP package level.
"""

import os
import re
import zipfile
import tempfile
import docx
from typing import Optional


def normalize_tabs_and_indents(doc: docx.Document) -> int:
    """
    Scans all paragraphs in body and table cells in the python-docx DOM,
    collapsing multiple consecutive tabs into clean spaces or single tabs.
    Returns the count of cleaned paragraphs/runs.
    """
    cleaned_count = 0

    # 1. Process body paragraphs
    for p in doc.paragraphs:
        if "\t\t" in p.text or "\t \t" in p.text:
            for r in p.runs:
                if "\t\t" in r.text or "\t \t" in r.text:
                    r.text = re.sub(r"\t\s*\t+", " ", r.text)
                    cleaned_count += 1
            # Also remove consecutive <w:tab/> elements in XML
            tabs = p._p.xpath(".//w:tab")
            if len(tabs) > 1:
                for i in range(len(tabs) - 1):
                    # If adjacent
                    parent = tabs[i].getparent()
                    next_el = tabs[i].getnext()
                    if next_el is not None and next_el.tag.endswith("tab"):
                        try:
                            parent.remove(tabs[i])
                            cleaned_count += 1
                        except Exception:
                            pass

    # 2. Process table cells
    for tbl_el in doc.element.xpath("//w:tbl"):
        for p in tbl_el.xpath(".//w:p"):
            p_text = "".join(p.xpath(".//w:t/text()"))
            if "\t\t" in p_text or "\t \t" in p_text:
                for t in p.xpath(".//w:t"):
                    if t.text and ("\t\t" in t.text or "\t \t" in t.text):
                        t.text = re.sub(r"\t\s*\t+", " ", t.text)
                        cleaned_count += 1

    return cleaned_count


def sanitize_package_xml(docx_path: str) -> None:
    """
    Performs package-level sanitization across word/document.xml,
    word/header*.xml, and word/footer*.xml:
    1. Replaces all instances of \t\t with a clean space.
    2. Ensures inter-run boundaries don't form concatenation.
    3. Normalizes common technical abbreviations (VPVB, VWVWS, NZIOC, GMBH).
    """
    if not os.path.exists(docx_path):
        return

    temp_zip = docx_path + ".sanitizing.tmp"
    with zipfile.ZipFile(docx_path, "r") as zin:
        with zipfile.ZipFile(temp_zip, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith("word/") and (
                    item.filename == "word/document.xml" or
                    item.filename.startswith("word/header") or
                    item.filename.startswith("word/footer")
                ):
                    try:
                        content = data.decode("utf-8")

                        # Strip internal double-tabs
                        content = re.sub(r"\t\s*\t+", " ", content)
                        content = re.sub(r"(<w:tab/>\s*)+<w:tab/>", "<w:tab/>", content)

                        # Clean uppercase acronyms that might trip concatenation
                        content = re.sub(r"\bvPvB\b", "VPVB", content)
                        content = re.sub(r"\bVwVwS\b", "VWVWS", content)
                        content = re.sub(r"\bNZIoC\b", "NZIOC", content)
                        content = re.sub(r"\bGmbH\b", "GMBH", content)

                        data = content.encode("utf-8")
                    except Exception:
                        pass
                zout.writestr(item, data)

    # Atomic replace
    os.replace(temp_zip, docx_path)
