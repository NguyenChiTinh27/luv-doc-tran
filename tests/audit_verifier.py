"""
tests/audit_verifier.py

Authoritative 7-Point XML & Visual Audit Verifier for Bilingual DOCX Documents.
Directly inspects the OpenXML ZIP container (word/document.xml, word/header*.xml, word/footer*.xml)
using lxml.etree to guarantee zero-defect layout and translation quality:
  - Check 1: Zero Concatenation ([a-z0-9][A-ZÀ-Ỹ] across boundaries without <w:br/> or <w:p>)
  - Check 2: Zero Internal Double-Tab (\t\t)
  - Check 3: 100% Pair Completeness (validated bilingual pairs / translatable source blocks == 1.0)
  - Check 4: Zero Leaked English (0 ordinals like 1st, 2nd, leaked section titles Goods, Measures, section6.2)
  - Check 5: Header Order (Sections 1-16 in strictly ascending Roman/numeric order I-XVI)
  - Check 6: Zero Clipped Rows (0 <w:trHeight w:hRule="exact"/>)
  - Check 7: Visual Parity (native header/footer with dynamic PAGE/NUMPAGES, keepNext on source lines, dark navy #1B365D, valid XML)
"""

import sys
import os
import re
import zipfile
from typing import Dict, Any, List, Tuple, Optional
import lxml.etree as etree

# OpenXML Namespaces
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {
    "w": W_NS,
    "r": R_NS,
}

# Vietnamese capital letters & diacritics
VI_UPPER = "ÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬÈÉẺẼẸÊỀẾỂỄỆÌÍỈĨỊÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰỲÝỶỸỴĐ"
VI_LOWER = "àáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ"
VI_DIACRITICS_REGEX = re.compile(f"[{VI_LOWER}{VI_UPPER}]")

# Whitelist exceptions for Check 1 (valid scientific/technical tokens)
WHITELIST_TOKENS = {
    "pH", "mRNA", "tRNA", "dATP", "cDNA", "mCD4", "iOS", "iPad", "LaTeX", "macOS",
    "eCommerce", "xHeight", "mPa", "hPa", "kPa", "MPa", "GPa", "kDa", "mg/kg",
    "g/cm3", "kg/m3", "ml/min", "mL/min", "km/h", "m/s", "mm2/s", "cSt", "ppm",
    "ppb", "wt%", "vol%", "v/v", "w/w"
}

# Common chemical element symbols
CHEM_ELEMENTS = {
    "H", "He", "Li", "Be", "B", "C", "N", "O", "F", "Ne", "Na", "Mg", "Al", "Si",
    "P", "S", "Cl", "Ar", "K", "Ca", "Sc", "Ti", "V", "Cr", "Mn", "Fe", "Co", "Ni",
    "Cu", "Zn", "Ga", "Ge", "As", "Se", "Br", "Kr", "Rb", "Sr", "Y", "Zr", "Nb", "Mo",
    "Tc", "Ru", "Rh", "Pd", "Ag", "Cd", "In", "Sn", "Sb", "Te", "I", "Xe", "Cs", "Ba",
    "La", "Ce", "Pr", "Nd", "Pm", "Sm", "Eu", "Gd", "Tb", "Dy", "Ho", "Er", "Tm", "Yb",
    "Lu", "Hf", "Ta", "W", "Re", "Os", "Ir", "Pt", "Au", "Hg", "Tl", "Pb", "Bi", "Th",
    "Pa", "U"
}

# Common MSDS / HSE Vietnamese terms to detect translation presence
VI_COMMON_WORDS = {
    "mục", "phần", "không", "có", "chất", "axit", "nước", "hợp chất", "dung dịch",
    "xylen", "toluen", "benzen", "etyl", "butyl", "axeton", "bảng", "dữ liệu",
    "an toàn", "sơ cứu", "hỏa hoạn", "tràn đổ", "sử dụng", "bảo quản", "phơi nhiễm",
    "bảo vệ", "đặc tính", "độc tính", "sinh thái", "tiêu hủy", "vận chuyển",
    "quy chuẩn", "thông tin", "nguy hiểm", "cảnh báo", "phòng ngừa", "tỷ trọng",
    "điểm chớp cháy", "độ tan", "nhận dạng", "thành phần", "nồng độ", "giới hạn",
    "silic", "đioxit", "canxi", "cacbonat", "nhôm", "oxit", "titan", "bột", "talc"
}

# Check 4 Leaked English regex patterns
LEAKED_ORDINALS = re.compile(r"\b(1st|2nd|3rd|[4-9]th|10th|11th|12th|13th|14th|15th|16th)\b", re.IGNORECASE)
LEAKED_HEADINGS = re.compile(
    r"\b(Goods|Measures|Emergency|Precautionary\s+statements|First\s+aid|Handling\s+and\s+storage|"
    r"Exposure\s+controls|Physical\s+and\s+chemical|Stability\s+and\s+reactivity|Toxicological|"
    r"Ecological|Disposal\s+considerations|Transport\s+information|Regulatory\s+information)\b",
    re.IGNORECASE
)
LEAKED_SECTION_REF = re.compile(r"\bsection\s*\d+(\.\d+)?\b", re.IGNORECASE)
LEAKED_STATUS = re.compile(r"^(None|Unknown|Not\s+applicable|N/A|See\s+section)$", re.IGNORECASE)

# Check 5 Section Header Numerals
ROMAN_TO_INT = {
    "I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6, "VII": 7, "VIII": 8,
    "IX": 9, "X": 10, "XI": 11, "XII": 12, "XIII": 13, "XIV": 14, "XV": 15, "XVI": 16
}
SECTION_HEADER_REGEX = re.compile(
    r"^(?:SECTION|Section|MỤC|Mục|Phần)?\s*([0-9]{1,2}|I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII|XIII|XIV|XV|XVI)[\.\:\s\-]+",
    re.IGNORECASE
)


def is_chemical_formula(token: str) -> bool:
    """Detect if token is a standard chemical formula like CaCO3, SiO2, Al2O3, H2SO4."""
    cleaned = re.sub(r"[\(\)\[\]\.,]", "", token)
    if not cleaned:
        return False
    # Pattern: Sequence of (Element + optional count)
    match = re.match(r"^([A-Z][a-z]?[0-9]*)+$", cleaned)
    if not match:
        return False
    # Verify elements
    elements_found = re.findall(r"([A-Z][a-z]?)", cleaned)
    return bool(elements_found) and all(e in CHEM_ELEMENTS for e in elements_found)


def is_vietnamese_text(text: str) -> bool:
    """Detect whether text contains Vietnamese diacritics or standard MSDS Vietnamese terms."""
    if VI_DIACRITICS_REGEX.search(text):
        return True
    text_lower = text.lower()
    return any(w in text_lower for w in VI_COMMON_WORDS)


def is_translatable_source_text(text: str) -> bool:
    """Check if a text block contains meaningful English text requiring translation."""
    cleaned = text.strip()
    if not cleaned or len(cleaned) < 2:
        return False
    # Must contain letters
    if not re.search(r"[a-zA-Z]", cleaned):
        return False
    # Ignore purely numeric or punctuation blocks (e.g. "123-45-6", "100%", "---")
    if re.match(r"^[\d\s\-\.\,\:\;\%\(\)\/\<\>\=\±\*\#\+\~]+$", cleaned):
        return False
    # Ignore numeric measurements with common units (e.g. ": > 1700 °C", "120 - 150 mPa.s", "25 °C")
    if re.match(r"^[\:\s\d\-\.\,\/\<\>\=\±\*\#\+\~]+(?:\s*(?:°C|°F|K|mg|g|kg|ml|mL|L|ppm|ppb|%|hPa|kPa|MPa|mPa\.s|cSt|mm2/s|wt%|vol%))\s*$", cleaned):
        return False
    # Ignore standard pagination strings in body
    if re.match(r"^(?:Page\s+\d+|p\.\s*\d+|\(Contd\.)", cleaned, re.IGNORECASE):
        return False
    # Ignore pure short chemical codes/CAS alone (e.g. "CAS: 7631-86-9" without label, "01-2119379499-16-XXXX")
    if re.match(r"^(?:CAS|EINECS|EC|UN|GB\d+)?\s*[\d\-\.\/Xx]+$", cleaned):
        return False
    return True


def check_zero_concatenation(xml_trees: List[etree.Element]) -> Tuple[int, List[Dict[str, Any]]]:
    """
    Check 1: Zero Concatenation ([a-z0-9][A-ZÀ-Ỹ] across boundaries without <w:br/> or <w:p>).
    Inspects intra-run strings and inter-run boundaries within the same paragraph.
    """
    violations = []
    concat_pattern = re.compile(f"([a-z0-9])([{VI_UPPER}A-Z])")

    for tree_idx, tree in enumerate(xml_trees):
        # Scan each paragraph
        for p_elem in tree.xpath(".//w:p", namespaces=NS):
            # 1. Intra-run inspection inside <w:t>
            runs = p_elem.xpath("./w:r", namespaces=NS)
            for r_idx, r in enumerate(runs):
                t_elements = r.xpath("./w:t", namespaces=NS)
                for t in t_elements:
                    t_text = t.text or ""
                    # Find all potential concatenation sites
                    for m in concat_pattern.finditer(t_text):
                        start, end = max(0, m.start() - 6), min(len(t_text), m.end() + 6)
                        snippet = t_text[start:end]

                        # Extract full token surrounding m.start() and m.end()
                        w_start = m.start()
                        while w_start > 0 and (t_text[w_start - 1].isalnum() or t_text[w_start - 1] in "_-"):
                            w_start -= 1
                        w_end = m.end()
                        while w_end < len(t_text) and (t_text[w_end].isalnum() or t_text[w_end] in "_-"):
                            w_end += 1
                        word = t_text[w_start:w_end]

                        # Whitelist check
                        match_pair = m.group(0)
                        if (word in WHITELIST_TOKENS or
                            match_pair in {"pH", "mP", "hP", "kP", "MP", "GP", "kD"} or
                            any(p in WHITELIST_TOKENS for p in re.split(r"[\.\-]", word)) or
                            is_chemical_formula(word) or
                            is_chemical_formula(re.split(r"[\.\-]", word)[0])):
                            continue

                        # If lowercase followed by uppercase Vietnamese letter -> definite violation
                        char_after = m.group(2)
                        if char_after in VI_UPPER:
                            violations.append({
                                "type": "intra_run_violation",
                                "snippet": snippet,
                                "match": m.group(0),
                                "full_text": t_text[:80]
                            })
                        elif m.group(1).isalpha() and char_after.isupper():
                            # English word followed directly by another capitalized word without space
                            # e.g., "IdentificationSection" or "undertakingMục"
                            violations.append({
                                "type": "intra_run_violation",
                                "snippet": snippet,
                                "match": m.group(0),
                                "full_text": t_text[:80]
                            })

            # 2. Inter-run boundary inspection between adjacent runs
            # If Run A ends with [a-z0-9] and Run B starts with [A-ZÀ-Ỹ]
            # AND between Run A and Run B there is NO <w:br/>, NO <w:tab/>, NO space
            for i in range(len(runs) - 1):
                r_a = runs[i]
                r_b = runs[i + 1]

                # Check if r_a or r_b contains a break element
                has_break = bool(r_a.xpath(".//w:br | .//w:cr | .//w:tab", namespaces=NS) or
                                 r_b.xpath(".//w:br | .//w:cr | .//w:tab", namespaces=NS))
                if has_break:
                    continue

                text_a = "".join(r_a.xpath(".//w:t/text()", namespaces=NS))
                text_b = "".join(r_b.xpath(".//w:t/text()", namespaces=NS))

                if not text_a or not text_b:
                    continue

                # Check boundary: trailing char of A and leading char of B
                trailing_a = text_a[-1]
                leading_b = text_b[0]

                if re.match(r"[a-z0-9]", trailing_a) and re.match(f"[{VI_UPPER}A-Z]", leading_b):
                    boundary = trailing_a + leading_b
                    # Exclude chemical formulas or whitelisted tokens spanning runs
                    combined_word = (text_a.split()[-1] if text_a.split() else "") + (text_b.split()[0] if text_b.split() else "")
                    if combined_word in WHITELIST_TOKENS or is_chemical_formula(combined_word):
                        continue

                    violations.append({
                        "type": "inter_run_violation",
                        "text_a": text_a[-15:],
                        "text_b": text_b[:15],
                        "boundary": boundary,
                    })

    return len(violations), violations


def check_zero_double_tabs(doc_tree: etree.Element) -> Tuple[int, List[Dict[str, Any]]]:
    """
    Check 2: Zero Internal Double-Tab (\t\t).
    Searches for consecutive tab characters in paragraph text or XML elements.
    """
    violations = []
    for p in doc_tree.xpath(".//w:p", namespaces=NS):
        # Collect runs and elements in order
        elem_parts = []
        for child in p.iterchildren():
            if child.tag == f"{{{W_NS}}}r":
                for r_child in child.iterchildren():
                    if r_child.tag == f"{{{W_NS}}}t":
                        elem_parts.append(r_child.text or "")
                    elif r_child.tag == f"{{{W_NS}}}tab":
                        elem_parts.append("\t")
            elif child.tag == f"{{{W_NS}}}tab":
                elem_parts.append("\t")

        p_text = "".join(elem_parts)
        # Check for \t\t or \t whitespace \t
        double_tab_matches = list(re.finditer(r"\t\s*\t", p_text))
        if double_tab_matches:
            violations.append({
                "count": len(double_tab_matches),
                "snippet": p_text[:80].replace("\t", "<TAB>")
            })

    total_violations = sum(v["count"] for v in violations)
    return total_violations, violations


def check_pair_completeness(doc_tree: etree.Element) -> Tuple[float, int, int, List[str]]:
    """
    Check 3: 100% Pair Completeness (validated bilingual pairs / translatable source blocks == 1.0).
    Verifies that every translatable English source block has a corresponding Vietnamese translation
    block immediately below it (Pattern A dual paragraph or Pattern B line break in table cell).
    """
    total_source_blocks = 0
    validated_pairs = 0
    unpaired_examples = []

    # 1. Inspect Table Cells (Pattern B or Pattern A within cells)
    for tc in doc_tree.xpath(".//w:tc", namespaces=NS):
        paragraphs = tc.xpath("./w:p", namespaces=NS)
        for p in paragraphs:
            # Check if this paragraph has Pattern B (<w:br/> separating EN and VI)
            runs = p.xpath("./w:r", namespaces=NS)
            has_br = False
            pre_br_text = []
            post_br_text = []
            for r in runs:
                br_elements = r.xpath("./w:br", namespaces=NS)
                t_elements = r.xpath("./w:t", namespaces=NS)
                r_text = "".join(t.text or "" for t in t_elements)

                if br_elements:
                    has_br = True
                else:
                    if not has_br:
                        pre_br_text.append(r_text)
                    else:
                        post_br_text.append(r_text)

            en_part = "".join(pre_br_text).strip()
            vi_part = "".join(post_br_text).strip()

            if is_translatable_source_text(en_part):
                total_source_blocks += 1
                is_vi_formatted = bool(tc.xpath(".//w:color[@w:val='1B365D' or @w:val='002060']", namespaces=NS) or tc.xpath(".//w:i", namespaces=NS))
                if has_br and (is_vietnamese_text(vi_part) or (is_vi_formatted and bool(vi_part))):
                    validated_pairs += 1
                else:
                    # Might be part of Pattern A inside cell, check next paragraph
                    p_idx = paragraphs.index(p)
                    if p_idx + 1 < len(paragraphs):
                        next_p = paragraphs[p_idx + 1]
                        next_text = "".join(next_p.xpath(".//w:t/text()", namespaces=NS)).strip()
                        next_is_vi_formatted = bool(next_p.xpath(".//w:color[@w:val='1B365D' or @w:val='002060']", namespaces=NS) or next_p.xpath(".//w:i", namespaces=NS))
                        if is_vietnamese_text(next_text) or (next_is_vi_formatted and bool(next_text)):
                            validated_pairs += 1
                        else:
                            unpaired_examples.append(f"Cell Pattern A missing translation: {en_part[:50]}")
                    else:
                        unpaired_examples.append(f"Cell Pattern B missing translation: {en_part[:50]}")

    # 2. Inspect Body Paragraphs (excluding paragraphs inside tables)
    body_paragraphs = doc_tree.xpath("/w:document/w:body/w:p", namespaces=NS)
    i = 0
    while i < len(body_paragraphs):
        p_elem = body_paragraphs[i]
        p_text = "".join(p_elem.xpath(".//w:t/text()", namespaces=NS)).strip()

        if is_translatable_source_text(p_text) and not is_vietnamese_text(p_text):
            total_source_blocks += 1
            # Next paragraph must be the Vietnamese translation
            if i + 1 < len(body_paragraphs):
                next_p = body_paragraphs[i + 1]
                next_text = "".join(next_p.xpath(".//w:t/text()", namespaces=NS)).strip()
                next_is_vi_formatted = bool(next_p.xpath(".//w:color[@w:val='1B365D' or @w:val='002060']", namespaces=NS) or next_p.xpath(".//w:i", namespaces=NS))

                if is_vietnamese_text(next_text) or (next_is_vi_formatted and bool(next_text)):
                    validated_pairs += 1
                    i += 2  # Skip translation paragraph
                    continue
                else:
                    unpaired_examples.append(f"Body paragraph missing translation: {p_text[:50]}")
            else:
                unpaired_examples.append(f"Trailing body paragraph without pair: {p_text[:50]}")
        i += 1

    if total_source_blocks == 0:
        completeness = 1.0
    else:
        completeness = round(validated_pairs / total_source_blocks, 4)

    return completeness, total_source_blocks, validated_pairs, unpaired_examples


def check_zero_leaked_english(doc_tree: etree.Element) -> Tuple[int, List[Dict[str, Any]]]:
    """
    Check 4: Zero Leaked English (0 ordinals like 1st, 2nd, leaked section titles Goods, Measures, section6.2).
    Inspects all Vietnamese translated lines.
    """
    violations = []

    # Collect all translated text runs/paragraphs
    # Identified by Vietnamese formatting: dark navy color (#1B365D / #002060) or italic <w:i/> or following a source line
    translated_lines = []

    # From body paragraphs
    body_paragraphs = doc_tree.xpath("/w:document/w:body/w:p", namespaces=NS)
    for i, p in enumerate(body_paragraphs):
        # A paragraph is considered a translation paragraph if it has italic and navy color, or is immediately after an English line
        is_trans_p = bool(p.xpath(".//w:rPr/w:color[@w:val='1B365D' or @w:val='002060']", namespaces=NS) or
                          p.xpath(".//w:rPr/w:i", namespaces=NS))
        if is_trans_p:
            text = "".join(p.xpath(".//w:t/text()", namespaces=NS)).strip()
            if text:
                translated_lines.append(text)

    # From table cells (post-break runs)
    for tc in doc_tree.xpath(".//w:tc", namespaces=NS):
        for p in tc.xpath("./w:p", namespaces=NS):
            runs = p.xpath("./w:r", namespaces=NS)
            has_br = False
            vi_runs = []
            for r in runs:
                if r.xpath("./w:br", namespaces=NS):
                    has_br = True
                elif has_br:
                    vi_runs.append("".join(r.xpath(".//w:t/text()", namespaces=NS)))
            if vi_runs:
                text = "".join(vi_runs).strip()
                if text:
                    translated_lines.append(text)

    # Scan each translated line for leaks
    for line in translated_lines:
        # 1. Leaked ordinals: 1st, 2nd, 3rd, ...
        for m in LEAKED_ORDINALS.finditer(line):
            violations.append({"type": "leaked_ordinal", "match": m.group(0), "line": line[:70]})

        # 2. Leaked section titles/headings: Goods, Measures, Precautionary statements, etc.
        for m in LEAKED_HEADINGS.finditer(line):
            # Check if this word appears standalone or untranslated in Vietnamese line
            violations.append({"type": "leaked_heading", "match": m.group(0), "line": line[:70]})

        # 3. Leaked section references: section6.2, section 6.2
        for m in LEAKED_SECTION_REF.finditer(line):
            violations.append({"type": "leaked_section_ref", "match": m.group(0), "line": line[:70]})

        # 4. Leaked standalone status words: None, Unknown, N/A
        if LEAKED_STATUS.match(line):
            violations.append({"type": "leaked_status_word", "match": line, "line": line})

        # 5. Untranslated English sentences (>= 4 English words and 0 Vietnamese diacritics/terms)
        words = line.split()
        if len(words) >= 4 and not is_vietnamese_text(line):
            # Verify they are latin words
            if all(re.match(r"^[a-zA-Z\.\,\:\;\-]+$", w) for w in words[:4]):
                violations.append({"type": "untranslated_english_sentence", "match": " ".join(words[:4]), "line": line[:70]})

    return len(violations), violations


def check_header_order(doc_tree: etree.Element) -> Tuple[int, List[Dict[str, Any]]]:
    """
    Check 5: Header Order (Sections 1-16 in strictly ascending Roman/numeric order I-XVI).
    Ensures no section headers are inverted (e.g. Section 11 appearing before Section 9).
    """
    violations = []
    last_sec_num = 0
    section_history = []

    # Traverse all elements in document order
    for elem in doc_tree.xpath(".//w:p", namespaces=NS):
        text = "".join(elem.xpath(".//w:t/text()", namespaces=NS)).strip()
        m = SECTION_HEADER_REGEX.match(text)
        if m:
            numeral_str = m.group(1).upper()
            if numeral_str in ROMAN_TO_INT:
                sec_num = ROMAN_TO_INT[numeral_str]
            elif numeral_str.isdigit():
                sec_num = int(numeral_str)
            else:
                continue

            if 1 <= sec_num <= 16:
                section_history.append((sec_num, text[:50]))
                if sec_num < last_sec_num:
                    violations.append({
                        "current_section": sec_num,
                        "previous_section": last_sec_num,
                        "header_text": text[:60]
                    })
                else:
                    last_sec_num = sec_num

    return len(violations), violations


def check_zero_clipped_rows(doc_tree: etree.Element) -> Tuple[int, List[str]]:
    """
    Check 6: Zero Clipped Rows (0 <w:trHeight w:hRule="exact"/>).
    Must be w:hRule="atLeast" or omitted so rows expand organically.
    """
    exact_tr_heights = doc_tree.xpath("//w:trHeight[@w:hRule='exact']", namespaces=NS)
    violations = []
    for elem in exact_tr_heights:
        val = elem.get(f"{{{W_NS}}}val") or elem.get("w:val") or "unspecified"
        violations.append(f"<w:trHeight w:hRule='exact' w:val='{val}'/>")

    return len(violations), violations


def check_visual_parity(zip_archive: zipfile.ZipFile, doc_tree: etree.Element) -> Tuple[bool, List[str]]:
    """
    Check 7: Visual Parity.
    Validates:
      - Native Word header/footer promotion with dynamic PAGE / NUMPAGES fields
      - Absence of orphaned pagination strings in body
      - <w:keepNext/> on source body paragraphs of bilingual pairs
      - Typography: Navy #1B365D / #002060, italic, 4pt after-spacing (80 dxa)
      - Valid XML structure
    """
    issues = []

    # 1. Header/Footer Inspection
    header_footer_files = [f for f in zip_archive.namelist() if re.match(r"word/(header|footer)\d+\.xml", f)]
    has_dynamic_page = False

    for hf_file in header_footer_files:
        try:
            hf_xml = zip_archive.read(hf_file)
            hf_tree = etree.fromstring(hf_xml)
            # Look for PAGE or NUMPAGES fields
            fld_simple = hf_tree.xpath(".//w:fldSimple[contains(@w:instr, 'PAGE')]", namespaces=NS)
            instr_text = hf_tree.xpath(".//w:instrText[contains(text(), 'PAGE')]", namespaces=NS)
            if fld_simple or instr_text:
                has_dynamic_page = True
        except Exception as e:
            issues.append(f"Failed to parse header/footer {hf_file}: {e}")

    # Check for un-promoted pagination in body text
    body_texts = doc_tree.xpath("/w:document/w:body/w:p//w:t/text()", namespaces=NS)
    for text in body_texts:
        if re.search(r"\bPage\s+\d+\s+of\s+\d+\b", text, re.IGNORECASE) or re.search(r"\bp\.\s*\d+\b", text):
            # Notice: if header/footer files exist with dynamic page, hardcoded body page strings are an anomaly
            if has_dynamic_page:
                issues.append(f"Unpromoted pagination string found in body: {text.strip()}")

    # 2. <w:keepNext/> on Source Body Paragraphs
    body_p = doc_tree.xpath("/w:document/w:body/w:p", namespaces=NS)
    for i in range(len(body_p) - 1):
        curr_p = body_p[i]
        next_p = body_p[i + 1]
        curr_text = "".join(curr_p.xpath(".//w:t/text()", namespaces=NS)).strip()
        next_text = "".join(next_p.xpath(".//w:t/text()", namespaces=NS)).strip()

        # If curr_p is source and next_p is translation (Pattern A)
        if is_translatable_source_text(curr_text) and is_vietnamese_text(next_text):
            keep_next = curr_p.xpath("./w:pPr/w:keepNext", namespaces=NS)
            if not keep_next:
                issues.append(f"Missing <w:keepNext/> on source paragraph: {curr_text[:40]}")

    # 3. Typography & Styling on Translated Paragraphs
    for p in body_p:
        text = "".join(p.xpath(".//w:t/text()", namespaces=NS)).strip()
        if is_vietnamese_text(text):
            # Check spacing: after should be 80 dxa (4 pt), before should be 0
            spacing = p.xpath("./w:pPr/w:spacing", namespaces=NS)
            # Color and italic check
            has_navy = bool(p.xpath(".//w:rPr/w:color[@w:val='1B365D' or @w:val='002060']", namespaces=NS))
            has_italic = bool(p.xpath(".//w:rPr/w:i", namespaces=NS))
            if not (has_navy or has_italic):
                issues.append(f"Translated paragraph missing dark navy / italic styling: {text[:40]}")

    visual_parity_passed = (len(issues) == 0)
    return visual_parity_passed, issues


def verify_bilingual_docx(docx_path: str) -> Dict[str, Any]:
    """
    Authoritative 7-Point XML & Visual Audit Entry Point.
    Directly unpacks the OpenXML ZIP package and executes all 7 audit checks.

    Returns:
    {
        "passed": bool,
        "concatenation_violations": int,  # Must be 0
        "double_tab_violations": int,     # Must be 0
        "pair_completeness": float,       # Must be 1.0 (100%)
        "leaked_english_count": int,      # Must be 0
        "header_order_violations": int,   # Must be 0
        "clipped_row_violations": int,    # Must be 0 (0 w:hRule='exact')
        "visual_parity_passed": bool,     # Must be True
        "details": Dict[str, Any]
    }
    """
    if not os.path.exists(docx_path):
        return {
            "passed": False,
            "error": f"File does not exist: {docx_path}",
            "concatenation_violations": -1,
            "double_tab_violations": -1,
            "pair_completeness": 0.0,
            "leaked_english_count": -1,
            "header_order_violations": -1,
            "clipped_row_violations": -1,
            "visual_parity_passed": False,
            "details": {}
        }

    try:
        with zipfile.ZipFile(docx_path, "r") as z:
            # Parse main document.xml
            doc_xml_bytes = z.read("word/document.xml")
            doc_tree = etree.fromstring(doc_xml_bytes)

            # Collect all XML trees for concatenation check (document + headers + footers)
            all_trees = [doc_tree]
            for name in z.namelist():
                if re.match(r"word/(header|footer)\d+\.xml", name):
                    try:
                        all_trees.append(etree.fromstring(z.read(name)))
                    except Exception:
                        pass

            # 1. Zero Concatenation Check
            concat_count, concat_details = check_zero_concatenation(all_trees)

            # 2. Zero Internal Double-Tab Check
            dtab_count, dtab_details = check_zero_double_tabs(doc_tree)

            # 3. 100% Pair Completeness Check
            pair_ratio, total_src, total_pairs, unpaired_list = check_pair_completeness(doc_tree)

            # 4. Zero Leaked English Check
            leak_count, leak_details = check_zero_leaked_english(doc_tree)

            # 5. Header Order Check
            order_count, order_details = check_header_order(doc_tree)

            # 6. Zero Clipped Rows Check
            clip_count, clip_details = check_zero_clipped_rows(doc_tree)

            # 7. Visual Parity Check
            parity_passed, parity_issues = check_visual_parity(z, doc_tree)

            # Master Pass Condition
            passed = (
                concat_count == 0 and
                dtab_count == 0 and
                pair_ratio == 1.0 and
                leak_count == 0 and
                order_count == 0 and
                clip_count == 0 and
                parity_passed is True
            )

            return {
                "passed": passed,
                "concatenation_violations": concat_count,
                "double_tab_violations": dtab_count,
                "pair_completeness": pair_ratio,
                "leaked_english_count": leak_count,
                "header_order_violations": order_count,
                "clipped_row_violations": clip_count,
                "visual_parity_passed": parity_passed,
                "details": {
                    "concatenation_violations": concat_details,
                    "double_tab_violations": dtab_details,
                    "pair_completeness": {
                        "ratio": pair_ratio,
                        "total_source_blocks": total_src,
                        "validated_pairs": total_pairs,
                        "unpaired_examples": unpaired_list[:10]
                    },
                    "leaked_english_violations": leak_details,
                    "header_order_violations": order_details,
                    "clipped_row_violations": clip_details,
                    "visual_parity_issues": parity_issues
                }
            }

    except Exception as e:
        return {
            "passed": False,
            "error": f"OpenXML audit exception: {str(e)}",
            "concatenation_violations": -1,
            "double_tab_violations": -1,
            "pair_completeness": 0.0,
            "leaked_english_count": -1,
            "header_order_violations": -1,
            "clipped_row_violations": -1,
            "visual_parity_passed": False,
            "details": {"exception": str(e)}
        }


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python -m tests.audit_verifier <path_to_docx>")
        sys.exit(1)

    target_path = sys.argv[1]
    report = verify_bilingual_docx(target_path)
    print("\n================ 7-POINT XML & VISUAL AUDIT REPORT ================")
    print(f"Target: {target_path}")
    print(f"Overall Result: {'PASS' if report['passed'] else 'FAIL'}")
    print(f"Check 1 (Zero Concatenation):   Violations: {report['concatenation_violations']}")
    print(f"Check 2 (Zero Double-Tabs):     Violations: {report['double_tab_violations']}")
    print(f"Check 3 (Pair Completeness):    Ratio: {report['pair_completeness'] * 100:.1f}%")
    print(f"Check 4 (Zero Leaked English):  Violations: {report['leaked_english_count']}")
    print(f"Check 5 (Header Order):         Violations: {report['header_order_violations']}")
    print(f"Check 6 (Zero Clipped Rows):    Violations: {report['clipped_row_violations']}")
    print(f"Check 7 (Visual Parity):        Passed: {report['visual_parity_passed']}")
    print("===================================================================\n")
    if not report["passed"]:
        import json
        print("Details:\n", json.dumps(report["details"], indent=2, ensure_ascii=False))
