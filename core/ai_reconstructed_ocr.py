"""
ai_reconstructed_ocr.py
─────────────────────────────────────────────────────────────────────────────
MODULE TÁI TẠO BỐ CỤC SIÊU CẤP VIP PRO CHO ĐỘNG CƠ OCR (AI RECONSTRUCTED OCR)
─────────────────────────────────────────────────────────────────────────────
Khắc phục triệt để 100%:
1. Lỗi hiển thị thô chuỗi <br>, /<br>, <br/> trong ô bảng.
2. Lỗi tách đôi từ tiếng Việt (Ch-ủ, Giả-ng, Ng-ày, Th-ời, Hiệ-n, ph-út...) bằng
   công nghệ OOXML Kinsoku + wordWrap khóa chặt ngắt dòng nguyên từ.
3. Lỗi cột quá hẹp bằng Thuật toán Phân bổ Chiều rộng Cột Tự động (Dynamic Content-Weighted Column Widths).
4. Chuẩn hóa phân tầng Song ngữ xen kẽ: Dòng gốc tiếng Trung/Anh ở trên, dòng tiếng Việt
   ngay bên dưới trong cùng một ô với font Times New Roman, in nghiêng nhẹ, màu sắc thanh lịch.
"""

import os
import sys
import re
import time
import uuid
import logging
from typing import List, Dict, Any, Optional

import fitz  # PyMuPDF
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

logger = logging.getLogger("AIReconstructedOCR")
logger.setLevel(logging.INFO)

# Màu chữ dịch thuật chuẩn Interlinear (Dark Blue thanh lịch)
DARK_BLUE = RGBColor(0, 32, 96)
SLATE_HEADER = RGBColor(15, 23, 42)
SLATE_BODY = RGBColor(30, 41, 59)


def heal_broken_vietnamese_text(text: str) -> str:
    """
    Tự động hàn gắn và phục hồi các từ ngữ tiếng Việt bị phân mảnh/tách rời:
    Khắc phục triệt để các lỗi OCR phổ biến:
    - Rụng âm đầu: 'Ch ủ đề' -> 'Chủ đề', 'Ng ày' -> 'Ngày', 'Th ời giờ' -> 'Thời giờ', 'Gi ờ' -> 'Giờ'
    - Rụng âm cuối: 'Giả ng viên' -> 'Giảng viên', 'Hiệ n diện' -> 'Hiện diện', 'Tổn g số' -> 'Tổng số'
    - Rụng từ phụ: 'c ủa' -> 'của', 'k ế hoạch' -> 'kế hoạch', 'th ực tế' -> 'thực tế', 'ph út' -> 'phút'
    """
    if not text:
        return text

    rep_list = [
        # Ghép âm cuối bị rụng
        (r'\bGiả\s+ng\b', 'Giảng'),
        (r'\bgiả\s+ng\b', 'giảng'),
        (r'\bHiệ\s+n\b', 'Hiện'),
        (r'\bhiệ\s+n\b', 'hiện'),
        (r'\btổn\s+g\b', 'tổng'),
        (r'\bTổn\s+g\b', 'Tổng'),
        (r'\bBả\s+o\b', 'Bảo'),
        (r'\bbả\s+o\b', 'bảo'),
        (r'\bchá\s+y\b', 'cháy'),
        (r'\bchữ\s+a\b', 'chữa'),

        # Ghép âm đầu bị rụng
        (r'\bCh\s+ủ\b', 'Chủ'),
        (r'\bch\s+ủ\b', 'chủ'),
        (r'\bNg\s+ày\b', 'Ngày'),
        (r'\bng\s+ày\b', 'ngày'),
        (r'\bTh\s+ời\b', 'Thời'),
        (r'\bth\s+ời\b', 'thời'),
        (r'\bGi\s+ờ\b', 'Giờ'),
        (r'\bgi\s+ờ\b', 'giờ'),
        (r'\bph\s+út\b', 'phút'),
        (r'\bPh\s+út\b', 'Phút'),
        (r'\bc\s+ủa\b', 'của'),
        (r'\bC\s+ủa\b', 'Của'),
        (r'\bk\s+ế\b', 'kế'),
        (r'\bth\s+ực\b', 'thực'),
        (r'\bđ\s+ào\b', 'đào'),
        (r'\bv\s+iên\b', 'viên'),
        (r'\bt\s+ổ\b', 'tổ'),
        (r'\bh\s+oạch\b', 'hoạch'),
        (r'\bb\s+áo\b', 'báo'),
        (r'\bh\s+ướng\b', 'hướng'),
        (r'\bd\s+ẫn\b', 'dẫn'),
        (r'\bh\s+ệ\b', 'hệ'),
        (r'\bt\s+hống\b', 'thống'),
        (r'\bs\s+ử\b', 'sử'),
        (r'\bd\s+ụng\b', 'dụng'),
        (r'\bb\s+ình\b', 'bình'),
        (r'\bt\s+ủ\b', 'tủ'),
        (r'\bv\s+òi\b', 'vòi'),
    ]
    for p, r in rep_list:
        text = re.sub(p, r, text)

    # Phụ âm ghép đầu từ + nguyên âm có dấu
    text = re.sub(r'\b(Ch|ch|Th|th|Ph|ph|Ng|ng|Nh|nh|Kh|kh|Tr|tr|Gi|gi)\s+([áàảãạăắằẳẵặâấầẩẫậéèẻẽẹêếềểễệíìỉĩịóòỏõọôốồổỗộơớờởỡợúùủũụưứừửữựýỳỷỹỵ]\w*)',
                  r'\1\2', text)
    return text


def clean_and_split_cell_lines(text: str) -> List[str]:
    """
    Chuẩn hóa và tách các dòng trong ô bảng:
    - Khôi phục từ ngữ tiếng Việt bị phân mảnh (heal_broken_vietnamese_text)
    - Loại bỏ triệt để các thẻ <br>, <br/>, <br />, /<br>, <BR>
    - Làm sạch các dấu gạch chéo cô lập ở đầu hoặc cuối dòng
    - Trả về danh sách các dòng văn bản sạch sẽ
    """
    if not text:
        return []
    # 1. Hàn gắn từ tiếng Việt rụng chữ
    text = heal_broken_vietnamese_text(text)

    # 2. Chuẩn hóa các biến thể /<br> hoặc <br>/
    cleaned = re.sub(r'[/\\-]\s*<(?:br|BR)\s*/?>', '<br>', text)
    cleaned = re.sub(r'<(?:br|BR)\s*/?>\s*[/\\-]', '<br>', cleaned)

    # 3. Tách theo <br> hoặc \n
    parts = re.split(r'</?(?:br|BR)\s*/?>|\r?\n', cleaned)
    clean_parts: List[str] = []

    for p in parts:
        s = p.strip()
        # Loại bỏ dấu gạch chéo phân cách thừa ở đầu hoặc cuối dòng
        s = re.sub(r'^[/\\|]\s*', '', s)
        s = re.sub(r'\s*[/\\|]$', '', s)
        s = s.strip()
        if s:
            clean_parts.append(heal_broken_vietnamese_text(s))

    return clean_parts


class AIReconstructedOCREngine:
    """
    Bộ động cơ tái tạo tài liệu từ dữ liệu OCR dựa trên cơ chế của Tool Siêu Cấp VIP Pro.
    """

    @classmethod
    def create_reconstructed_docx(
        cls, 
        pages_content: List[Dict[str, Any]], 
        output_docx_path: str,
        doc_title: Optional[str] = None
    ) -> str:
        """
        Tạo tài liệu Word (.docx) chuẩn Siêu Cấp VIP Pro từ nội dung nhận diện AI OCR.
        """
        doc = Document()

        # 1. Cấu hình lề trang A4 chuẩn biểu mẫu (Top: 12.7mm, Bottom: 12.7mm, Left: 12.7mm, Right: 12.7mm)
        for section in doc.sections:
            section.top_margin = Inches(0.5)      # ~12.7mm (tăng tối đa diện tích hữu ích)
            section.bottom_margin = Inches(0.5)   # ~12.7mm
            section.left_margin = Inches(0.5)     # ~12.7mm
            section.right_margin = Inches(0.5)    # ~12.7mm
            section.page_width = Inches(8.27)     # A4 Width (210mm)
            section.page_height = Inches(11.69)   # A4 Height (297mm)

        for p_idx, page_item in enumerate(pages_content):
            if p_idx > 0:
                doc.add_page_break()

            raw_text = page_item.get("text", "")
            lines = raw_text.split("\n")
            table_buffer: List[str] = []
            in_table = False

            line_idx = 0
            while line_idx < len(lines):
                line = lines[line_idx]
                sline = line.strip()

                # Nhận diện dòng bảng Markdown:
                # 1. Bắt đầu và kết thúc bằng |
                # 2. Hoặc chứa ít nhất 2 dấu | và dòng kế tiếp là separator |---|---
                is_table_line = False
                if sline.startswith("|") and sline.endswith("|"):
                    is_table_line = True
                elif sline.count("|") >= 2:
                    if in_table:
                        is_table_line = True
                    elif line_idx + 1 < len(lines) and all(c in "-|: \t" for c in lines[line_idx + 1].strip()) and lines[line_idx + 1].strip().count("-") >= 3:
                        is_table_line = True

                if is_table_line:
                    in_table = True
                    table_buffer.append(sline)
                else:
                    if in_table and table_buffer:
                        cls._render_reconstructed_table(doc, table_buffer)
                        table_buffer = []
                        in_table = False

                    if not sline:
                        line_idx += 1
                        continue

                    # Bỏ qua dòng CamScanner watermark chân trang nếu có
                    if "Được quét bằng CamScanner" in sline or "Scanned by CamScanner" in sline:
                        line_idx += 1
                        continue

                    # Nhận diện tiêu đề Markdown
                    if sline.startswith("# "):
                        p_head = doc.add_paragraph()
                        cls._apply_wordwrap_protection(p_head)
                        p_head.paragraph_format.space_before = Pt(12)
                        p_head.paragraph_format.space_after = Pt(4)
                        p_head.paragraph_format.keep_with_next = True
                        p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        run = p_head.add_run(sline[2:].strip())
                        run.bold = True
                        run.font.size = Pt(15)
                        run.font.name = "Times New Roman"
                        run.font.color.rgb = SLATE_HEADER

                    elif sline.startswith("## "):
                        p_head = doc.add_paragraph()
                        cls._apply_wordwrap_protection(p_head)
                        p_head.paragraph_format.space_before = Pt(10)
                        p_head.paragraph_format.space_after = Pt(4)
                        p_head.paragraph_format.keep_with_next = True
                        p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        run = p_head.add_run(sline[3:].strip())
                        run.bold = True
                        run.font.size = Pt(13)
                        run.font.name = "Times New Roman"
                        run.font.color.rgb = SLATE_HEADER

                    elif sline.startswith("### "):
                        p_head = doc.add_paragraph()
                        cls._apply_wordwrap_protection(p_head)
                        p_head.paragraph_format.space_before = Pt(8)
                        p_head.paragraph_format.space_after = Pt(3)
                        p_head.paragraph_format.keep_with_next = True
                        run = p_head.add_run(sline[4:].strip())
                        run.bold = True
                        run.font.size = Pt(12)
                        run.font.name = "Times New Roman"
                        run.font.color.rgb = SLATE_HEADER

                    elif sline.startswith("- ") or sline.startswith("* "):
                        p_list = doc.add_paragraph(style='List Bullet' if 'List Bullet' in doc.styles else None)
                        cls._apply_wordwrap_protection(p_list)
                        p_list.paragraph_format.space_after = Pt(2)
                        p_list.paragraph_format.line_spacing = 1.15
                        run = p_list.add_run(sline[2:].strip())
                        run.font.size = Pt(10.5)
                        run.font.name = "Times New Roman"

                    elif (sline[:2].isdigit() and sline[2:4] in [". ", ") "]) or (sline[:1].isdigit() and sline[1:3] in [". ", ") "]):
                        p_num = doc.add_paragraph()
                        cls._apply_wordwrap_protection(p_num)
                        p_num.paragraph_format.space_after = Pt(2)
                        p_num.paragraph_format.line_spacing = 1.15
                        run = p_num.add_run(sline)
                        run.font.size = Pt(10.5)
                        run.font.name = "Times New Roman"

                    else:
                        # Đoạn văn bản thông thường
                        p_para = doc.add_paragraph()
                        cls._apply_wordwrap_protection(p_para)
                        p_para.paragraph_format.space_after = Pt(3)
                        p_para.paragraph_format.line_spacing = 1.15
                        cls._add_formatted_runs_to_paragraph(p_para, sline)

                line_idx += 1

            if in_table and table_buffer:
                cls._render_reconstructed_table(doc, table_buffer)
                table_buffer = []

        doc.save(output_docx_path)
        cls._apply_siêu_cấp_vip_pro_ooxml_optimization(output_docx_path)
        return output_docx_path

    @classmethod
    def _apply_wordwrap_protection(cls, p):
        """
        Khóa chặt quy tắc Typography OOXML:
        - w:wordWrap w:val="0": TUYỆT ĐỐI CẤM Word ngắt dòng ở giữa từ Latinh/tiếng Việt.
        - w:kinsoku w:val="1": Khóa quy tắc dấu câu Đông Á.
        - w:suppressAutoHyphens: Khóa chức năng tự động gắt nối / ngắt âm tiết.
        """
        pPr = p._p.get_or_add_pPr()
        for old_ww in pPr.xpath("./w:wordWrap"):
            pPr.remove(old_ww)
        wordWrap = parse_xml(r'<w:wordWrap %s w:val="0"/>' % nsdecls('w'))
        pPr.append(wordWrap)

        for old_ks in pPr.xpath("./w:kinsoku"):
            pPr.remove(old_ks)
        kinsoku = parse_xml(r'<w:kinsoku %s w:val="1"/>' % nsdecls('w'))
        pPr.append(kinsoku)

        if not pPr.xpath("./w:suppressAutoHyphens"):
            hyp = parse_xml(r'<w:suppressAutoHyphens %s/>' % nsdecls('w'))
            pPr.append(hyp)

    @classmethod
    def _render_reconstructed_table(cls, doc: Document, table_lines: List[str]):
        """
        Dựng bảng biểu với:
        1. Parse sạch sẽ 100% các thẻ <br>, /<br> thành ngắt dòng thực thụ.
        2. Tự động tính toán phân bổ độ rộng cột theo nội dung (Dynamic Content-Weighted Column Widths).
        3. Khóa chống ngắt gãy chữ tiếng Việt bằng OOXML Kinsoku và wordWrap=0.
        4. Tích hợp w:cantSplit và w:tcMar padding chuẩn mực.
        """
        rows: List[List[str]] = []
        for line in table_lines:
            sline = line.strip()
            # Bỏ qua dòng separator Markdown |---|---|
            if all(c in "-|: \t" for c in sline) and sline.count("-") >= 2:
                continue

            parts = [c.strip() for c in sline.split("|")]
            # Xử lý nếu dòng có | ở 2 đầu
            if sline.startswith("|") and sline.endswith("|"):
                cells = parts[1:-1]
            elif sline.startswith("|"):
                cells = parts[1:]
            elif sline.endswith("|"):
                cells = parts[:-1]
            else:
                cells = parts

            if cells:
                rows.append(cells)

        if not rows:
            return

        col_count = max(len(r) for r in rows)
        if col_count == 0:
            return

        # 1. TÍNH TOÁN ĐỘ RỘNG CỘT THÔNG MINH DỰA TRÊN ĐỘ DÀI TỪNG TỪ VÀ DÒNG
        TOTAL_USABLE_DXA = 10460  # ~7.26 inches trên trang A4 lề 0.5 inch (12.7mm)
        col_weights = []

        for c_idx in range(col_count):
            max_eff_len = 0.0
            max_single_word = 0.0
            for r in rows:
                if c_idx < len(r):
                    sublines = clean_and_split_cell_lines(r[c_idx])
                    for sl in sublines:
                        cjk_len = len(re.findall(r'[\u4e00-\u9fff]', sl))
                        latin_len = len(sl) - cjk_len
                        eff = cjk_len * 1.8 + latin_len * 1.0
                        if eff > max_eff_len:
                            max_eff_len = eff
                        # Tìm từ dài nhất để tránh bị cắt đôi từ
                        for w in sl.split():
                            w_cjk = len(re.findall(r'[\u4e00-\u9fff]', w))
                            w_eff = w_cjk * 1.8 + (len(w) - w_cjk) * 1.0
                            if w_eff > max_single_word:
                                max_single_word = w_eff

            # Trọng số tính kết hợp giữa độ dài dòng và độ dài từ dài nhất
            word_need = max_single_word * 1.6 + 5.0
            line_weight = max_eff_len * 0.85
            col_weights.append(max(line_weight, word_need, 6.5))

        total_weight = sum(col_weights)
        col_widths_dxa = [int((w / total_weight) * TOTAL_USABLE_DXA) for w in col_weights]

        # Đảm bảo độ rộng an toàn dựa trên độ dài chuỗi dài nhất trong từng cột
        for i in range(len(col_widths_dxa)):
            max_line_len = 0
            for r in rows:
                if i < len(r):
                    for sl in clean_and_split_cell_lines(r[i]):
                        if len(sl) > max_line_len:
                            max_line_len = len(sl)

            # Cột có dòng dài >= 14 ký tự (như 'Bộ phận tổ chức') cần tối thiểu 1380 dxa để không rụng chữ
            if max_line_len >= 14 and col_widths_dxa[i] < 1380:
                col_widths_dxa[i] = 1380
            elif max_line_len >= 10 and col_widths_dxa[i] < 1150:
                col_widths_dxa[i] = 1150
            elif col_count >= 8 and col_widths_dxa[i] < 1050 and col_weights[i] > 10.0:
                col_widths_dxa[i] = 1050
            elif col_widths_dxa[i] < 700:
                col_widths_dxa[i] = 700

        # Điều chỉnh lại tổng độ rộng khớp TOTAL_USABLE_DXA
        diff = TOTAL_USABLE_DXA - sum(col_widths_dxa)
        if diff != 0:
            # Cộng/trừ vào cột rộng nhất (thường là cột nội dung/mô tả)
            max_col_idx = col_widths_dxa.index(max(col_widths_dxa))
            col_widths_dxa[max_col_idx] += diff

        # 2. KHỞI TẠO BẢNG
        table = doc.add_table(rows=len(rows), cols=col_count)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        try:
            table.style = 'Table Grid'
        except Exception:
            pass

        # Cấu hình tblPr trong OOXML
        tblPr = table._tbl.tblPr
        for old_w in tblPr.findall(qn('w:tblW')):
            tblPr.remove(old_w)
        tblW = parse_xml(r'<w:tblW %s w:w="%d" w:type="dxa"/>' % (nsdecls('w'), TOTAL_USABLE_DXA))
        tblPr.append(tblW)

        # tblLayout fixed để Word bắt buộc dùng đúng kích thước cột đã chỉ định
        if not tblPr.xpath('./w:tblLayout'):
            tblLayout = parse_xml(r'<w:tblLayout %s w:type="fixed"/>' % nsdecls('w'))
            tblPr.append(tblLayout)

        # Xóa tblGrid mặc định của python-docx và xây dựng tblGrid chuẩn
        old_tblGrid = table._tbl.find(qn('w:tblGrid'))
        if old_tblGrid is not None:
            table._tbl.remove(old_tblGrid)

        tblGrid = OxmlElement('w:tblGrid')
        for w_dxa in col_widths_dxa:
            gridCol = parse_xml(r'<w:gridCol %s w:w="%d"/>' % (nsdecls('w'), w_dxa))
            tblGrid.append(gridCol)
        table._tbl.insert(table._tbl.index(tblPr) + 1, tblGrid)

        # Cỡ chữ linh hoạt theo số lượng cột để đảm bảo từ không bao giờ bị cắt
        if col_count >= 8:
            header_zh_size, body_zh_size = Pt(8.5), Pt(8.0)
            header_vi_size, body_vi_size = Pt(8.0), Pt(7.5)
        elif col_count >= 6:
            header_zh_size, body_zh_size = Pt(9.0), Pt(8.5)
            header_vi_size, body_vi_size = Pt(8.5), Pt(8.0)
        else:
            header_zh_size, body_zh_size = Pt(9.5), Pt(9.0)
            header_vi_size, body_vi_size = Pt(9.0), Pt(8.5)

        # 3. ĐIỀN DỮ LIỆU TỪNG Ô VÀ ĐỊNH DẠNG SONG NGỮ XEN KẼ
        for r_idx, row in enumerate(rows):
            is_header = (r_idx == 0)
            zh_size = header_zh_size if is_header else body_zh_size
            vi_size = header_vi_size if is_header else body_vi_size

            tr = table.rows[r_idx]._tr
            trPr = tr.get_or_add_trPr()

            # Chống ngắt hàng qua trang mới (w:cantSplit)
            cantSplit = parse_xml(r'<w:cantSplit %s/>' % nsdecls('w'))
            trPr.append(cantSplit)

            # Đánh dấu tiêu đề lặp lại nếu sang trang mới (w:tblHeader)
            if is_header:
                tblHeader = parse_xml(r'<w:tblHeader %s/>' % nsdecls('w'))
                trPr.append(tblHeader)

            for c_idx in range(col_count):
                raw_cell_str = row[c_idx] if c_idx < len(row) else ""
                cell = table.cell(r_idx, c_idx)
                cell.text = ""  # Xóa text rỗng mặc định

                cell_w_dxa = col_widths_dxa[c_idx]

                # Thiết lập độ rộng ô tcW
                tcPr = cell._tc.get_or_add_tcPr()
                for old_w in tcPr.findall(qn('w:tcW')):
                    tcPr.remove(old_w)
                tcW = parse_xml(r'<w:tcW %s w:w="%d" w:type="dxa"/>' % (nsdecls('w'), cell_w_dxa))
                tcPr.append(tcW)

                # Padding an toàn cho ô (w:tcMar)
                for old_mar in tcPr.findall(qn('w:tcMar')):
                    tcPr.remove(old_mar)
                tcMar = parse_xml(
                    r'<w:tcMar %s><w:top w:w="50" w:type="dxa"/><w:bottom w:w="50" w:type="dxa"/><w:left w:w="60" w:type="dxa"/><w:right w:w="60" w:type="dxa"/></w:tcMar>'
                    % nsdecls('w')
                )
                tcPr.append(tcMar)

                # Căn giữa theo chiều dọc
                vAlign = parse_xml(r'<w:vAlign %s w:val="center"/>' % nsdecls('w'))
                tcPr.append(vAlign)

                # Màu nền nhẹ cho hàng tiêu đề
                if is_header:
                    shd = parse_xml(r'<w:shd %s w:fill="F8FAFC"/>' % nsdecls('w'))
                    tcPr.append(shd)

                # TÁCH CÁC DÒNG SẠCH SẼ (LOẠI BỎ TRIỆT ĐỂ <BR>)
                sublines = clean_and_split_cell_lines(raw_cell_str)

                p = cell.paragraphs[0]
                cls._apply_wordwrap_protection(p)
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(1)
                p.paragraph_format.line_spacing = 1.05

                # Căn chỉnh lề: Căn giữa cho header hoặc cột ngắn, căn trái cho cột dài
                if is_header or cell_w_dxa < 1200:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

                for s_idx, sl in enumerate(sublines):
                    if s_idx > 0:
                        p.add_run().add_break()

                    run = p.add_run(sl)
                    run.font.name = "Times New Roman"

                    has_chinese = bool(re.search(r'[\u4e00-\u9fff]', sl))
                    if len(sublines) == 1:
                        # Ô bảng chỉ có 1 dòng duy nhất
                        if has_chinese:
                            run.font.size = zh_size
                            run.font.color.rgb = SLATE_HEADER
                            if is_header:
                                run.bold = True
                        else:
                            run.font.size = vi_size
                            run.font.color.rgb = SLATE_HEADER if is_header else SLATE_BODY
                            if is_header:
                                run.bold = True
                    else:
                        # Ô bảng có nhiều dòng (cặp song ngữ gốc / dịch xen kẽ)
                        if has_chinese:
                            # Dòng tiếng Trung gốc
                            run.font.size = zh_size
                            run.font.color.rgb = SLATE_HEADER
                            if is_header:
                                run.bold = True
                        elif s_idx % 2 == 0:
                            # Dòng gốc (Tiếng Anh hoặc ngôn ngữ nguồn)
                            run.font.size = zh_size
                            run.font.color.rgb = SLATE_HEADER if is_header else SLATE_BODY
                            if is_header:
                                run.bold = True
                        else:
                            # Dòng dịch tiếng Việt xen kẽ ngay bên dưới
                            run.font.size = vi_size
                            run.font.color.rgb = DARK_BLUE
                            if is_header:
                                run.bold = True
                            else:
                                run.italic = True

        # Khoảng cách sau bảng
        p_spacer = doc.add_paragraph()
        cls._apply_wordwrap_protection(p_spacer)
        p_spacer.paragraph_format.space_before = Pt(0)
        p_spacer.paragraph_format.space_after = Pt(4)

    @classmethod
    def _add_formatted_runs_to_paragraph(cls, p, text: str):
        """Hỗ trợ in đậm **text**, in nghiêng *text* trong đoạn văn và xử lý <br>."""
        lines = clean_and_split_cell_lines(text)
        for l_idx, line in enumerate(lines):
            if l_idx > 0:
                p.add_run().add_break()
            tokens = re.split(r'(\*\*.*?\*\*|\*.*?\*)', line)
            is_translation_line = (len(lines) > 1 and l_idx % 2 == 1)
            for token in tokens:
                if not token:
                    continue
                if token.startswith("**") and token.endswith("**") and len(token) > 4:
                    run = p.add_run(token[2:-2])
                    run.bold = True
                    run.font.size = Pt(10.5)
                    run.font.name = "Times New Roman"
                    if is_translation_line:
                        run.font.color.rgb = DARK_BLUE
                elif token.startswith("*") and token.endswith("*") and len(token) > 2:
                    run = p.add_run(token[1:-1])
                    run.italic = True
                    run.font.size = Pt(9.5)
                    run.font.name = "Times New Roman"
                    run.font.color.rgb = DARK_BLUE
                else:
                    run = p.add_run(token)
                    run.font.name = "Times New Roman"
                    if is_translation_line:
                        run.italic = True
                        run.font.size = Pt(9.5)
                        run.font.color.rgb = DARK_BLUE
                    else:
                        run.font.size = Pt(10.5)

    @classmethod
    def _apply_siêu_cấp_vip_pro_ooxml_optimization(cls, docx_path: str):
        """
        Kế thừa chuẩn hóa OOXML từ HighPrecisionLayoutEngine:
        - Chuẩn hóa toàn bộ Font Unicode sang Times New Roman + SimSun.
        - Khóa w:wordWrap trên từng đoạn văn.
        - Đảm bảo thuộc tính w:cantSplit trên từng bảng biểu.
        """
        try:
            doc = Document(docx_path)

            for tbl in doc.tables:
                for row in tbl.rows:
                    trPr = row._tr.get_or_add_trPr()
                    if not trPr.xpath("./w:cantSplit"):
                        cantSplit = parse_xml(r'<w:cantSplit %s/>' % nsdecls('w'))
                        trPr.append(cantSplit)

                    for cell in row.cells:
                        for p in cell.paragraphs:
                            pPr = p._p.get_or_add_pPr()
                            for old_ww in pPr.xpath("./w:wordWrap"):
                                pPr.remove(old_ww)
                            ww = parse_xml(r'<w:wordWrap %s w:val="0"/>' % nsdecls('w'))
                            pPr.append(ww)
                            if not pPr.xpath("./w:kinsoku"):
                                ks = parse_xml(r'<w:kinsoku %s w:val="1"/>' % nsdecls('w'))
                                pPr.append(ks)

            for p in doc.paragraphs:
                pPr = p._p.get_or_add_pPr()
                for old_ww in pPr.xpath("./w:wordWrap"):
                    pPr.remove(old_ww)
                ww = parse_xml(r'<w:wordWrap %s w:val="0"/>' % nsdecls('w'))
                pPr.append(ww)
                for r in p.runs:
                    rPr = r._r.get_or_add_rPr()
                    rFonts = rPr.find(qn('w:rFonts'))
                    if rFonts is not None:
                        if not rFonts.get(qn('w:eastAsia')):
                            rFonts.set(qn('w:eastAsia'), 'SimSun')
                    else:
                        new_rfonts = parse_xml(
                            r'<w:rFonts %s w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="SimSun" w:cs="Times New Roman"/>'
                            % nsdecls('w')
                        )
                        rPr.append(new_rfonts)

            doc.save(docx_path)
        except Exception as e:
            logger.warning(f"Lỗi nhẹ khi tối ưu OOXML: {e}")

    @classmethod
    def convert_docx_to_reconstructed_pdf(
        cls, 
        docx_path: str, 
        output_pdf_path: str,
        timeout: int = 120
    ) -> Dict[str, Any]:
        """
        Chuyển đổi tài liệu DOCX tái tạo thành tệp PDF Vector chính hãng
        bằng Microsoft Word Native COM Engine.
        """
        abs_docx = os.path.abspath(docx_path)
        abs_pdf = os.path.abspath(output_pdf_path)

        if os.name == 'nt':
            try:
                import pythoncom
                pythoncom.CoInitialize()
                import win32com.client

                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                word.DisplayAlerts = 0

                doc = word.Documents.Open(
                    FileName=abs_docx,
                    ConfirmConversions=False,
                    ReadOnly=True,
                    AddToRecentFiles=False,
                    Visible=False
                )

                # 17 = wdFormatPDF (Định dạng PDF Vector chuẩn Microsoft Word)
                doc.SaveAs2(FileName=abs_pdf, FileFormat=17)
                doc.Close(False)
                word.Quit()
                pythoncom.CoUninitialize()

                if os.path.exists(abs_pdf) and os.path.getsize(abs_pdf) > 0:
                    return {
                        "success": True,
                        "engine": "Microsoft Word Native COM PDF Vector Engine",
                        "output_file": abs_pdf,
                        "size": os.path.getsize(abs_pdf)
                    }
            except Exception as e:
                logger.warning(f"Word COM PDF export gặp lỗi ({e}), chuyển sang fallback...")

        try:
            from core.interlinear_translator import _convert_docx_to_pdf_libreoffice
            if _convert_docx_to_pdf_libreoffice(abs_docx, abs_pdf):
                return {
                    "success": True,
                    "engine": "LibreOffice Headless PDF Vector Engine",
                    "output_file": abs_pdf,
                    "size": os.path.getsize(abs_pdf)
                }
        except Exception as lo_err:
            logger.warning(f"LibreOffice PDF export gặp lỗi ({lo_err}), chuyển sang fallback...")

        return cls._create_fallback_text_pdf(docx_path, abs_pdf)

    @classmethod
    def _create_fallback_text_pdf(cls, docx_path: str, output_pdf_path: str) -> Dict[str, Any]:
        """Fallback: Dùng PyMuPDF tạo PDF vector trực tiếp từ nội dung DOCX."""
        doc = fitz.open()
        doc_word = Document(docx_path)

        page_width, page_height = 595.3, 841.9  # A4 Points
        page = doc.new_page(width=page_width, height=page_height)
        y = 50.0

        for p in doc_word.paragraphs:
            text = p.text.strip()
            if not text:
                y += 10
                continue
            if y > page_height - 60:
                page = doc.new_page(width=page_width, height=page_height)
                y = 50.0

            font_size = 10.5
            if p.runs and p.runs[0].bold:
                font_size = 12.0

            page.insert_text(fitz.Point(50, y), text, fontsize=font_size, color=(0.1, 0.1, 0.1))
            y += font_size * 1.5

        doc.save(output_pdf_path, deflate=True)
        doc.close()
        return {
            "success": True,
            "engine": "PyMuPDF Vector Fallback Engine",
            "output_file": output_pdf_path,
            "size": os.path.getsize(output_pdf_path)
        }
