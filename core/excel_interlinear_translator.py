"""
core/excel_interlinear_translator.py
───────────────────────────────────
Động cơ Dịch thuật Song ngữ Kỹ thuật Xen kẽ cho Bảng tính Excel (.xlsx, .xls)
Tuân thủ tuyệt đối quy chuẩn DOC-TRAN:
- Dòng trên: Văn bản gốc (giữ nguyên font, bold, color).
- Dòng dưới: Bản dịch tiếng Việt (\\n), font Times New Roman, in nghiêng, màu xanh đậm #002060.
- Bảo toàn 100% công thức toán học (=SUM, =AVERAGE, =IF, etc.), merged cells, độ rộng cột, viền và màu nền.
- Bảo vệ mã số CAS (\\b\\d{2,7}-\\d{2}-\\d\\b) và các định danh kỹ thuật hóa học.
"""

from __future__ import annotations

import os
import sys
import re
import io
import shutil
import subprocess
import uuid
import datetime
import logging
from typing import Dict, List, Tuple, Any, Optional, Union
from pathlib import Path

import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.cell.cell import MergedCell, Cell
from openpyxl.utils import get_column_letter

logger = logging.getLogger("uvicorn.error") if logging.getLogger("uvicorn.error").hasHandlers() else logging.getLogger(__name__)

from core.interlinear_translator import (
    protect_technical_tokens,
    restore_technical_tokens,
    should_translate_item,
    InterlinearTranslator,
    LANG_NAME_MAP,
    ProtectedText,
    CJK_FONT_MAP,
    is_different_language_pair,
    detect_script_type,
)
from core.deepseek_ocr import DeepSeekOCR
from core.context_profiler import AIContextProfiler, ContextProfile
from core.quality_auditor import AIQualityAuditor, AuditReport


class ExcelInterlinearTranslator:
    """
    Động cơ Dịch Song Ngữ Kỹ Thuật DOC-TRAN cho Bảng Tính Excel.
    Sử dụng openpyxl với CellRichText để tạo dòng song ngữ xen kẽ trực tiếp trong từng ô.
    """

    DARK_BLUE_HEX = "002060"
    DARK_BLUE_ARGB = "FF002060"
    STANDARD_FONT = "Times New Roman"

    _CACHE: Dict[Tuple[str, str, str], str] = {}

    @classmethod
    def _find_soffice_binary(cls) -> Optional[str]:
        """Tìm đường dẫn tệp thực thi soffice (LibreOffice) trên hệ thống."""
        env_soffice = os.environ.get("SOFFICE_PATH")
        if env_soffice and os.path.exists(env_soffice):
            return env_soffice

        found_in_path = shutil.which("soffice") or shutil.which("libreoffice")
        if found_in_path:
            return found_in_path

        if os.name == 'nt':
            candidates = [
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                r"D:\Program Files\LibreOffice\program\soffice.exe",
            ]
            for c in candidates:
                if os.path.exists(c):
                    return c
        else:
            candidates = [
                "/usr/bin/soffice",
                "/usr/bin/libreoffice",
                "/usr/local/bin/soffice",
                "/usr/local/bin/libreoffice",
                "/usr/lib/libreoffice/program/soffice.bin",
                "/opt/libreoffice/program/soffice",
            ]
            for c in candidates:
                if os.path.exists(c):
                    return c
        return None

    @classmethod
    def convert_xls_to_xlsx(cls, xls_path: str, output_xlsx_path: Optional[str] = None) -> str:
        """
        Chuyển đổi tệp .xls (Excel 97-2003 nhị phân) sang định dạng chuẩn .xlsx:
        - Trên Windows: Ưu tiên LibreOffice CLI, fallback sang Excel COM.
        - Trên Linux (Render / Docker): Sử dụng headless LibreOffice CLI.
        """
        abs_xls = os.path.abspath(xls_path)
        if not os.path.exists(abs_xls):
            raise FileNotFoundError(f"Tệp nguồn không tồn tại: {xls_path}")

        if not output_xlsx_path:
            out_dir = os.path.dirname(abs_xls)
            stem = Path(abs_xls).stem
            output_xlsx_path = os.path.join(out_dir, f"{stem}.xlsx")

        abs_xlsx = os.path.abspath(output_xlsx_path)
        out_dir = os.path.dirname(abs_xlsx) or "."
        os.makedirs(out_dir, exist_ok=True)

        # 1. Thử chuyển đổi bằng LibreOffice CLI (Headless)
        soffice = cls._find_soffice_binary()
        if soffice:
            creationflags = 0
            if os.name == 'nt' and hasattr(subprocess, 'CREATE_NO_WINDOW'):
                creationflags = subprocess.CREATE_NO_WINDOW

            cmd = [
                soffice,
                "--headless",
                "--convert-to",
                "xlsx",
                "--outdir",
                out_dir,
                abs_xls
            ]
            try:
                proc = subprocess.run(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    timeout=120,
                    creationflags=creationflags
                )
                stem = Path(abs_xls).stem
                default_converted = os.path.join(out_dir, f"{stem}.xlsx")
                if os.path.exists(default_converted):
                    if os.path.abspath(default_converted) != abs_xlsx:
                        if os.path.exists(abs_xlsx):
                            try:
                                os.remove(abs_xlsx)
                            except Exception:
                                pass
                        shutil.move(default_converted, abs_xlsx)
                    if os.path.exists(abs_xlsx) and os.path.getsize(abs_xlsx) > 0:
                        logger.info("Chuyển đổi .xls sang .xlsx thành công qua LibreOffice: %s", abs_xlsx)
                        return abs_xlsx
            except Exception as e:
                logger.warning("LibreOffice .xls sang .xlsx conversion thất bại: %s", e)

        # 2. Trên Windows: Thử chuyển đổi bằng Excel COM
        if os.name == 'nt':
            try:
                import pythoncom
                pythoncom.CoInitialize()
                import win32com.client
                excel = None
                wb = None
                try:
                    excel = win32com.client.DispatchEx("Excel.Application")
                    excel.Visible = False
                    excel.DisplayAlerts = False
                    excel.ScreenUpdating = False
                    wb = excel.Workbooks.Open(abs_xls, ReadOnly=True)
                    # FileFormat 51 = xlOpenXMLWorkbook (.xlsx)
                    wb.SaveAs(abs_xlsx, FileFormat=51)
                    if os.path.exists(abs_xlsx) and os.path.getsize(abs_xlsx) > 0:
                        logger.info("Chuyển đổi .xls sang .xlsx thành công qua Excel COM: %s", abs_xlsx)
                        return abs_xlsx
                finally:
                    if wb:
                        try:
                            wb.Close(SaveChanges=False)
                        except Exception:
                            pass
                    if excel:
                        try:
                            excel.Quit()
                        except Exception:
                            pass
                    try:
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass
            except Exception as e:
                logger.warning("Excel COM .xls sang .xlsx conversion thất bại: %s", e)

        raise RuntimeError(
            f"Không thể chuyển đổi tệp .xls sang .xlsx: {xls_path}. "
            "Vui lòng cài đặt LibreOffice hoặc Microsoft Office trên máy chủ."
        )

    @classmethod
    def convert_xlsx_to_pdf(cls, xlsx_path: str, output_pdf_path: str) -> str:
        """
        Chuyển đổi bảng tính .xlsx đã dịch sang tài liệu PDF chuẩn in ấn:
        - Sử dụng LibreOffice headless CLI hoặc Excel COM.
        """
        abs_xlsx = os.path.abspath(xlsx_path)
        abs_pdf = os.path.abspath(output_pdf_path)
        out_dir = os.path.dirname(abs_pdf) or "."
        os.makedirs(out_dir, exist_ok=True)

        # 1. Thử LibreOffice CLI
        soffice = cls._find_soffice_binary()
        if soffice:
            creationflags = 0
            if os.name == 'nt' and hasattr(subprocess, 'CREATE_NO_WINDOW'):
                creationflags = subprocess.CREATE_NO_WINDOW

            cmd = [
                soffice,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                out_dir,
                abs_xlsx
            ]
            try:
                proc = subprocess.run(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    timeout=120,
                    creationflags=creationflags
                )
                stem = Path(abs_xlsx).stem
                default_pdf = os.path.join(out_dir, f"{stem}.pdf")
                if os.path.exists(default_pdf):
                    if os.path.abspath(default_pdf) != abs_pdf:
                        if os.path.exists(abs_pdf):
                            try:
                                os.remove(abs_pdf)
                            except Exception:
                                pass
                        shutil.move(default_pdf, abs_pdf)
                    if os.path.exists(abs_pdf) and os.path.getsize(abs_pdf) > 0:
                        return abs_pdf
            except Exception as e:
                logger.warning("LibreOffice .xlsx to pdf failed: %s", e)

        # 2. Thử Excel COM trên Windows
        if os.name == 'nt':
            try:
                import pythoncom
                pythoncom.CoInitialize()
                import win32com.client
                excel = None
                wb = None
                try:
                    excel = win32com.client.DispatchEx("Excel.Application")
                    excel.Visible = False
                    excel.DisplayAlerts = False
                    excel.ScreenUpdating = False
                    wb = excel.Workbooks.Open(abs_xlsx, ReadOnly=True)
                    # 0 = xlTypePDF
                    wb.ExportAsFixedFormat(0, abs_pdf)
                    if os.path.exists(abs_pdf) and os.path.getsize(abs_pdf) > 0:
                        return abs_pdf
                finally:
                    if wb:
                        try:
                            wb.Close(SaveChanges=False)
                        except Exception:
                            pass
                    if excel:
                        try:
                            excel.Quit()
                        except Exception:
                            pass
                    try:
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass
            except Exception as e:
                logger.warning("Excel COM .xlsx to pdf failed: %s", e)

        raise RuntimeError(f"Không thể xuất PDF từ file Excel: {xlsx_path}. Thiếu LibreOffice hoặc Excel COM.")

    @classmethod
    def _extract_font_props(cls, cell: Cell) -> Tuple[str, float, bool, bool, str]:
        """Trích xuất an toàn các thuộc tính font của ô gốc."""
        f = cell.font
        name = f.name if f and f.name else "Calibri"
        sz = float(f.size) if f and f.size else 11.0
        b = bool(f.bold) if f and f.bold else False
        i = bool(f.italic) if f and f.italic else False
        color = "FF000000"
        if f and f.color and getattr(f.color, 'rgb', None):
            c_rgb = str(f.color.rgb).upper()
            if len(c_rgb) == 8:
                color = c_rgb
            elif len(c_rgb) == 6:
                color = "FF" + c_rgb
        return name, sz, b, i, color

    @classmethod
    def _is_formula_cell(cls, cell: Cell) -> bool:
        """Kiểm tra ô có phải là công thức toán học không (=SUM, =AVERAGE, etc.)."""
        if getattr(cell, 'data_type', None) == 'f':
            return True
        val = getattr(cell, 'value', None)
        if isinstance(val, str) and val.strip().startswith('='):
            return True
        return False

    @classmethod
    def _is_ignorable_value(cls, val: Any) -> bool:
        """Kiểm tra giá trị ô có phải là số thuần, boolean, ngày tháng hoặc rỗng không."""
        if val is None:
            return True
        if isinstance(val, (int, float, bool, datetime.date, datetime.datetime, datetime.time)):
            return True
        return False

    @staticmethod
    def _contains_cjk(text: str) -> bool:
        """Kiểm tra chuỗi có chứa ký tự tiếng Trung/Nhật/Hàn hay không."""
        if not text:
            return False
        return bool(re.search(r"[\u4e00-\u9fff\u3400-\u4dbf\u3040-\u30ff\uac00-\ud7af]", text))

    @classmethod
    def apply_interlinear_rich_text(
        cls,
        cell: Cell,
        orig_text: str,
        trans_text: str,
        output_mode: str = "bilingual",
        font_family: str = "auto",
        source_lang: str = "auto",
        target_lang: str = "vi"
    ) -> None:
        """
        Áp dụng định dạng Song ngữ Xen kẽ CellRichText cho ô:
        - Dòng trên: TextBlock với font gốc hoặc font CJK tối ưu nếu là chữ Hán/Nhật/Hàn.
        - Dòng dưới: TextBlock với font tối ưu (Times New Roman cho Latin/Việt, font CJK nếu tiếng đích là CJK), màu xanh đậm FF002060, in nghiêng.
        - Bắt buộc wrap_text=True và vertical='center'.
        """
        if not trans_text or not trans_text.strip():
            return

        name, sz, b, i, color = cls._extract_font_props(cell)
        cfg = CJK_FONT_MAP.get((font_family or "auto").lower().strip(), CJK_FONT_MAP["auto"])

        clean_orig = str(orig_text or "").strip()
        clean_trans = str(trans_text or "").strip()

        # Xác định font cho dòng gốc
        orig_font_name = name
        if cls._contains_cjk(clean_orig) or (source_lang and source_lang.lower().strip() in ["zh", "chinese", "ja", "ko"]):
            orig_font_name = cfg.get("cjk", "Microsoft YaHei")

        # Xác định font cho dòng dịch
        if cls._contains_cjk(clean_trans) or (target_lang and target_lang.lower().strip() in ["zh", "chinese", "ja", "ko"]):
            trans_font_name = cfg.get("cjk", "Microsoft YaHei")
        else:
            trans_font_name = cfg.get("latin", cls.STANDARD_FONT)

        font_orig = InlineFont(
            rFont=orig_font_name,
            sz=sz,
            b=b,
            i=i,
            color=color
        )
        font_trans = InlineFont(
            rFont=trans_font_name,
            sz=sz or 11.0,
            color=cls.DARK_BLUE_ARGB,
            b=False,
            i=True
        )

        if output_mode.lower() == "bilingual":
            rt = CellRichText([
                TextBlock(font_orig, clean_orig),
                TextBlock(font_trans, f"\n{clean_trans}")
            ])
            cell.value = rt
        else:
            # Chế độ bản dịch thuần (target only)
            rt = CellRichText([
                TextBlock(font_trans, clean_trans)
            ])
            cell.value = rt

        # Cập nhật căn lề
        curr_align = cell.alignment
        horiz = curr_align.horizontal if (curr_align and curr_align.horizontal) else "left"
        vert = curr_align.vertical if (curr_align and curr_align.vertical) else "center"
        cell.alignment = Alignment(wrap_text=True, vertical=vert, horizontal=horiz)

    @classmethod
    def apply_bilingual_replace_rich_text(
        cls,
        cell: Cell,
        pairs: List[Tuple[str, str]],
        trans_map: Dict[str, str],
        output_mode: str = "bilingual",
        font_family: str = "auto",
        source_lang: str = "auto",
        target_lang: str = "vi"
    ) -> None:
        """
        Ghi đè ô song ngữ có sẵn bằng cách thay thế trực tiếp vào vị trí ngôn ngữ thứ 2:
        - Giữ nguyên dòng 1 (hoặc các dòng chẵn) với font gốc hoặc font CJK tối ưu.
        - Thay thế dòng 2 (hoặc các dòng lẻ) bằng bản dịch mới với font Times New Roman / CJK, màu xanh #002060, in nghiêng.
        - Bảo toàn chính xác số lượng dòng gốc (ví dụ: 2 dòng ra đúng 2 dòng).
        """
        name, sz, b, i, color = cls._extract_font_props(cell)
        cfg = CJK_FONT_MAP.get((font_family or "auto").lower().strip(), CJK_FONT_MAP["auto"])

        rich_blocks = []
        for p_idx, (src_text, old_sec_text) in enumerate(pairs):
            clean_src = str(src_text or "").strip()
            clean_trans = str(trans_map.get(clean_src, clean_src) or "").strip()

            # Font cho dòng gốc
            orig_font_name = name
            if cls._contains_cjk(clean_src) or (source_lang and source_lang.lower().strip() in ["zh", "chinese", "ja", "ko"]):
                orig_font_name = cfg.get("cjk", "Microsoft YaHei")

            # Font cho dòng dịch
            if cls._contains_cjk(clean_trans) or (target_lang and target_lang.lower().strip() in ["zh", "chinese", "ja", "ko"]):
                trans_font_name = cfg.get("cjk", "Microsoft YaHei")
            else:
                trans_font_name = cfg.get("latin", cls.STANDARD_FONT)

            font_orig = InlineFont(
                rFont=orig_font_name,
                sz=sz,
                b=b,
                i=i,
                color=color
            )
            font_trans = InlineFont(
                rFont=trans_font_name,
                sz=sz or 11.0,
                color=cls.DARK_BLUE_ARGB,
                b=False,
                i=True
            )

            if output_mode.lower() == "bilingual":
                prefix = "" if p_idx == 0 else "\n"
                rich_blocks.append(TextBlock(font_orig, f"{prefix}{clean_src}"))
                rich_blocks.append(TextBlock(font_trans, f"\n{clean_trans}"))
            else:
                prefix = "" if p_idx == 0 else "\n"
                rich_blocks.append(TextBlock(font_trans, f"{prefix}{clean_trans}"))

        cell.value = CellRichText(rich_blocks)

        curr_align = cell.alignment
        horiz = curr_align.horizontal if (curr_align and curr_align.horizontal) else "left"
        vert = curr_align.vertical if (curr_align and curr_align.vertical) else "center"
        cell.alignment = Alignment(wrap_text=True, vertical=vert, horizontal=horiz)

    @classmethod
    def translate_excel(
        cls,
        input_path: str,
        output_path: str,
        source_lang: str = "en",
        target_lang: str = "vi",
        provider: str = "google",
        api_key: Optional[str] = None,
        output_mode: str = "bilingual",
        batch_size: int = 30,
        font_family: str = "auto"
    ) -> Dict[str, Any]:
        r"""
        Phương thức chính thực hiện Dịch Song Ngữ Bảng Tính Excel:
        1. Nhận diện và chuyển đổi .xls sang .xlsx nếu cần.
        2. Tải workbook với data_only=False, keep_vba=True, keep_links=True.
        3. Duyệt toàn bộ các sheet (wb.worksheets):
           - Bỏ qua các ô công thức (giữ nguyên chuỗi công thức 100%).
           - Bỏ qua các ô MergedCell (chỉ xử lý ô top-left).
           - Bỏ qua các ô số thuần, ngày tháng, boolean, rỗng.
           - Bảo vệ số CAS (\b\d{2,7}-\d{2}-\d\b) và các mã kỹ thuật.
        4. Tích hợp động cơ dịch DeepSeek AI và Free Bing Translator Engine.
        5. Ghi lại song ngữ xen kẽ bằng CellRichText chuẩn DOC-TRAN (#002060 italic).
        6. Tự động tối ưu chiều cao dòng và lưu tệp kết quả.
        """
        abs_in = os.path.abspath(input_path)
        abs_out = os.path.abspath(output_path)
        out_dir = os.path.dirname(abs_out) or "."
        os.makedirs(out_dir, exist_ok=True)

        ext = os.path.splitext(abs_in)[1].lower()
        temp_files_to_cleanup = []
        work_path = abs_in

        # 1. Xử lý tệp .xls cũ
        if ext == ".xls":
            temp_xlsx = os.path.join(out_dir, f"temp_conv_{uuid.uuid4().hex[:6]}.xlsx")
            work_path = cls.convert_xls_to_xlsx(abs_in, temp_xlsx)
            temp_files_to_cleanup.append(temp_xlsx)

        actual_key = (api_key or "").strip() or DeepSeekOCR.get_saved_api_key()

        # ── VÒNG 1: AI CONTEXT PROFILER (Phân tích bối cảnh bảng tính chạy ngầm) ──
        context_profile = AIContextProfiler.profile_document(
            work_path,
            api_key=actual_key,
            source_lang=source_lang,
            target_lang=target_lang
        )

        # 2. Nạp workbook bảo toàn 100% công thức và liên kết
        is_vba = str(work_path).lower().endswith(('.xlsm', '.xltm'))
        wb = openpyxl.load_workbook(
            work_path,
            data_only=False,
            keep_vba=is_vba,
            keep_links=True
        )

        records: List[Tuple[Any, int, int, Cell, str, ProtectedText]] = []
        seen_texts: set = set()
        unique_to_translate: List[str] = []

        # 3. Quét tất cả worksheets
        for ws in wb.worksheets:
            for row_idx, row in enumerate(ws.iter_rows(), start=1):
                for col_idx, cell in enumerate(row, start=1):
                    # Bỏ qua ô MergedCell (tránh lỗi AttributeError read-only)
                    if isinstance(cell, MergedCell):
                        continue

                    # Bỏ qua ô công thức toán học
                    if cls._is_formula_cell(cell):
                        continue

                    # Bỏ qua số, ngày tháng, boolean, ô rỗng
                    if cls._is_ignorable_value(cell.value):
                        continue
                    if getattr(cell, 'data_type', None) in ('n', 'b', 'd', 'e'):
                        continue

                    val_str = str(cell.value or "").strip()
                    if not val_str:
                        continue

                    # Kiểm tra ô có chứa từ ngữ tự nhiên cần dịch không
                    if not should_translate_item(val_str):
                        continue

                    lines = [l.strip() for l in val_str.splitlines() if l.strip()]

                    # Tự động nhận diện ô vốn đã là song ngữ (chứa các cặp dòng khác ngôn ngữ)
                    is_bilingual_cell = (
                        len(lines) >= 2
                        and len(lines) % 2 == 0
                        and all(
                            is_different_language_pair(lines[2 * k], lines[2 * k + 1])
                            for k in range(len(lines) // 2)
                        )
                    )

                    if is_bilingual_cell:
                        pairs = [(lines[2 * k], lines[2 * k + 1]) for k in range(len(lines) // 2)]
                        cell_meta = {
                            "type": "bilingual_cell",
                            "pairs": pairs
                        }
                        for src, _ in pairs:
                            if should_translate_item(src) and src not in seen_texts:
                                seen_texts.add(src)
                                unique_to_translate.append(src)
                    else:
                        cell_meta = {
                            "type": "monolingual",
                            "source": val_str
                        }
                        if val_str not in seen_texts:
                            seen_texts.add(val_str)
                            unique_to_translate.append(val_str)

                    # Lưu bản ghi ô cần dịch kèm metadata phân loại
                    records.append((ws, row_idx, col_idx, cell, val_str, cell_meta))

        # ── VÒNG 2: CONTEXT-AWARE TRANSLATION (Dịch theo ngữ cảnh) ──
        trans_map: Dict[str, str] = {}
        total_unique = len(unique_to_translate)
        engine_used = "Động cơ Bảng tính Excel DOC-TRAN (openpyxl + Bing Engine)"

        active_provider = (provider or "google").lower().strip()
        is_valid_deepseek = (
            active_provider == "deepseek"
            and actual_key
            and not str(actual_key).startswith("sk-test-")
            and len(str(actual_key).strip()) > 20
        )

        if is_valid_deepseek:
            engine_used = "Động cơ Bảng tính Excel DOC-TRAN (openpyxl + DeepSeek-V4.1-Flash)"
            for i in range(0, total_unique, batch_size):
                batch = unique_to_translate[i : i + batch_size]
                try:
                    translated = InterlinearTranslator.translate_texts_deepseek(
                        batch,
                        actual_key,
                        source_lang=source_lang,
                        target_lang=target_lang,
                        context_profile=context_profile
                    )
                    for src, res in zip(batch, translated):
                        trans_map[src] = res
                except Exception as ex:
                    logger.warning("DeepSeek batch translation failed, falling back to Free Bing: %s", ex)
                    rem_batch = unique_to_translate[i:]
                    translated_rem = InterlinearTranslator.translate_texts_free(
                        rem_batch,
                        source_lang=source_lang,
                        target_lang=target_lang
                    )
                    for s, r in zip(rem_batch, translated_rem):
                        trans_map[s] = AIContextProfiler.apply_glossary_to_text(s, r, context_profile)
                    engine_used += " (Fallback Bing)"
                    break
        else:
            for i in range(0, total_unique, batch_size):
                batch = unique_to_translate[i : i + batch_size]
                translated = InterlinearTranslator.translate_texts_free(
                    batch,
                    source_lang=source_lang,
                    target_lang=target_lang
                )
                for src, res in zip(batch, translated):
                    trans_map[src] = AIContextProfiler.apply_glossary_to_text(src, res, context_profile)

        # ── VÒNG 3: AI QUALITY AUDIT & AUTO-HEAL (Kiểm định chất lượng & Tự sửa lỗi) ──
        translation_pairs = [(src, trans_map.get(src, src)) for src in unique_to_translate]
        healed_translations, audit_report = AIQualityAuditor.audit_and_heal_pairs(
            translation_pairs,
            profile=context_profile,
            source_lang=source_lang,
            target_lang=target_lang
        )
        for src, healed in zip(unique_to_translate, healed_translations):
            trans_map[src] = healed

        # 5. Ghi kết quả song ngữ xen kẽ vào các ô bảng tính
        translated_count = 0
        for ws, row_idx, col_idx, cell, orig_text, cell_meta in records:
            meta_type = cell_meta.get("type", "monolingual")
            if meta_type == "bilingual_cell":
                pairs = cell_meta["pairs"]
                cls.apply_bilingual_replace_rich_text(
                    cell=cell,
                    pairs=pairs,
                    trans_map=trans_map,
                    output_mode=output_mode,
                    font_family=font_family,
                    source_lang=source_lang,
                    target_lang=target_lang
                )
            else:
                restored_trans = trans_map.get(orig_text, "").strip()
                if restored_trans:
                    cls.apply_interlinear_rich_text(
                        cell=cell,
                        orig_text=orig_text,
                        trans_text=restored_trans,
                        output_mode=output_mode,
                        font_family=font_family,
                        source_lang=source_lang,
                        target_lang=target_lang
                    )

                # Điều chỉnh chiều cao dòng nếu dòng đã bị gán chiều cao cố định thấp
                curr_h = ws.row_dimensions[row_idx].height
                if curr_h is not None and curr_h < 30.0:
                    ws.row_dimensions[row_idx].height = max(round(curr_h * 1.85, 1), 30.0)

                translated_count += 1

        # 6. Lưu workbook hoàn chỉnh
        wb.save(abs_out)

        # Kiểm tra tính hợp lệ của tệp Excel (.xlsx) đầu ra
        file_valid, file_err = AIQualityAuditor.verify_xlsx_integrity(abs_out)
        audit_report.file_valid = file_valid
        if not file_valid:
            audit_report.file_error = file_err
            logger.error("Cảnh báo tệp Excel đầu ra lỗi cấu trúc: %s", file_err)

        # Dọn dẹp tệp trung gian
        for tf in temp_files_to_cleanup:
            if os.path.exists(tf):
                try:
                    os.remove(tf)
                except Exception:
                    pass

        src_str = LANG_NAME_MAP.get(source_lang, source_lang)
        tgt_str = LANG_NAME_MAP.get(target_lang, target_lang)

        return {
            "success": True,
            "output_path": abs_out,
            "filename": os.path.basename(abs_out),
            "total_sheets": len(wb.worksheets),
            "sheet_count": len(wb.worksheets),
            "total_translatable_cells": len(records),
            "translated_cells": translated_count,
            "total_translated_cells": translated_count,
            "engine": engine_used,
            "source_lang": source_lang,
            "target_lang": target_lang,
            "output_mode": output_mode,
            "context_profile": context_profile.to_dict(),
            "audit_report": audit_report.to_dict(),
            "stages_completed": [
                "Vòng 1 (AI Context Profiler): Phân tích bối cảnh & lập hồ sơ thuật ngữ",
                "Vòng 2 (Context-Aware Translation): Dịch song ngữ xen kẽ theo ngữ cảnh",
                "Vòng 3 (AI Quality Audit & Auto-heal): Kiểm định chất lượng & tự sửa lỗi"
            ],
            "details": f"Đã dịch song ngữ ({src_str} ➔ {tgt_str}) {translated_count} ô bảng tính trên {len(wb.worksheets)} sheet. Bảo toàn 100% công thức toán học và cấu trúc ô gộp."
        }

    @classmethod
    def render_excel_to_html(cls, file_bytes_or_path: Union[bytes, str], max_rows: int = 150) -> Dict[str, Any]:
        """
        Kết xuất bảng tính Excel sang HTML dạng Sheet Tabs + Table phục vụ xem trước Split-View.
        """
        if isinstance(file_bytes_or_path, (str, Path)):
            wb = openpyxl.load_workbook(file_bytes_or_path, data_only=True, rich_text=True)
            filename = os.path.basename(str(file_bytes_or_path))
        else:
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes_or_path), data_only=True, rich_text=True)
            filename = "document.xlsx"

        sheets_data = []
        for ws in wb.worksheets:
            sheet_title = ws.title
            rows_html = []
            
            # Quét các hàng
            for r_idx, row in enumerate(ws.iter_rows(), start=1):
                if r_idx > max_rows:
                    rows_html.append(f"<tr><td colspan='10' class='text-muted text-center'><em>(Đã rút gọn hiển thị {max_rows} hàng đầu tiên)</em></td></tr>")
                    break

                # Bỏ qua hàng hoàn toàn rỗng
                if all(c.value is None for c in row):
                    continue

                cells_html = []
                for cell in row:
                    if isinstance(cell, MergedCell):
                        continue

                    val = cell.value
                    display_text = ""
                    cell_classes = ["excel-preview-cell"]

                    if val is None:
                        display_text = ""
                    elif isinstance(val, CellRichText):
                        parts = []
                        for block in val:
                            t = str(getattr(block, 'text', block))
                            font = getattr(block, 'font', None)
                            is_trans = False
                            if font:
                                color = getattr(font, 'color', None)
                                color_rgb = getattr(color, 'rgb', '') if color else ''
                                if '002060' in str(color_rgb) or getattr(font, 'i', False):
                                    is_trans = True
                            if is_trans:
                                parts.append(f"<div class='excel-trans-line' style='color:#002060; font-style:italic;'>{t.strip()}</div>")
                            else:
                                parts.append(f"<div class='excel-orig-line'>{t.strip()}</div>")
                        display_text = "".join(parts)
                    else:
                        display_text = str(val).replace("\n", "<br/>")

                    # Canh lề
                    align_style = ""
                    if cell.alignment:
                        if cell.alignment.horizontal:
                            align_style += f"text-align:{cell.alignment.horizontal}; "
                        if cell.alignment.vertical:
                            align_style += f"vertical-align:{cell.alignment.vertical}; "

                    cells_html.append(f"<td style='{align_style}' class='{' '.join(cell_classes)}'>{display_text}</td>")

                if cells_html:
                    rows_html.append(f"<tr>{''.join(cells_html)}</tr>")

            sheet_html = f"""
            <div class='excel-sheet-content' data-sheet='{sheet_title}'>
                <div class='table-responsive'>
                    <table class='table table-bordered table-sm excel-table'>
                        <tbody>
                            {''.join(rows_html)}
                        </tbody>
                    </table>
                </div>
            </div>
            """
            sheets_data.append({
                "title": sheet_title,
                "html": sheet_html
            })

        return {
            "success": True,
            "filename": filename,
            "sheets": [s["title"] for s in sheets_data],
            "sheets_data": sheets_data
        }
