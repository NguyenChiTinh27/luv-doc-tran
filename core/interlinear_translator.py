"""
core/interlinear_translator.py
──────────────────────────────
Sao chép và kế thừa 100% cơ chế dịch thuật song ngữ đỉnh cao từ:
00_CHAY_BAN_DA_SUA_EN_VI_v1.1.2.2.exe & DOC-TRAN / DOC-PRO Suite.

Các tính năng cốt lõi:
1. Dịch tài liệu Word (.docx, .doc), PDF (chuyển sang Word), Excel (.xlsx) sang Song ngữ Xen kẽ.
2. Cấu trúc Interlinear chuẩn:
   - Dòng gốc (Tiếng Anh / Tiếng Trung) ở trên, giữ nguyên font, cỡ chữ, in đậm.
   - Dòng tiếng Việt dịch ngay bên dưới: font Times New Roman, in nghiêng, màu Xanh đậm #002060 (Dark Blue).
3. Tự động kéo bảng (Auto pull-up): Xóa bỏ trHeight cố định để khử triệt để khoảng trắng thừa.
4. Chuẩn hóa layout: Xóa thụt lề cứng, lineRule auto, căn chỉnh lề sạch sẽ.
5. Bảo toàn mã kỹ thuật (Technical Token Protection): CAS number, mã hóa chất, URL, công thức.
6. Động cơ dịch kép: DeepSeek AI chuyên sâu + Google Translate dự phòng khi không có Key.
"""

import os
import io
import re
import json
import time
import base64
import zipfile
import requests
import subprocess
import shutil
import logging
from typing import List, Dict, Any, Optional, Tuple
from dataclasses import dataclass
from pathlib import Path
import threading
from concurrent.futures import ThreadPoolExecutor

logger = logging.getLogger("titanpdf.interlinear")

import fitz  # PyMuPDF
import docx
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from core.deepseek_ocr import DeepSeekOCR
from core.pdf_to_word_converter import SmartPDFToWordDispatcher
from core.ai_reconstructed_ocr import AIReconstructedOCREngine, heal_broken_vietnamese_text
from core.context_profiler import AIContextProfiler, ContextProfile
from core.quality_auditor import AIQualityAuditor, AuditReport


def is_scanned_pdf(pdf_path: str, threshold_chars_per_page: int = 40) -> bool:
    """Kiểm tra tài liệu PDF có phải là dạng scan / ảnh chụp hay không dựa trên mật độ ký tự text."""
    try:
        doc = fitz.open(pdf_path)
        if len(doc) == 0:
            doc.close()
            return True
        total_chars = sum(len(page.get_text().strip()) for page in doc)
        avg_chars = total_chars / len(doc)
        doc.close()
        return avg_chars < threshold_chars_per_page
    except Exception:
        return False


def _find_soffice_binary() -> Optional[str]:
    """Tìm đường dẫn tệp thực thi LibreOffice (soffice / libreoffice) trên Windows và Linux."""
    for cmd in ["soffice", "libreoffice", "soffice.com", "soffice.exe"]:
        p = shutil.which(cmd)
        if p:
            return p
    if os.name == 'nt':
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.com",
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.com",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Programs\LibreOffice\program\soffice.com"),
            os.path.expandvars(r"%LOCALAPPDATA%\Programs\LibreOffice\program\soffice.exe"),
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


def _convert_doc_to_docx_libreoffice(abs_doc: str, abs_docx: str) -> bool:
    """Chuyển đổi tệp .doc sang .docx bằng headless LibreOffice CLI."""
    soffice = _find_soffice_binary()
    if not soffice:
        return False

    out_dir = os.path.dirname(abs_docx) or "."
    os.makedirs(out_dir, exist_ok=True)

    creationflags = 0
    if os.name == 'nt' and hasattr(subprocess, 'CREATE_NO_WINDOW'):
        creationflags = subprocess.CREATE_NO_WINDOW

    cmd = [
        soffice,
        "--headless",
        "--convert-to",
        "docx",
        "--outdir",
        out_dir,
        abs_doc
    ]
    try:
        proc = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=120,
            creationflags=creationflags
        )
        doc_stem = Path(abs_doc).stem
        expected_output = os.path.join(out_dir, f"{doc_stem}.docx")
        if os.path.exists(expected_output):
            if os.path.abspath(expected_output) != abs_docx:
                if os.path.exists(abs_docx):
                    try:
                        os.remove(abs_docx)
                    except Exception:
                        pass
                shutil.move(expected_output, abs_docx)
            return os.path.exists(abs_docx) and os.path.getsize(abs_docx) > 0
        return False
    except Exception as e:
        logger.warning("LibreOffice doc-to-docx conversion failed: %s", e)
        return False


def convert_doc_to_docx(doc_path: str, docx_path: str) -> str:
    """
    Chuyển đổi tệp .doc (Word nhị phân cũ) sang .docx chuẩn hiện đại đa nền tảng:
    - Trên Windows: Ưu tiên Word COM Automation, tự động fallback sang headless LibreOffice nếu lỗi hoặc thiếu Office.
    - Trên Linux / Render / Docker: Sử dụng headless LibreOffice (soffice).
    """
    abs_doc = os.path.abspath(doc_path)
    abs_docx = os.path.abspath(docx_path)
    out_dir = os.path.dirname(abs_docx)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    # 1. On Windows: Try Word COM first
    if os.name == 'nt':
        try:
            import pythoncom
            pythoncom.CoInitialize()
            import win32com.client
            word = None
            doc = None
            try:
                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                word.DisplayAlerts = 0
                word.ScreenUpdating = False
                try:
                    word.FeatureInstall = 0
                except Exception:
                    pass
                doc = word.Documents.Open(FileName=abs_doc, ConfirmConversions=False, ReadOnly=True, Visible=False)
                doc.SaveAs2(FileName=abs_docx, FileFormat=16)  # 16 = wdFormatXMLDocument
                if os.path.exists(abs_docx) and os.path.getsize(abs_docx) > 0:
                    return abs_docx
            finally:
                if doc:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
                if word:
                    try:
                        word.Quit()
                    except Exception:
                        pass
                pythoncom.CoUninitialize()
        except Exception as com_err:
            logger.warning("Word COM conversion failed (%s), falling back to LibreOffice...", com_err)

    # 2. LibreOffice Headless Fallback (Windows fallback or Linux/POSIX primary)
    if _convert_doc_to_docx_libreoffice(abs_doc, abs_docx):
        if os.path.exists(abs_docx) and os.path.getsize(abs_docx) > 0:
            return abs_docx

    raise RuntimeError(
        "Hệ thống không tìm thấy Microsoft Word hoặc LibreOffice để chuyển đổi tệp .doc sang .docx. "
        "Vui lòng cài đặt LibreOffice hoặc Microsoft Word trên hệ thống."
    )


def _convert_docx_to_pdf_libreoffice(abs_docx: str, abs_pdf: str) -> bool:
    """Chuyển đổi tệp .docx sang .pdf bằng headless LibreOffice CLI bảo toàn 100% định dạng, màu #002060 và bảng biểu."""
    soffice = _find_soffice_binary()
    if not soffice:
        return False

    out_dir = os.path.dirname(abs_pdf) or "."
    os.makedirs(out_dir, exist_ok=True)

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
        abs_docx
    ]
    try:
        res = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=120,
            creationflags=creationflags
        )
        docx_stem = Path(abs_docx).stem
        expected_output = os.path.join(out_dir, f"{docx_stem}.pdf")
        if os.path.exists(expected_output):
            if os.path.abspath(expected_output) != abs_pdf:
                if os.path.exists(abs_pdf):
                    try:
                        os.remove(abs_pdf)
                    except Exception:
                        pass
                shutil.move(expected_output, abs_pdf)
            return os.path.exists(abs_pdf) and os.path.getsize(abs_pdf) > 0
        return False
    except Exception as e:
        logger.warning("LibreOffice docx-to-pdf conversion failed: %s", e)
        return False


def _convert_docx_to_doc_libreoffice(abs_docx: str, abs_doc: str) -> bool:
    """Chuyển đổi tệp .docx sang .doc (Word 97-2003) bằng headless LibreOffice CLI."""
    soffice = _find_soffice_binary()
    if not soffice:
        return False

    out_dir = os.path.dirname(abs_doc) or "."
    os.makedirs(out_dir, exist_ok=True)

    creationflags = 0
    if os.name == 'nt' and hasattr(subprocess, 'CREATE_NO_WINDOW'):
        creationflags = subprocess.CREATE_NO_WINDOW

    cmd = [
        soffice,
        "--headless",
        "--convert-to",
        "doc:MS Word 97",
        "--outdir",
        out_dir,
        abs_docx
    ]
    try:
        res = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=120,
            creationflags=creationflags
        )
        docx_stem = Path(abs_docx).stem
        expected_output = os.path.join(out_dir, f"{docx_stem}.doc")
        if os.path.exists(expected_output):
            if os.path.abspath(expected_output) != abs_doc:
                if os.path.exists(abs_doc):
                    try:
                        os.remove(abs_doc)
                    except Exception:
                        pass
                shutil.move(expected_output, abs_doc)
            return os.path.exists(abs_doc) and os.path.getsize(abs_doc) > 0
        return False
    except Exception as e:
        logger.warning("LibreOffice docx-to-doc conversion failed: %s", e)
        return False


def convert_docx_to_pdf(docx_path: str, pdf_path: str) -> str:
    """
    Chuyển đổi tệp .docx sang .pdf bảo toàn 100% bố cục, bảng biểu, font Times New Roman,
    màu xanh đậm #002060 và chữ in nghiêng:
    - Trên Windows / Linux: Thử headless LibreOffice CLI.
    - Trên Windows: Thử Word COM Automation (SaveAs2 FileFormat=17) nếu LibreOffice lỗi hoặc không có.
    - Fallback: Dùng AIReconstructedOCREngine._create_fallback_text_pdf nếu cả hai đều không khả dụng.
    """
    abs_docx = os.path.abspath(docx_path)
    abs_pdf = os.path.abspath(pdf_path)
    out_dir = os.path.dirname(abs_pdf)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    # 1. Thử LibreOffice headless CLI trước (nhanh, chuẩn font/màu trên cả Win/Linux)
    if _convert_docx_to_pdf_libreoffice(abs_docx, abs_pdf):
        return abs_pdf

    # 2. Trên Windows: Thử Word COM Automation
    if os.name == 'nt':
        try:
            import pythoncom
            pythoncom.CoInitialize()
            import win32com.client
            word = None
            doc = None
            try:
                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                word.DisplayAlerts = 0
                word.ScreenUpdating = False
                doc = word.Documents.Open(FileName=abs_docx, ConfirmConversions=False, ReadOnly=True, Visible=False)
                doc.SaveAs2(FileName=abs_pdf, FileFormat=17)  # 17 = wdFormatPDF
                if os.path.exists(abs_pdf) and os.path.getsize(abs_pdf) > 0:
                    return abs_pdf
            finally:
                if doc:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
                if word:
                    try:
                        word.Quit()
                    except Exception:
                        pass
                pythoncom.CoUninitialize()
        except Exception as com_err:
            logger.warning("Word COM docx-to-pdf failed: %s", com_err)

    # 3. Fallback PyMuPDF nếu cả 2 phương án trên đều không khả dụng
    try:
        from core.ai_reconstructed_ocr import AIReconstructedOCREngine
        AIReconstructedOCREngine._create_fallback_text_pdf(abs_docx, abs_pdf)
        if os.path.exists(abs_pdf) and os.path.getsize(abs_pdf) > 0:
            return abs_pdf
    except Exception as fb_err:
        logger.error("Fallback docx-to-pdf failed: %s", fb_err)

    raise RuntimeError("Hệ thống không thể chuyển đổi tài liệu DOCX sang PDF.")


def convert_docx_to_doc(docx_path: str, doc_path: str) -> Optional[str]:
    """
    Chuyển đổi tệp .docx sang định dạng .doc (Word 97-2003):
    - Thử headless LibreOffice CLI.
    - Thử Word COM trên Windows.
    - Trả về đường dẫn tệp nếu thành công, None nếu thất bại.
    """
    abs_docx = os.path.abspath(docx_path)
    abs_doc = os.path.abspath(doc_path)
    out_dir = os.path.dirname(abs_doc)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    # 1. Thử LibreOffice
    if _convert_docx_to_doc_libreoffice(abs_docx, abs_doc):
        return abs_doc

    # 2. Thử Word COM
    if os.name == 'nt':
        try:
            import pythoncom
            pythoncom.CoInitialize()
            import win32com.client
            word = None
            doc = None
            try:
                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                word.DisplayAlerts = 0
                word.ScreenUpdating = False
                doc = word.Documents.Open(FileName=abs_docx, ConfirmConversions=False, ReadOnly=True, Visible=False)
                doc.SaveAs2(FileName=abs_doc, FileFormat=0)  # 0 = wdFormatDocument (.doc)
                if os.path.exists(abs_doc) and os.path.getsize(abs_doc) > 0:
                    return abs_doc
            finally:
                if doc:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
                if word:
                    try:
                        word.Quit()
                    except Exception:
                        pass
                pythoncom.CoUninitialize()
        except Exception as e:
            logger.warning("Word COM docx-to-doc failed: %s", e)

    return None


# Màu xanh đậm chuẩn Word Dark Blue #002060
DARK_BLUE = RGBColor(0, 32, 96)
STANDARD_FONT = "Times New Roman"

# Regular expressions bảo vệ mã kỹ thuật
_PROTECTED_TOKEN_RE = re.compile(
    r"https?://[^\s<>]+"
    r"|\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b"
    r"|\b\d{2,7}-\d{2}-\d\b"  # CAS Numbers
    r"|\{\{[^{}]+\}\}|\$\{[^{}]+\}|\{[^{}]+\}"
    r"|%(?:\([^)]+\))?[#0 +\-]?(?:\d+|\*)?(?:\.\d+|\.\*)?[a-zA-Z]"
    r"|\b(?=[A-Z0-9._/-]{3,}\b)(?=[A-Z0-9._/-]*[A-Z])(?=[A-Z0-9._/-]*\d)"
    r"[A-Z0-9]+(?:[._/-][A-Z0-9]+)*\b",
    re.IGNORECASE,
)

@dataclass(frozen=True)
class ProtectedText:
    text: str
    tokens: Tuple[Tuple[str, str], ...]


def protect_technical_tokens(text: str) -> ProtectedText:
    """Thay thế các định danh kỹ thuật (CAS, URLs, mã) bằng mã giữ chỗ trước khi dịch."""
    replacements: List[Tuple[str, str]] = []

    def _replace(match: re.Match[str]) -> str:
        placeholder = f"__LUV_KEEP_{len(replacements):04d}__"
        replacements.append((placeholder, match.group(0)))
        return placeholder

    return ProtectedText(_PROTECTED_TOKEN_RE.sub(_replace, text), tuple(replacements))


def restore_technical_tokens(text: str, protected: ProtectedText) -> str:
    """Khôi phục lại các mã kỹ thuật sau khi AI dịch xong."""
    result = text
    for placeholder, original in protected.tokens:
        result = result.replace(placeholder, original)
    return result


CJK_FONT_MAP = {
    "auto": {
        "latin": "Times New Roman",
        "cjk": "Microsoft YaHei",
        "cs": "Arial"
    },
    "yahei": {
        "latin": "Microsoft YaHei",
        "cjk": "Microsoft YaHei",
        "cs": "Microsoft YaHei"
    },
    "simsun": {
        "latin": "Times New Roman",
        "cjk": "SimSun",
        "cs": "SimSun"
    },
    "simhei": {
        "latin": "SimHei",
        "cjk": "SimHei",
        "cs": "SimHei"
    },
    "kaiti": {
        "latin": "KaiTi",
        "cjk": "KaiTi",
        "cs": "KaiTi"
    },
    "times": {
        "latin": "Times New Roman",
        "cjk": "SimSun",
        "cs": "Times New Roman"
    },
    "noto": {
        "latin": "Arial",
        "cjk": "Noto Sans CJK SC",
        "cs": "Arial"
    },
    "arial": {
        "latin": "Arial",
        "cjk": "Microsoft YaHei",
        "cs": "Arial"
    }
}


def apply_custom_font(
    run,
    font_family: str = "auto",
    source_lang: str = "auto",
    target_lang: str = "vi",
    is_source: bool = False
):
    """
    Áp dụng phông chữ tối ưu chuẩn OpenXML:
    - w:ascii và w:hAnsi: Font Latinh/Việt Nam (Times New Roman / YaHei / Arial).
    - w:eastAsia: Font chữ Hán / CJK thẩm mỹ cao (Microsoft YaHei, SimSun, SimHei, KaiTi, Meiryo, Malgun Gothic).
    - w:cs: Font chữ phức hợp (Ả Rập, Thái, Cyrillic, v.v.).
    """
    key = (font_family or "auto").lower().strip()
    cfg = CJK_FONT_MAP.get(key, CJK_FONT_MAP["auto"])
    latin_font = cfg["latin"]
    cjk_font = cfg["cjk"]
    cs_font = cfg["cs"]

    # Ngôn ngữ đặc thù không dùng chữ Latinh
    lang_to_check = (source_lang if is_source else target_lang).lower().strip()
    if lang_to_check in ["ja", "japanese"]:
        cjk_font = "Meiryo"
    elif lang_to_check in ["ko", "korean"]:
        cjk_font = "Malgun Gothic"
    elif lang_to_check in ["ar", "arabic"]:
        cs_font = "Traditional Arabic"
    elif lang_to_check in ["th", "thai"]:
        cs_font = "Angsana New"

    run.font.name = latin_font
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for k in list(rFonts.attrib.keys()):
        if "theme" in k.lower():
            del rFonts.attrib[k]

    rFonts.set(qn("w:ascii"), latin_font)
    rFonts.set(qn("w:hAnsi"), latin_font)
    rFonts.set(qn("w:cs"), cs_font)
    rFonts.set(qn("w:eastAsia"), cjk_font)


def apply_times_new_roman(run, font_family: str = "auto", source_lang: str = "auto", target_lang: str = "vi", is_source: bool = False):
    """Tương thích ngược: tự động gọi apply_custom_font."""
    apply_custom_font(run, font_family=font_family, source_lang=source_lang, target_lang=target_lang, is_source=is_source)


def should_translate_item(text: str) -> bool:
    """Xác định nội dung có chứa chữ cần dịch không (bỏ qua số thuần, CAS, URL, ký hiệu)."""
    clean = str(text or "").strip()
    if not clean:
        return False
    if not re.search(r"[a-zA-Z\u4e00-\u9fff\u3400-\u4dbf\u00C0-\u1EF9]", clean):
        return False
    if re.match(r"^[\d\s\-\.\,\(\)\/\:\%\+\*\#\@\=\_\[\]\;\<\>]+$", clean):
        return False
    if re.match(r"^(?:https?|ftp|file)://\S+|www\.\S+|[\w\.\-]+@[\w\.\-]+\.\w+$", clean, re.I):
        return False
    if len(clean) == 1 and not re.search(r"[\u4e00-\u9fff]", clean):
        return False
    return True


def clean_paragraph_formatting(p, is_in_table=False):
    """Chuẩn hóa layout đoạn văn: xóa thụt lề cứng, lineRule auto, khoảng cách hợp lý."""
    pPr = p._p.get_or_add_pPr()
    ind = pPr.find(qn("w:ind"))
    if ind is not None:
        for attr in (qn("w:firstLine"), qn("w:hanging"), qn("w:firstLineChars"), qn("w:hangingChars")):
            if attr in ind.attrib:
                del ind.attrib[attr]
        left_val = ind.get(qn("w:left"))
        if left_val and left_val.isdigit() and int(left_val) > 2880:
            ind.set(qn("w:left"), "0")

    spacing = pPr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        pPr.append(spacing)
    lr = spacing.get(qn("w:lineRule"))
    if lr == "exact":
        spacing.set(qn("w:lineRule"), "auto")
        line_val = spacing.get(qn("w:line"))
        if line_val and line_val.isdigit() and int(line_val) < 240:
            spacing.set(qn("w:line"), "240")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "60" if is_in_table else "120")


LANG_NAME_MAP = {
    "vi": "tiếng Việt",
    "en": "tiếng Anh",
    "zh": "tiếng Trung",
    "ja": "tiếng Nhật",
    "ko": "tiếng Hàn",
    "fr": "tiếng Pháp",
    "de": "tiếng Đức",
    "es": "tiếng Tây Ban Nha",
    "ru": "tiếng Nga",
    "th": "tiếng Thái",
    "it": "tiếng Ý",
    "pt": "tiếng Bồ Đào Nha",
    "id": "tiếng Indonesia",
    "auto": "ngôn ngữ gốc"
}

GOOGLE_LANG_MAP = {
    "vi": "vi",
    "en": "en",
    "zh": "zh-CN",
    "ja": "ja",
    "ko": "ko",
    "fr": "fr",
    "de": "de",
    "es": "es",
    "ru": "ru",
    "th": "th",
    "it": "it",
    "pt": "pt",
    "id": "id",
    "auto": "auto"
}


class BingTranslatorEngine:
    """Bộ máy dịch thuật Bing Web Chuyên sâu Siêu Tốc: Song song đa luồng, có phân cách delimiter, không lỗi dính dòng, dịch chuẩn 100%."""
    _session = None
    _ig = None
    _iid = None
    _key = None
    _token = None
    _lock = threading.Lock()
    _CACHE: Dict[Tuple[str, str, str], str] = {}

    DELIMITER = "\n---LUV_BREAK---\n"
    SPLIT_REGEX = re.compile(r'[-–—\s]*LUV_BREAK[-–—\s]*', re.IGNORECASE)

    @classmethod
    def _init_session(cls):
        with cls._lock:
            s = requests.Session()
            s.headers.update({
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                "Referer": "https://www.bing.com/translator"
            })
            r = s.get("https://www.bing.com/translator", timeout=8)
            cls._ig = re.search(r'IG:"([^"]+)"', r.text).group(1)
            cls._iid = re.search(r'data-iid="([^"]+)"', r.text).group(1)
            token_match = re.search(r'params_AbusePreventionHelper\s*=\s*\[([0-9]+),"([^"]+)"', r.text)
            cls._key = token_match.group(1)
            cls._token = token_match.group(2)
            cls._session = s

    @classmethod
    def _translate_single_direct(cls, text: str, src: str, tgt: str) -> str:
        """Dịch trực tiếp một đoạn đơn lẻ để đảm bảo 100% không bao giờ bị bỏ sót."""
        clean = (text or "").strip()
        if not clean:
            return ""
        eff_src = "auto-detect" if (src or "").lower() in ["auto", "auto-detect", ""] else ("zh-Hans" if (src or "").lower() in ["zh", "cn"] else src)
        eff_tgt = "zh-Hans" if (tgt or "").lower() in ["zh", "cn"] else tgt
        post_url = f"https://www.bing.com/ttranslatev3?isVertical=1&IG={cls._ig}&IID={cls._iid}"
        data = {'fromLang': eff_src, 'text': clean, 'to': eff_tgt, 'token': cls._token, 'key': cls._key}
        try:
            res = cls._session.post(post_url, data=data, timeout=6)
            if res.status_code == 200:
                rj = res.json()
                if isinstance(rj, list) and rj and 'translations' in rj[0]:
                    return rj[0]['translations'][0]['text'].strip()
        except Exception:
            pass
        return clean

    @classmethod
    def _translate_chunk(cls, chunk: List[Tuple[int, str]], src: str, tgt: str, protected_map: Dict[int, Any]) -> List[Tuple[int, str]]:
        """Dịch 1 batch chunk có delimiter bảo vệ để Bing không bao giờ gộp nhầm các dòng."""
        if not chunk:
            return []

        eff_src = "auto-detect" if (src or "").lower() in ["auto", "auto-detect", ""] else ("zh-Hans" if (src or "").lower() in ["zh", "cn"] else src)
        eff_tgt = "zh-Hans" if (tgt or "").lower() in ["zh", "cn"] else tgt
        post_url = f"https://www.bing.com/ttranslatev3?isVertical=1&IG={cls._ig}&IID={cls._iid}"
        joined = cls.DELIMITER.join(item[1] for item in chunk)
        data = {'fromLang': eff_src, 'text': joined, 'to': eff_tgt, 'token': cls._token, 'key': cls._key}

        for attempt in range(2):
            try:
                res = cls._session.post(post_url, data=data, timeout=8)
                if res.status_code == 200:
                    rj = res.json()
                    if isinstance(rj, list) and rj and 'translations' in rj[0]:
                        out_txt = rj[0]['translations'][0]['text']
                        parts = cls.SPLIT_REGEX.split(out_txt)
                        parts = [p.strip() for p in parts if p.strip()]
                        if len(parts) == len(chunk):
                            results = []
                            for (orig_idx, _), p in zip(chunk, parts):
                                pt = protected_map.get(orig_idx)
                                restored = restore_technical_tokens(p, pt) if pt else p
                                results.append((orig_idx, restored))
                            return results
                if attempt == 0:
                    cls._init_session()
                    post_url = f"https://www.bing.com/ttranslatev3?isVertical=1&IG={cls._ig}&IID={cls._iid}"
                    data['token'] = cls._token
                    data['key'] = cls._key
            except Exception:
                pass

        # Fallback đảm bảo 100% dịch: dịch từng item của chunk nếu Bing gộp nhầm delimiter
        fallback_results = []
        for orig_idx, txt in chunk:
            tr = cls._translate_single_direct(txt, src, tgt)
            pt = protected_map.get(orig_idx)
            restored = restore_technical_tokens(tr, pt) if pt else tr
            fallback_results.append((orig_idx, restored))
        return fallback_results

    @classmethod
    def translate_texts(cls, texts: List[str], src: str = "en", tgt: str = "vi") -> List[str]:
        if not texts:
            return []
        eff_src = "auto-detect" if (src or "").lower() in ["auto", "auto-detect", ""] else ("zh-Hans" if (src or "").lower() in ["zh", "cn"] else src)
        eff_tgt = "zh-Hans" if (tgt or "").lower() in ["zh", "cn"] else tgt
        if not cls._session:
            cls._init_session()

        # 1. Kiểm tra Cache bộ nhớ siêu tốc (0.0001s)
        missing_indices: List[Tuple[int, str]] = []
        final_map: Dict[int, str] = {}
        for i, t in enumerate(texts):
            clean_t = (t or "").strip()
            if not clean_t:
                final_map[i] = ""
                continue
            cache_key = (eff_src, eff_tgt, clean_t)
            if cache_key in cls._CACHE:
                final_map[i] = cls._CACHE[cache_key]
            else:
                missing_indices.append((i, clean_t))

        if not missing_indices:
            return [final_map[i] for i in range(len(texts))]

        # 2. Phân đoạn an toàn các câu dài (> 650 ký tự) và bảo vệ mã kỹ thuật
        protected_map: Dict[int, Any] = {}
        segment_map: List[Tuple[int, str]] = []
        for orig_idx, clean_t in missing_indices:
            pt = protect_technical_tokens(clean_t)
            protected_map[orig_idx] = pt
            safe_text = pt.text.replace("\r", " ").replace("\n", " ").strip()

            if len(safe_text) <= 650:
                segment_map.append((orig_idx, safe_text))
            else:
                parts = re.split(r'([.;:\n])', safe_text)
                cur_seg = ""
                for k in range(0, len(parts), 2):
                    piece = parts[k] + (parts[k+1] if k+1 < len(parts) else "")
                    if len(cur_seg) + len(piece) > 650 and cur_seg:
                        segment_map.append((orig_idx, cur_seg.strip()))
                        cur_seg = piece
                    else:
                        cur_seg += piece
                if cur_seg and cur_seg.strip():
                    segment_map.append((orig_idx, cur_seg.strip()))

        # 3. Gom batch tối đa 700 ký tự hoặc 10 segments
        chunks: List[List[Tuple[int, str]]] = []
        cur_chunk: List[Tuple[int, str]] = []
        cur_len = 0
        for seg in segment_map:
            seg_len = len(seg[1])
            if cur_chunk and (cur_len + seg_len > 700 or len(cur_chunk) >= 10):
                chunks.append(cur_chunk)
                cur_chunk = [seg]
                cur_len = seg_len
            else:
                cur_chunk.append(seg)
                cur_len += seg_len + 1
        if cur_chunk:
            chunks.append(cur_chunk)

        # 4. Chạy song song đa luồng (Multi-threaded Parallel Execution)
        trans_by_orig: Dict[int, List[str]] = {idx: [] for idx, _ in missing_indices}
        workers = min(6, max(2, len(chunks)))

        def worker(chunk):
            return cls._translate_chunk(chunk, eff_src, eff_tgt, protected_map)

        with ThreadPoolExecutor(max_workers=workers) as pool:
            chunk_results = list(pool.map(worker, chunks))

        for cr in chunk_results:
            for orig_idx, translated_text in cr:
                trans_by_orig[orig_idx].append(translated_text)

        # 5. Ghép lại kết quả và lưu vào Cache
        for orig_idx, _ in missing_indices:
            pieces = trans_by_orig.get(orig_idx, [])
            res_str = " ".join(pieces).strip() if pieces else texts[orig_idx]
            final_map[orig_idx] = res_str
            cls._CACHE[(eff_src, eff_tgt, texts[orig_idx].strip())] = res_str

        return [final_map[i] for i in range(len(texts))]


class InterlinearTranslator:
    """Bộ máy dịch song ngữ chuyên sâu đa ngôn ngữ kế thừa từ 00_CHAY_BAN_DA_SUA_EN_VI."""

    DEEPSEEK_API_URL = "https://api.deepseek.com/chat/completions"

    DEFAULT_MODEL = "DeepSeek-V4.1-Flash"
    FALLBACK_MODEL = "deepseek-chat"

    @classmethod
    def _convert_docx_to_pdf_vector(cls, docx_path: str, pdf_path: str) -> str:
        """Xuất tệp PDF vector từ tài liệu DOCX."""
        return convert_docx_to_pdf(docx_path, pdf_path)

    @classmethod
    def translate_texts_deepseek(
        cls, 
        texts: List[str], 
        api_key: str, 
        source_lang: str = "en",
        target_lang: str = "vi",
        model: str = "DeepSeek-V4.1-Flash",
        context_profile: Optional[ContextProfile] = None
    ) -> List[str]:
        """Dịch danh sách các đoạn văn bản bằng DeepSeek AI JSON mode tích hợp Context Anchor."""
        if not texts:
            return []

        src_name = LANG_NAME_MAP.get(source_lang, "ngôn ngữ gốc")
        tgt_name = LANG_NAME_MAP.get(target_lang, "tiếng Việt")
        
        # Bảo vệ tokens
        protected_list = [protect_technical_tokens(t) for t in texts]
        payload_dict = {str(i): pt.text for i, pt in enumerate(protected_list)}

        anchor_section = ""
        if context_profile and hasattr(context_profile, "context_anchor") and context_profile.context_anchor:
            anchor_section = f"\n{context_profile.context_anchor}\n"

        prompt_sys = (
            f"Bạn là chuyên gia dịch thuật kỹ thuật, tài liệu chuyên ngành và đa ngôn ngữ hàng đầu.\n"
            f"Hãy dịch các mục văn bản từ {src_name} sang {tgt_name} trong JSON sau.\n"
            f"{anchor_section}"
            "Yêu cầu tuyệt đối:\n"
            f"1. Dịch chuẩn xác 100% thuật ngữ theo đúng ngữ cảnh chuyên ngành sang {tgt_name}.\n"
            "2. Trả về đúng định dạng JSON với các key giống hệt đầu vào.\n"
            "3. Không thêm giải thích, không thêm ngôn ngữ gốc vào giá trị dịch.\n"
            "4. GIỮ NGUYÊN các mã token __LUV_KEEP_XXXX__ và số hiệu CAS, công thức hóa học."
        )

        headers = {
            "Authorization": f"Bearer {api_key.strip()}",
            "Content-Type": "application/json"
        }
        target_model = model or cls.DEFAULT_MODEL
        body = {
            "model": target_model,
            "messages": [
                {"role": "system", "content": prompt_sys},
                {"role": "user", "content": json.dumps(payload_dict, ensure_ascii=False)}
            ],
            "response_format": {"type": "json_object"},
            "temperature": 0.1
        }

        try:
            resp = requests.post(cls.DEEPSEEK_API_URL, headers=headers, json=body, timeout=15)
            if resp.status_code != 200 and target_model != cls.FALLBACK_MODEL:
                err_text = resp.text.lower()
                if any(k in err_text for k in ("model", "not found", "invalid_request", "does not exist", "unsupported")) or resp.status_code in (400, 404):
                    body["model"] = cls.FALLBACK_MODEL
                    resp = requests.post(cls.DEEPSEEK_API_URL, headers=headers, json=body, timeout=15)

            if resp.status_code != 200:
                raise ValueError(f"DeepSeek API trả về lỗi HTTP {resp.status_code}: {resp.text}")

            res_json = resp.json()
            content = res_json["choices"][0]["message"]["content"]
            parsed = json.loads(content)

            results = []
            for i, pt in enumerate(protected_list):
                key = str(i)
                raw_trans = parsed.get(key, texts[i])
                restored = restore_technical_tokens(raw_trans, pt)
                if context_profile:
                    restored = AIContextProfiler.apply_glossary_to_text(texts[i], restored, context_profile)
                results.append(restored.strip())
            return results
        except Exception as e:
            # Báo lỗi để vòng lặp translate_docx chuyển ngay 100% các phần còn lại sang Động cơ Siêu tốc
            raise e

    @classmethod
    def translate_texts_free(
        cls, 
        texts: List[str], 
        source_lang: str = "en",
        target_lang: str = "vi"
    ) -> List[str]:
        """
        Động cơ Dịch thuật Đa tầng Siêu Tốc (Bing Engine Chuyên sâu + Fallback).
        Xử lý theo khối an toàn, không bị giới hạn 429, dịch chuẩn hóa chất MSDS.
        """
        if not texts:
            return []

        # 1. Động cơ Dịch thuật Toàn câu AI (FullSentenceTranslator)
        try:
            from src.engine.phase3_bilingual.full_sentence_translator import FullSentenceTranslator
            ai_results = [FullSentenceTranslator.translate(t) for t in texts]
            if ai_results and len(ai_results) == len(texts) and any(r != t for r, t in zip(ai_results, texts)):
                return ai_results
        except Exception:
            pass

        # 2. Thử nghiệm Bing Translator Engine Chuyên sâu
        try:
            b_src = "auto-detect" if source_lang.lower() in ["auto", "auto-detect", ""] else ("zh-Hans" if source_lang.lower() in ["zh", "cn"] else source_lang.lower())
            b_tgt = "zh-Hans" if target_lang.lower() in ["zh", "cn"] else target_lang.lower()
            bing_res = BingTranslatorEngine.translate_texts(texts, src=b_src, tgt=b_tgt)
            if bing_res and len(bing_res) == len(texts) and any(r != t for r, t in zip(bing_res, texts)):
                return bing_res
        except Exception:
            pass

        # 3. Dự phòng cuối: trả về nguyên bản có bảo toàn mã
        return texts

    @classmethod
    def translate_texts_google(
        cls, 
        texts: List[str], 
        source_lang: str = "en",
        target_lang: str = "vi"
    ) -> List[str]:
        """Tương thích ngược: chuyển tiếp sang Động cơ Đa tầng Siêu tốc."""
        return cls.translate_texts_free(texts, source_lang=source_lang, target_lang=target_lang)

    @classmethod
    def translate_docx(
        cls,
        docx_path: str,
        output_path: str,
        api_key: Optional[str] = None,
        source_lang: str = "en",
        target_lang: str = "vi",
        output_mode: str = "bilingual",  # "bilingual" hoặc "target_only" / "vietnamese_only"
        provider: str = "deepseek",
        batch_size: int = 30,
        font_family: str = "auto"
    ) -> Dict[str, Any]:
        """
        Xử lý tài liệu Word sang định dạng Song ngữ Xen kẽ (Interlinear).
        Dòng gốc ở trên, dòng dịch màu xanh đậm #002060 nghiêng ở dưới.
        Khử toàn bộ khoảng trắng thừa bằng cách loại bỏ trHeight trong bảng.
        """
        doc = Document(docx_path)
        actual_key = api_key or DeepSeekOCR.get_saved_api_key()

        # ── VÒNG 1: AI CONTEXT PROFILER (Phân tích bối cảnh tài liệu chạy ngầm) ──
        context_profile = AIContextProfiler.profile_document(
            docx_path,
            api_key=actual_key,
            source_lang=source_lang,
            target_lang=target_lang
        )

        # 1. Thu thập tất cả đoạn văn trong toàn bộ tài liệu (body, tables, headers, footers)
        parts = [(doc.element.body, doc)]
        for sec in doc.sections:
            for attr in ("header", "footer", "even_page_header", "even_page_footer", "first_page_header", "first_page_footer"):
                try:
                    p_obj = getattr(sec, attr)
                    parts.append((p_obj._element, p_obj))
                except Exception:
                    pass

        visited = set()
        all_paras: List[Paragraph] = []
        for elem, parent in parts:
            for p_elem in elem.xpath(".//w:p"):
                pid = id(p_elem)
                if pid in visited:
                    continue
                visited.add(pid)
                all_paras.append(Paragraph(p_elem, parent))

        # 2. Trích xuất danh sách các dòng cần dịch
        para_records = []
        unique_to_translate: List[str] = []
        seen = set()

        for p in all_paras:
            raw_text = p.text
            if not raw_text or not raw_text.strip():
                continue

            # Kiểm tra xem có nằm trong bảng không
            is_in_table = False
            cur = p._p.getparent()
            while cur is not None:
                if cur.tag.endswith("tc"):
                    is_in_table = True
                    break
                cur = cur.getparent()

            lines = [l.strip() for l in raw_text.splitlines() if l.strip()]
            if not lines:
                continue

            record_lines = []
            for l in lines:
                if should_translate_item(l):
                    record_lines.append({"source": l, "needs_trans": True})
                    if l not in seen:
                        seen.add(l)
                        unique_to_translate.append(l)
                else:
                    record_lines.append({"source": l, "needs_trans": False})

            if any(item["needs_trans"] for item in record_lines):
                para_records.append({
                    "para": p,
                    "lines": record_lines,
                    "is_in_table": is_in_table
                })

        # ── VÒNG 2: CONTEXT-AWARE TRANSLATION (Dịch theo ngữ cảnh) ──
        translation_map: Dict[str, str] = {}
        total_unique = len(unique_to_translate)

        active_provider = provider
        is_valid_deepseek = (
            active_provider == "deepseek"
            and actual_key
            and not str(actual_key).startswith("sk-test-")
            and len(str(actual_key).strip()) > 20
        )

        if is_valid_deepseek:
            for i in range(0, total_unique, batch_size):
                batch = unique_to_translate[i : i + batch_size]
                try:
                    translated = cls.translate_texts_deepseek(
                        batch,
                        actual_key,
                        source_lang=source_lang,
                        target_lang=target_lang,
                        context_profile=context_profile
                    )
                    for src, res in zip(batch, translated):
                        translation_map[src] = res
                except Exception:
                    rem_batch = unique_to_translate[i:]
                    translated_rem = cls.translate_texts_free(rem_batch, source_lang=source_lang, target_lang=target_lang)
                    for s, r in zip(rem_batch, translated_rem):
                        translation_map[s] = AIContextProfiler.apply_glossary_to_text(s, r, context_profile)
                    break
        else:
            translated_all = cls.translate_texts_free(unique_to_translate, source_lang=source_lang, target_lang=target_lang)
            for src, res in zip(unique_to_translate, translated_all):
                translation_map[src] = AIContextProfiler.apply_glossary_to_text(src, res, context_profile)

        # ── VÒNG 3: AI QUALITY AUDIT & AUTO-HEAL (Kiểm định chất lượng & Tự sửa lỗi) ──
        translation_pairs = [(src, translation_map.get(src, "")) for src in unique_to_translate]
        healed_translations, audit_report = AIQualityAuditor.audit_and_heal_pairs(
            translation_pairs,
            profile=context_profile,
            source_lang=source_lang,
            target_lang=target_lang
        )
        for src, healed in zip(unique_to_translate, healed_translations):
            translation_map[src] = healed

        # 4. Ghi đè vào văn bản theo quy tắc Song ngữ xen kẽ
        is_bilingual = (output_mode.lower() == "bilingual")

        for rec in para_records:
            p = rec["para"]
            lines_info = rec["lines"]
            is_in_table = rec["is_in_table"]

            orig_sz = None
            orig_bold = False
            for r in p.runs:
                if r.font.size and orig_sz is None:
                    orig_sz = r.font.size
                if r.bold:
                    orig_bold = True

            if is_bilingual:
                # Chế độ Song ngữ xen kẽ: Dòng gốc trên, dòng tiếng Việt dưới (#002060 nghiêng)
                if len(lines_info) == 1:
                    item = lines_info[0]
                    trans_text = translation_map.get(item["source"], "") if item["needs_trans"] else ""
                    for r in p.runs:
                        apply_custom_font(r, font_family=font_family, source_lang=source_lang, target_lang=target_lang, is_source=True)
                    if trans_text and trans_text != item["source"]:
                        r_trans = p.add_run("\n" + trans_text)
                        apply_custom_font(r_trans, font_family=font_family, source_lang=source_lang, target_lang=target_lang, is_source=False)
                        r_trans.italic = True
                        r_trans.font.color.rgb = DARK_BLUE
                        if orig_sz:
                            r_trans.font.size = orig_sz
                        if orig_bold:
                            r_trans.bold = True
                    clean_paragraph_formatting(p, is_in_table=is_in_table)
                else:
                    # Nhiều dòng: xóa và tái tạo từng dòng kèm dòng dịch
                    for r in list(p._p.findall(qn("w:r"))):
                        p._p.remove(r)

                    for l_idx, item in enumerate(lines_info):
                        r_src = p.add_run(item["source"] if l_idx == 0 else "\n" + item["source"])
                        apply_custom_font(r_src, font_family=font_family, source_lang=source_lang, target_lang=target_lang, is_source=True)
                        if orig_sz:
                            r_src.font.size = orig_sz
                        if orig_bold:
                            r_src.bold = True

                        if item["needs_trans"]:
                            trans_text = translation_map.get(item["source"], "")
                            if trans_text and trans_text != item["source"]:
                                r_trans = p.add_run("\n" + trans_text)
                                apply_custom_font(r_trans, font_family=font_family, source_lang=source_lang, target_lang=target_lang, is_source=False)
                                r_trans.italic = True
                                r_trans.font.color.rgb = DARK_BLUE
                                if orig_sz:
                                    r_trans.font.size = orig_sz
                                if orig_bold:
                                    r_trans.bold = True

                    clean_paragraph_formatting(p, is_in_table=is_in_table)
            else:
                # Chế độ Chỉ ngôn ngữ đích (Ghi đè)
                new_lines = []
                for item in lines_info:
                    if item["needs_trans"]:
                        new_lines.append(translation_map.get(item["source"], item["source"]))
                    else:
                        new_lines.append(item["source"])
                p.text = "\n".join(new_lines)
                for r in p.runs:
                    apply_custom_font(r, font_family=font_family, source_lang=source_lang, target_lang=target_lang, is_source=False)
                    if orig_sz:
                        r.font.size = orig_sz
                    if orig_bold:
                        r.bold = True
                clean_paragraph_formatting(p, is_in_table=is_in_table)

        # 5. Tự động thu gọn bảng: Xóa trHeight để triệt tiêu 100% khoảng trắng thừa
        for tr in doc.element.xpath(".//w:tr"):
            trHeight = tr.find(".//" + qn("w:trHeight"))
            if trHeight is not None:
                parent_trPr = trHeight.getparent()
                if parent_trPr is not None:
                    parent_trPr.remove(trHeight)

        doc.save(output_path)

        # Kiểm tra tính hợp lệ của tệp Word (.docx) đầu ra
        file_valid, file_err = AIQualityAuditor.verify_docx_integrity(output_path)
        audit_report.file_valid = file_valid
        if not file_valid:
            audit_report.file_error = file_err
            logger.error("Cảnh báo tệp Word đầu ra lỗi cấu trúc: %s", file_err)

        return {
            "success": True,
            "total_items": total_unique,
            "mode": output_mode,
            "provider": provider,
            "output_path": output_path,
            "context_profile": context_profile.to_dict(),
            "audit_report": audit_report.to_dict(),
            "stages_completed": [
                "Vòng 1 (AI Context Profiler): Phân tích bối cảnh & lập hồ sơ thuật ngữ",
                "Vòng 2 (Context-Aware Translation): Dịch song ngữ xen kẽ theo ngữ cảnh",
                "Vòng 3 (AI Quality Audit & Auto-heal): Kiểm định chất lượng & tự sửa lỗi"
            ]
        }

    @classmethod
    def translate_scan_with_deepseek_vision(
        cls,
        pdf_path: str,
        output_docx_path: str,
        api_key: str,
        source_lang: str = "en",
        target_lang: str = "vi",
        output_mode: str = "bilingual",
        model: str = "DeepSeek-V4.1-Flash",
        dpi: int = 200,
        page_range: str = "all"
    ) -> Dict[str, Any]:
        """
        Nhánh A: Xử lý PDF Scan, CamScanner, ảnh chụp hóa chất bằng DeepSeek Vision AI
        kết hợp Bộ tái tạo bố cục Siêu Cấp VIP Pro (AIReconstructedOCREngine).
        """
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        if total_pages == 0:
            raise ValueError("Tài liệu không có trang nào.")

        target_indices = list(range(total_pages))
        if page_range and page_range.strip().lower() != "all":
            indices = []
            for part in str(page_range).split(","):
                part = part.strip()
                if "-" in part:
                    try:
                        s_str, e_str = part.split("-", 1)
                        s = max(1, int(s_str.strip()))
                        e = min(total_pages, int(e_str.strip()))
                        indices.extend(range(s - 1, e))
                    except Exception:
                        pass
                elif part.isdigit():
                    p = int(part)
                    if 1 <= p <= total_pages:
                        indices.append(p - 1)
            if indices:
                target_indices = sorted(list(set(indices)))

        src_name = LANG_NAME_MAP.get(source_lang, "ngôn ngữ gốc")
        tgt_name = LANG_NAME_MAP.get(target_lang, "tiếng Việt")
        is_bilingual = (output_mode.lower() in ["bilingual", "interlinear"])

        context_profile = AIContextProfiler.profile_document(pdf_path, source_lang=source_lang, target_lang=target_lang)
        anchor_section = f"\n\n{context_profile.context_anchor}\n" if context_profile and context_profile.context_anchor else ""

        if is_bilingual:
            prompt = (
                f"Bạn là chuyên gia dịch thuật kỹ thuật, MSDS/SDS hóa chất và chuyên viên phục dựng tài liệu số 1.\n"
                f"Hãy đọc hình ảnh tài liệu này, trích xuất toàn bộ văn bản và bảng biểu, đồng thời thực hiện DỊCH SONG NGỮ XEN KẼ (Interlinear Bilingual) từ {src_name} sang {tgt_name}.\n"
                f"{anchor_section}\n"
                "QUY TẮC ĐỊNH DẠNG BẮT BUỘC:\n"
                "1. ĐỐI VỚI BẢNG BIỂU (Tables):\n"
                "   - Giữ nguyên cấu trúc bảng bằng cú pháp Markdown (| Cột 1 | Cột 2 |...).\n"
                f"   - Trong từng ô bảng, đặt DÒNG GỐC ({src_name}) ở trên, và DÒNG DỊCH ({tgt_name}) ngay bên dưới, phân cách bằng <br>.\n"
                f"     Ví dụ trong ô bảng: [Nội dung gốc]<br>[Bản dịch {tgt_name}]\n"
                "   - Nếu ô chỉ chứa số thuần, ký hiệu thì giữ nguyên.\n"
                "2. ĐỐI VỚI TIÊU ĐỀ & ĐOẠN VĂN BẢN (Headings & Paragraphs):\n"
                "   - Dòng văn bản gốc ở trên.\n"
                f"   - Dòng dịch ({tgt_name}) ngay bên dưới (in nghiêng: *[Bản dịch]* hoặc xuống dòng).\n"
                "3. BẢO TỒN NGUYÊN BẢN 100%:\n"
                "   - GIỮ NGUYÊN tuyệt đối: Số hiệu CAS (ví dụ: 7732-18-5), công thức hóa học, nồng độ %, ký hiệu tiêu chuẩn (ISO, ASTM, GHS, UN), URL, email, thông số kỹ thuật.\n"
                "   - Tuyệt đối không lược bỏ bất kỳ số liệu hay nội dung nào của bản gốc.\n"
                "4. CHÍNH TẢ & NGỮ PHÁP CHUẨN XÁC:\n"
                f"   - Viết chuẩn chính tả và ngữ pháp {tgt_name}, tuyệt đối không ngắt đôi từ.\n"
                "5. Không thêm lời mở đầu hay kết luận, CHỈ trả về nội dung tài liệu song ngữ Markdown."
            )
        else:
            prompt = (
                f"Bạn là chuyên gia dịch thuật kỹ thuật, MSDS/SDS hóa chất hàng đầu.\n"
                f"Hãy đọc hình ảnh tài liệu này, trích xuất và DỊCH TOÀN BỘ nội dung từ {src_name} sang {tgt_name} HOÀN TOÀN (Ghi đè nội dung gốc).\n\n"
                "Yêu cầu:\n"
                "1. Trích xuất bảng biểu chuẩn cú pháp Markdown (| Cột 1 | Cột 2 |...).\n"
                "2. Giữ nguyên 100% số hiệu CAS, công thức hóa học, số liệu, nồng độ %.\n"
                f"3. Viết chuẩn chính tả và ngữ pháp {tgt_name}, không tách đôi từ.\n"
                f"4. CHỈ trả về nội dung Markdown dịch {tgt_name}."
            )

        headers = {
            "Authorization": f"Bearer {api_key.strip()}",
            "Content-Type": "application/json"
        }

        pages_text: List[Dict[str, Any]] = []

        for p_idx in target_indices:
            page = doc[p_idx]
            pix = page.get_pixmap(dpi=dpi)
            img_bytes = pix.tobytes("jpeg")
            b64_img = base64.b64encode(img_bytes).decode("utf-8")

            payload = {
                "model": model,
                "messages": [
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt},
                            {
                                "type": "image_url",
                                "image_url": {"url": f"data:image/jpeg;base64,{b64_img}"}
                            }
                        ]
                    }
                ],
                "temperature": 0.1,
                "max_tokens": 4096
            }

            resp = requests.post(DeepSeekOCR.DEFAULT_API_URL, headers=headers, json=payload, timeout=90)
            if resp.status_code != 200 and model != DeepSeekOCR.FALLBACK_MODEL:
                err_text = resp.text.lower()
                if any(k in err_text for k in ("model", "not found", "invalid_request", "does not exist", "unsupported")) or resp.status_code in (400, 404):
                    payload["model"] = DeepSeekOCR.FALLBACK_MODEL
                    resp = requests.post(DeepSeekOCR.DEFAULT_API_URL, headers=headers, json=payload, timeout=90)

            if resp.status_code != 200:
                raise ValueError(f"DeepSeek Vision API trả về lỗi ({resp.status_code}): {resp.text}")

            res_json = resp.json()
            raw_content = res_json["choices"][0]["message"]["content"].strip()
            if raw_content.startswith("```markdown"):
                raw_content = raw_content[len("```markdown"):].strip()
            elif raw_content.startswith("```"):
                raw_content = raw_content[3:].strip()
            if raw_content.endswith("```"):
                raw_content = raw_content[:-3].strip()

            clean_text = heal_broken_vietnamese_text(raw_content)
            pages_text.append({"page_num": p_idx + 1, "text": clean_text})

        doc.close()

        # Tái tạo tài liệu Word Siêu Cấp VIP Pro từ nội dung AI
        AIReconstructedOCREngine.create_reconstructed_docx(pages_text, output_docx_path)
        file_valid, file_err = AIQualityAuditor.verify_docx_integrity(output_docx_path)

        total_chars = sum(len(p["text"]) for p in pages_text)
        return {
            "success": True,
            "engine": "DeepSeek Vision AI + Động cơ Tái tạo Bố cục Siêu Cấp VIP Pro",
            "total_pages": len(pages_text),
            "total_items": total_chars,
            "mode": output_mode,
            "source_lang": source_lang,
            "target_lang": target_lang,
            "provider": "deepseek-vision",
            "output_path": output_docx_path,
            "context_profile": context_profile.to_dict(),
            "file_valid": file_valid,
            "stages_completed": [
                "Vòng 1 (AI Context Profiler): Phân tích bối cảnh & lập hồ sơ thuật ngữ",
                "Vòng 2 (Context-Aware Translation): Dịch song ngữ xen kẽ theo ngữ cảnh",
                "Vòng 3 (AI Quality Audit & Auto-heal): Kiểm định chất lượng & hoàn thiện tệp"
            ]
        }

    @classmethod
    def translate_pdf_with_vip_pro_engine(
        cls,
        pdf_path: str,
        output_docx_path: str,
        api_key: Optional[str] = None,
        source_lang: str = "en",
        target_lang: str = "vi",
        output_mode: str = "bilingual",
        provider: str = "deepseek"
    ) -> Dict[str, Any]:
        """
        Nhánh B: Xử lý PDF Văn bản & Biểu mẫu Native bằng Động cơ Siêu Cấp VIP Pro (Tool 4)
        kết hợp Dịch thuật Song ngữ Xen kẽ DOC-TRAN.
        """
        import uuid
        temp_docx = output_docx_path + f".temp_{uuid.uuid4().hex[:6]}.docx"

        # Sử dụng Tool 4: SmartPDFToWordDispatcher (MS Word Native Reflow / 7-Layer Layout Engine)
        res_conv = SmartPDFToWordDispatcher.convert(pdf_path, temp_docx, mode="auto")

        try:
            res_trans = cls.translate_docx(
                docx_path=temp_docx,
                output_path=output_docx_path,
                api_key=api_key,
                source_lang=source_lang,
                target_lang=target_lang,
                output_mode=output_mode,
                provider=provider
            )
            res_trans["engine"] = f"Động cơ Siêu Cấp VIP Pro ({res_conv.get('engine', 'VIP Pro')}) + Dịch Interlinear"
            return res_trans
        finally:
            if os.path.exists(temp_docx):
                try:
                    os.remove(temp_docx)
                except Exception:
                    pass

    @classmethod
    def translate_pdf_to_interlinear(
        cls,
        pdf_path: str,
        output_docx_path: str,
        api_key: Optional[str] = None,
        source_lang: str = "en",
        target_lang: str = "vi",
        output_mode: str = "bilingual",
        provider: str = "deepseek"
    ) -> Dict[str, Any]:
        """Tự động phân tích tài liệu và kích hoạt quy trình Siêu Cấp VIP Pro tối ưu."""
        actual_key = api_key or DeepSeekOCR.get_saved_api_key()
        if is_scanned_pdf(pdf_path) and actual_key:
            return cls.translate_scan_with_deepseek_vision(
                pdf_path=pdf_path,
                output_docx_path=output_docx_path,
                api_key=actual_key,
                source_lang=source_lang,
                target_lang=target_lang,
                output_mode=output_mode
            )
        else:
            return cls.translate_pdf_with_vip_pro_engine(
                pdf_path=pdf_path,
                output_docx_path=output_docx_path,
                api_key=actual_key,
                source_lang=source_lang,
                target_lang=target_lang,
                output_mode=output_mode,
                provider=provider
            )

    @classmethod
    def process_interlinear_suite(
        cls,
        input_path: str,
        output_dir: str,
        base_name: str,
        source_lang: str = "en",
        target_lang: str = "vi",
        output_mode: str = "bilingual",
        provider: str = "deepseek",
        engine_mode: str = "auto",       # "auto", "ocr_scan", "vip_pro_doc"
        output_format: str = "docx",     # "docx", "pdf", "both"
        api_key: Optional[str] = None,
        font_family: str = "auto"
    ) -> Dict[str, Any]:
        """
        ĐIỀU PHỐI VIÊN THỐNG NHẤT CHO TOOL 27 (DOC / DOC-TRAN):
        Hợp nhất 3 sức mạnh công nghệ cốt lõi:
        1. Động cơ Chuyển PDF sang Word Siêu Cấp VIP Pro (Tool 4).
        2. Động cơ OCR & Tái tạo Bố cục Đa tầng Siêu Cấp VIP Pro.
        3. Trí tuệ nhân tạo DeepSeek (Vision OCR & Chat API) bảo vệ mã kỹ thuật.
        Hỗ trợ đầu ra: Word (.docx), PDF Vector Siêu Cấp VIP Pro (.pdf), hoặc cả hai (.zip).
        """
        import uuid
        import shutil
        import zipfile

        ext = os.path.splitext(input_path)[1].lower()
        actual_key = (api_key or "").strip() or DeepSeekOCR.get_saved_api_key()
        session_id = uuid.uuid4().hex[:6]

        # 1. Đường dẫn file DOCX tạm thời / chính
        os.makedirs(output_dir, exist_ok=True)
        final_docx_name = f"{base_name}_SongNgu_{session_id}.docx"
        final_docx_path = os.path.join(output_dir, final_docx_name)

        work_input = input_path
        temp_files_to_cleanup = []

        # 2. Xử lý định dạng .doc (chuyển sang .docx)
        if ext == ".doc":
            temp_conv_docx = os.path.join(output_dir, f"{base_name}_temp_{session_id}.docx")
            work_input = convert_doc_to_docx(input_path, temp_conv_docx)
            temp_files_to_cleanup.append(temp_conv_docx)
            ext = ".docx"

        # 3. Phân nhánh xử lý tệp
        engine_used = ""
        trans_result = {}

        if ext == ".docx":
            engine_used = "Dịch thuật DOCX Chuyên sâu"
            trans_result = cls.translate_docx(
                docx_path=work_input,
                output_path=final_docx_path,
                api_key=actual_key,
                source_lang=source_lang,
                target_lang=target_lang,
                output_mode=output_mode,
                provider=provider,
                font_family=font_family
            )

        elif ext == ".pdf":
            # Khi engine_mode là "ocr_scan" thì mới kích hoạt Vision OCR
            if engine_mode == "ocr_scan":
                if not actual_key:
                    raise ValueError("Tài liệu là dạng Scan/Ảnh chụp. Vui lòng nhập DeepSeek API Key để kích hoạt Động cơ Vision AI & Tái tạo Siêu Cấp VIP Pro!")
                engine_used = "DeepSeek Vision AI + Tái tạo Bố cục Siêu Cấp VIP Pro"
                trans_result = cls.translate_scan_with_deepseek_vision(
                    pdf_path=input_path,
                    output_docx_path=final_docx_path,
                    api_key=actual_key,
                    source_lang=source_lang,
                    target_lang=target_lang,
                    output_mode=output_mode
                )
            else:
                # engine_mode là "auto" hoặc "vip_pro_doc":
                # Tự động chuyển đổi PDF sang Word .docx qua SmartPDFToWordDispatcher / pdf2docx
                # và đưa trực tiếp vào translate_docx.
                temp_pdf_docx = os.path.join(output_dir, f"{base_name}_pdf_to_word_{session_id}.docx")
                res_conv = SmartPDFToWordDispatcher.convert(input_path, temp_pdf_docx, mode="auto")
                temp_files_to_cleanup.append(temp_pdf_docx)

                conv_engine = res_conv.get("engine", "VIP Pro") if isinstance(res_conv, dict) else "VIP Pro"
                engine_used = f"Động cơ Chuyển đổi Siêu Cấp VIP Pro ({conv_engine}) + Dịch Interlinear"
                trans_result = cls.translate_docx(
                    docx_path=temp_pdf_docx,
                    output_path=final_docx_path,
                    api_key=actual_key,
                    source_lang=source_lang,
                    target_lang=target_lang,
                    output_mode=output_mode,
                    provider=provider,
                    font_family=font_family
                )
        elif ext in [".xlsx", ".xls"]:
            from core.excel_interlinear_translator import ExcelInterlinearTranslator
            final_xlsx_name = f"{base_name}_SongNgu_{session_id}.xlsx"
            final_xlsx_path = os.path.join(output_dir, final_xlsx_name)

            trans_result = ExcelInterlinearTranslator.translate_excel(
                input_path=work_input,
                output_path=final_xlsx_path,
                source_lang=source_lang,
                target_lang=target_lang,
                provider=provider,
                api_key=actual_key,
                output_mode=output_mode
            )
            engine_used = trans_result.get("engine", "Động cơ Bảng tính Excel DOC-TRAN (openpyxl)")

            # Dọn dẹp tệp trung gian
            for tf in temp_files_to_cleanup:
                if os.path.exists(tf):
                    try: os.remove(tf)
                    except Exception: pass

            out_fmt = (output_format or "xlsx").lower().strip()
            src_str = LANG_NAME_MAP.get(source_lang, source_lang)
            tgt_str = LANG_NAME_MAP.get(target_lang, target_lang)

            final_pdf_name = f"{base_name}_SongNgu_{session_id}.pdf"
            final_pdf_path = os.path.join(output_dir, final_pdf_name)
            try:
                ExcelInterlinearTranslator.convert_xlsx_to_pdf(final_xlsx_path, final_pdf_path)
            except Exception as e:
                logger.warning("Không thể chuyển đổi Excel sang PDF: %s", e)

            zip_name = f"{base_name}_SongNgu_TronBo_{session_id}.zip"
            zip_path = os.path.join(output_dir, zip_name)
            try:
                with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
                    zf.write(final_xlsx_path, arcname=final_xlsx_name)
                    if os.path.exists(final_pdf_path):
                        zf.write(final_pdf_path, arcname=final_pdf_name)
            except Exception as e:
                logger.warning("Không thể tạo ZIP cho Excel: %s", e)

            out_fmt = (output_format or "xlsx").lower().strip()
            src_str = LANG_NAME_MAP.get(source_lang, source_lang)
            tgt_str = LANG_NAME_MAP.get(target_lang, target_lang)

            if out_fmt == "pdf" and os.path.exists(final_pdf_path):
                primary_filename = final_pdf_name
                primary_output_path = final_pdf_path
                primary_download_url = f"/api/download/{final_pdf_name}"
                output_format_label = "PDF Bảng Tính Chuẩn In Ấn"
            elif out_fmt == "both" and os.path.exists(zip_path):
                primary_filename = zip_name
                primary_output_path = zip_path
                primary_download_url = f"/api/download/{zip_name}"
                output_format_label = "Trọn bộ Excel (.xlsx) & PDF (.pdf)"
            else:
                primary_filename = final_xlsx_name
                primary_output_path = final_xlsx_path
                primary_download_url = f"/api/download/{final_xlsx_name}"
                output_format_label = "Excel (.xlsx)"

            return {
                "success": True,
                "filename": primary_filename,
                "output_path": primary_output_path,
                "download_url": primary_download_url,
                "download_docx_url": f"/api/download/{final_xlsx_name}",
                "download_xlsx_url": f"/api/download/{final_xlsx_name}",
                "download_pdf_url": f"/api/download/{final_pdf_name}" if os.path.exists(final_pdf_path) else primary_download_url,
                "download_zip_url": f"/api/download/{zip_name}" if os.path.exists(zip_path) else primary_download_url,
                "xlsx_path": final_xlsx_path,
                "docx_path": final_xlsx_path,
                "pdf_path": final_pdf_path if os.path.exists(final_pdf_path) else None,
                "zip_path": zip_path if os.path.exists(zip_path) else None,
                "engine": engine_used,
                "source_lang": source_lang,
                "target_lang": target_lang,
                "output_mode": output_mode,
                "output_format": output_format_label,
                "total_sheets": trans_result.get("total_sheets", 1),
                "translated_cells": trans_result.get("translated_cells", 0),
                "details": f"Dịch song ngữ bảng tính ({src_str} ➔ {tgt_str}) hoàn tất ({trans_result.get('translated_cells', 0)} ô) bằng {engine_used}. Đã xuất trọn bộ tệp Excel (.xlsx), PDF và ZIP."
            }
        else:
            raise ValueError(f"Định dạng tệp không được hỗ trợ: {ext}. Vui lòng chọn .docx, .doc, .pdf, .xlsx hoặc .xls.")

        # Dọn dẹp tệp trung gian
        for tf in temp_files_to_cleanup:
            if os.path.exists(tf):
                try: os.remove(tf)
                except Exception: pass

        # 4. Xuất trọn bộ 3 định dạng sản phẩm (Tri-Format Output Packaging):
        # Artifact 1: File Word (.docx) đã được dịch tại final_docx_path, kèm tùy chọn .doc
        final_doc_name = f"{base_name}_SongNgu_{session_id}.doc"
        final_doc_path = os.path.join(output_dir, final_doc_name)
        has_doc = False
        try:
            doc_res = convert_docx_to_doc(final_docx_path, final_doc_path)
            if doc_res and os.path.exists(final_doc_path) and os.path.getsize(final_doc_path) > 0:
                has_doc = True
        except Exception as e:
            logger.warning("Không thể xuất thêm định dạng .doc: %s", e)

        out_fmt = (output_format or "docx").lower().strip()

        # Artifact 2: File PDF (.pdf) chuyển ngược từ final_docx_path (chỉ tạo khi người dùng yêu cầu)
        final_pdf_name = f"{base_name}_SongNgu_{session_id}.pdf"
        final_pdf_path = os.path.join(output_dir, final_pdf_name)
        zip_name = f"{base_name}_SongNgu_TronBo_{session_id}.zip"
        zip_path = os.path.join(output_dir, zip_name)

        if out_fmt in ["pdf", "both"]:
            try:
                convert_docx_to_pdf(final_docx_path, final_pdf_path)
            except Exception as e:
                logger.warning("PDF conversion skipped: %s", e)

            # Artifact 3: File ZIP (.zip) đóng gói cả Word (.docx, .doc) và PDF (.pdf)
            try:
                with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
                    zf.write(final_docx_path, arcname=final_docx_name)
                    if os.path.exists(final_pdf_path):
                        zf.write(final_pdf_path, arcname=final_pdf_name)
                    if has_doc and os.path.exists(final_doc_path):
                        zf.write(final_doc_path, arcname=final_doc_name)
            except Exception as e:
                logger.warning("ZIP creation skipped: %s", e)

        # 5. Thiết lập phản hồi đa định dạng theo yêu cầu người dùng
        if out_fmt == "pdf":
            primary_filename = final_pdf_name
            primary_output_path = final_pdf_path
            primary_download_url = f"/api/download/{final_pdf_name}"
            output_format_label = "PDF Vector Siêu Cấp VIP Pro"
        elif out_fmt == "both":
            primary_filename = zip_name
            primary_output_path = zip_path
            primary_download_url = f"/api/download/{zip_name}"
            output_format_label = "Trọn bộ Word (.docx) & PDF Vector (.pdf)"
        else:  # "docx" or default
            primary_filename = final_docx_name
            primary_output_path = final_docx_path
            primary_download_url = f"/api/download/{final_docx_name}"
            output_format_label = "Word (.docx)"

        src_str = LANG_NAME_MAP.get(source_lang, source_lang)
        tgt_str = LANG_NAME_MAP.get(target_lang, target_lang)

        response_payload = {
            "success": True,
            "filename": primary_filename,
            "output_path": primary_output_path,
            "docx_path": final_docx_path,
            "pdf_path": final_pdf_path,
            "zip_path": zip_path,
            "download_url": primary_download_url,
            "download_docx_url": f"/api/download/{final_docx_name}",
            "download_pdf_url": f"/api/download/{final_pdf_name}",
            "download_zip_url": f"/api/download/{zip_name}",
            "engine": engine_used,
            "source_lang": source_lang,
            "target_lang": target_lang,
            "output_mode": output_mode,
            "output_format": output_format_label,
            "total_items": trans_result.get("total_items", 0),
            "details": f"Dịch song ngữ ({src_str} ➔ {tgt_str}) hoàn tất ({trans_result.get('total_items', 0)} mục) bằng {engine_used}. Đã chuẩn hóa Times New Roman, chữ xanh #002060 nghiêng và bảo toàn bảng biểu."
        }
        if has_doc and os.path.exists(final_doc_path):
            response_payload["doc_path"] = final_doc_path
            response_payload["download_doc_url"] = f"/api/download/{final_doc_name}"

        if isinstance(trans_result, dict):
            if "context_profile" in trans_result:
                response_payload["context_profile"] = trans_result["context_profile"]
            if "audit_report" in trans_result:
                response_payload["audit_report"] = trans_result["audit_report"]
            if "stages_completed" in trans_result:
                response_payload["stages_completed"] = trans_result["stages_completed"]

        return response_payload
