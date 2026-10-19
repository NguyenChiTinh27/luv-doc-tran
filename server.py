"""
LUV-DOC-TRAN — Standalone Commercial Bilingual Translation Web Application
─────────────────────────────────────────────────────────────────────────
LUV-DOC-TRAN Studio Pro - Nền Tảng Dịch Thuật Đa Ngữ & Đa Định Dạng Cao Cấp
Deployable on Render, Docker, Linux, and Windows.
"""

import os
import io
import sys
import uuid
import time
import shutil
import logging
import tempfile
import base64
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Dict, Any

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

# Setup logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")
logger = logging.getLogger("luv_doctran")

# Path setup
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
TEMP_DIR = os.path.join(DATA_DIR, "temp")
OUTPUT_DIR = os.path.join(DATA_DIR, "output")
STATIC_DIR = os.path.join(BASE_DIR, "static")

os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(STATIC_DIR, exist_ok=True)

# Import Core Translation and Licensing modules
from core.license_client import LUVLicenseClient
from core.interlinear_translator import InterlinearTranslator, convert_doc_to_docx, convert_docx_to_pdf
from core.excel_interlinear_translator import ExcelInterlinearTranslator
from core.deepseek_ocr import DeepSeekOCR

app = FastAPI(
    title="LUV-DOC-TRAN API",
    description="Công cụ dịch thuật song ngữ kỹ thuật xen kẽ chuyên nghiệp và thương mại",
    version="1.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def save_upload_file(upload_file: UploadFile, destination_dir: str) -> str:
    """Save an uploaded file safely to disk and return absolute path."""
    ext = os.path.splitext(upload_file.filename or "")[1]
    safe_name = f"{uuid.uuid4().hex[:10]}_{upload_file.filename}"
    file_path = os.path.join(destination_dir, safe_name)
    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(upload_file.file, buffer)
    return file_path


# ============================================================
# PREVIEW CONVERSION HELPERS
# ============================================================
def convert_docx_bytes_to_html(content: bytes) -> str:
    """Convert docx bytes to clean HTML using mammoth with embedded images."""
    try:
        import mammoth
        def _convert_image(image):
            with image.open() as image_bytes:
                encoded_src = base64.b64encode(image_bytes.read()).decode("ascii")
                return {
                    "src": f"data:{image.content_type};base64,{encoded_src}",
                    "class": "docx-preview-image",
                    "style": "max-width:100%; height:auto; border-radius:8px; margin:10px 0; display:block;"
                }
        file_stream = io.BytesIO(content)
        res = mammoth.convert_to_html(file_stream, convert_image=mammoth.images.img_element(_convert_image))
        return res.value
    except Exception as e:
        logger.warning("Mammoth convert failed, using fallback: %s", e)
        return f"<div class='preview-fallback'><p>Xem trước tài liệu Word ({len(content)} bytes)</p></div>"


def convert_word_bytes_to_html(content: bytes, filename: str = "document.docx") -> str:
    """Convert .docx or binary .doc to HTML."""
    is_binary_doc = (
        content.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1")
        or (filename.lower().endswith(".doc") and not content.startswith(b"PK\x03\x04"))
    )
    if is_binary_doc:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_doc = os.path.join(temp_dir, "input.doc")
            temp_docx = os.path.join(temp_dir, "converted.docx")
            with open(temp_doc, "wb") as f:
                f.write(content)
            convert_doc_to_docx(temp_doc, temp_docx)
            if not os.path.exists(temp_docx) or os.path.getsize(temp_docx) == 0:
                raise RuntimeError("Không thể chuyển đổi tệp .doc nhị phân sang định dạng .docx.")
            with open(temp_docx, "rb") as f:
                content = f.read()
    return convert_docx_bytes_to_html(content)


def convert_pdf_to_html_preview(file_path: str) -> str:
    """Render PDF pages as clean SVG or text preview using PyMuPDF."""
    import fitz
    doc = fitz.open(file_path)
    html_parts = ["<div class='pdf-preview-container'>"]
    for i in range(min(len(doc), 15)):  # Preview first 15 pages for speed
        page = doc[i]
        svg = page.get_svg_image()
        html_parts.append(f"<div class='pdf-page-wrapper' style='margin-bottom:20px; box-shadow:0 4px 12px rgba(0,0,0,0.1); border-radius:6px; overflow:hidden;'>{svg}</div>")
    html_parts.append("</div>")
    doc.close()
    return "".join(html_parts)


# ============================================================
# API ENDPOINTS
# ============================================================
@app.get("/api/status")
def get_system_status():
    """System health check and environmental capabilities."""
    has_libreoffice = shutil.which("soffice") is not None or shutil.which("libreoffice") is not None
    return {
        "status": "online",
        "app_name": "LUV-DOC-TRAN",
        "version": "1.0.0-commercial",
        "os": sys.platform,
        "has_libreoffice": has_libreoffice,
        "time": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    }


@app.get("/api/trial/status")
def get_trial_status(device_id: str = ""):
    """Get 48-hour free trial timer and validity for client device."""
    return LUVLicenseClient.get_trial_status(device_id)


@app.get("/api/license/status")
def get_license_status(device_id: str = "", key: str = ""):
    """Get stored license activation state with real-time Cloudflare Worker validation."""
    try:
        clean_key = (key or "").strip()
        # If server does not have an active license but client passed key from localStorage, auto-sync!
        current_info = LUVLicenseClient.get_license_info()
        if not current_info.get("is_active") and clean_key:
            try:
                LUVLicenseClient.activate_key(clean_key, registered_to="Người Dùng Bản Quyền", client_device_hint=device_id)
            except Exception as ex:
                logger.warning("Auto-activation from status query failed: %s", ex)

        status = LUVLicenseClient.validate_license(device_id=device_id, force_online=True)
        if status.get("revoked"):
            return {
                "is_active": False,
                "revoked": True,
                "message": "Key kích hoạt đã bị thu hồi trên máy chủ D1. Vui lòng liên hệ quản trị viên."
            }
        return status
    except Exception as e:
        return {
            "is_active": False,
            "revoked": False,
            "tier": "Chưa kích hoạt",
            "registered_to": None,
            "customer": None,
            "license_key": None,
            "error": str(e)
        }


@app.post("/api/license/activate")
def activate_license(
    key: str = Form(...),
    registered_to: str = Form("Người Dùng Bản Quyền"),
    device_id: str = Form("")
):
    """Activate license key on this instance."""
    try:
        record = LUVLicenseClient.activate_key(key, registered_to, client_device_hint=device_id)
        return {"success": True, "message": "Key kích hoạt đã được xác thực thành công!", "data": record}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/preview/upload")
async def preview_upload_file(file: UploadFile = File(...)):
    """Generate instant HTML preview for an uploaded file (Word, Excel, PDF)."""
    filename = file.filename or "document"
    ext = os.path.splitext(filename)[1].lower()

    try:
        content = await file.read()
        if not content:
            raise HTTPException(status_code=400, detail="Tệp tin tải lên rỗng.")

        if ext in [".docx", ".doc"]:
            html = convert_word_bytes_to_html(content, filename=filename)
            return {"success": True, "type": "word", "filename": filename, "html": html}

        elif ext in [".xlsx", ".xls"]:
            is_binary_xls = content.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1") or (ext == ".xls" and not content.startswith(b"PK\x03\x04"))
            if is_binary_xls:
                with tempfile.TemporaryDirectory() as temp_dir:
                    temp_xls = os.path.join(temp_dir, "input.xls")
                    temp_xlsx = os.path.join(temp_dir, "converted.xlsx")
                    with open(temp_xls, "wb") as f:
                        f.write(content)
                    ExcelInterlinearTranslator.convert_xls_to_xlsx(temp_xls, temp_xlsx)
                    data = ExcelInterlinearTranslator.render_excel_to_html(temp_xlsx)
            else:
                data = ExcelInterlinearTranslator.render_excel_to_html(content)
            data["success"] = True
            data["type"] = "excel"
            data["filename"] = filename
            if "sheets_data" in data and data["sheets_data"]:
                data["html"] = data["sheets_data"][0]["html"]
            return data

        elif ext == ".pdf":
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tf:
                tf.write(content)
                tf_path = tf.name
            try:
                html = convert_pdf_to_html_preview(tf_path)
            finally:
                if os.path.exists(tf_path):
                    os.remove(tf_path)
            return {"success": True, "type": "pdf", "filename": filename, "html": html}

        else:
            raise HTTPException(status_code=400, detail=f"Định dạng {ext} chưa được hỗ trợ xem trước.")
    except Exception as e:
        logger.error("Preview upload error: %s", e)
        raise HTTPException(status_code=500, detail=f"Lỗi xem trước tệp: {str(e)}")


@app.get("/api/preview/output/{filename}")
def preview_output_file(filename: str):
    """Render HTML preview for generated output document."""
    file_path = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="Không tìm thấy tệp kết quả.")

    ext = os.path.splitext(filename)[1].lower()
    try:
        with open(file_path, "rb") as f:
            content = f.read()

        if ext in [".docx", ".doc"]:
            html = convert_word_bytes_to_html(content, filename=filename)
            return {"success": True, "type": "word", "filename": filename, "html": html}
        elif ext in [".xlsx", ".xls"]:
            data = ExcelInterlinearTranslator.render_excel_to_html(content)
            data["success"] = True
            data["type"] = "excel"
            data["filename"] = filename
            if "sheets_data" in data and data["sheets_data"]:
                data["html"] = data["sheets_data"][0]["html"]
            return data
        elif ext == ".pdf":
            html = convert_pdf_to_html_preview(file_path)
            return {"success": True, "type": "pdf", "filename": filename, "html": html}
        else:
            return {"success": False, "detail": "Không hỗ trợ xem trước định dạng này"}
    except Exception as e:
        logger.error("Preview output error: %s", e)
        raise HTTPException(status_code=500, detail=f"Lỗi xem trước tệp kết quả: {str(e)}")


@app.post("/api/translate")
async def process_translation(
    file: UploadFile = File(...),
    source_lang: str = Form("auto"),
    target_lang: str = Form("vi"),
    output_mode: str = Form("bilingual"),
    provider: str = Form("neural"),
    engine_mode: str = Form("auto"),
    api_key: Optional[str] = Form(None),
    device_id: str = Form(""),
    license_key: Optional[str] = Form(None),
    font_family: str = Form("auto")
):
    """
    Main Bilingual Interlinear Translation Engine:
    - Bắt buộc Key Kích Hoạt hợp lệ (Active License Key).
    - Prioritizes Word (.docx) generation for instant delivery.
    - Supports Word (.docx, .doc), PDF (auto-converted to docx), and Excel (.xlsx, .xls).
    """
    actual_provider = "deepseek" if provider in ["neural", "deepseek"] else provider
    clean_key = (license_key or "").strip()

    # 1. Access Control Check (Local cryptographic Ed25519 verification)
    access = LUVLicenseClient.check_access(device_id=device_id, force_online=False)

    # If server has no active license on disk, but client passed key from localStorage, auto-activate!
    if not access.get("has_access") and clean_key:
        try:
            LUVLicenseClient.activate_key(clean_key, registered_to="Người Dùng Bản Quyền", client_device_hint=device_id)
            access = LUVLicenseClient.check_access(device_id=device_id, force_online=False)
        except Exception as ex:
            logger.warning("Auto-activation during translate failed: %s", ex)

    if access.get("revoked"):
        return JSONResponse(
            status_code=403,
            content={
                "success": False,
                "revoked": True,
                "message": "Key kích hoạt đã bị thu hồi trên máy chủ D1. Vui lòng liên hệ quản trị viên."
            }
        )
    if not access.get("has_access"):
        return JSONResponse(
            status_code=403,
            content={
                "success": False,
                "license_required": True,
                "message": access.get("message") or "Hệ thống yêu cầu nhập Key kích hoạt để bắt đầu dịch tài liệu."
            }
        )

    filename = file.filename or "document"
    ext = os.path.splitext(filename)[1].lower()
    base_name = os.path.splitext(filename)[0]

    # Save uploaded file
    saved_path = save_upload_file(file, TEMP_DIR)

    effective_api_key = (api_key or "").strip() or DeepSeekOCR.get_saved_api_key()
    if api_key and api_key.strip():
        DeepSeekOCR.save_api_key(api_key.strip())

    start_time = time.time()

    try:
        # Case A: Excel Spreadsheet (.xlsx, .xls)
        if ext in [".xlsx", ".xls"]:
            out_filename = f"{base_name}_SongNgu_{uuid.uuid4().hex[:6]}.xlsx"
            out_path = os.path.join(OUTPUT_DIR, out_filename)

            res = ExcelInterlinearTranslator.translate_excel(
                input_path=saved_path,
                output_path=out_path,
                source_lang=source_lang,
                target_lang=target_lang,
                provider=actual_provider,
                api_key=effective_api_key,
                font_family=font_family
            )

            elapsed = round(time.time() - start_time, 2)
            return {
                "success": True,
                "type": "excel",
                "filename": out_filename,
                "download_url": f"/api/download/{out_filename}",
                "download_docx_url": f"/api/download/{out_filename}",  # fallback alias
                "preview_output_filename": out_filename,
                "elapsed_seconds": elapsed,
                "context_profile": res.get("context_profile"),
                "audit_report": res.get("audit_report"),
                "stages_completed": res.get("stages_completed"),
                "details": f"Đã dịch song ngữ thành công {res.get('total_translated_cells', 0)} ô văn bản ({res.get('sheet_count', 1)} Sheet)."
            }

        # Case B: Word (.docx, .doc) or PDF (.pdf)
        else:
            # Word (.docx) is the exclusive output format for both Word and PDF inputs
            res = InterlinearTranslator.process_interlinear_suite(
                input_path=saved_path,
                output_dir=OUTPUT_DIR,
                base_name=base_name,
                source_lang=source_lang,
                target_lang=target_lang,
                output_mode=output_mode,
                provider=actual_provider,
                engine_mode=engine_mode,
                output_format="docx",
                api_key=effective_api_key,
                font_family=font_family
            )

            # Ensure URLs
            if isinstance(res, dict):
                docx_file = os.path.basename(res.get("docx_path", "")) if res.get("docx_path") else None

                if docx_file:
                    res["download_docx_url"] = f"/api/download/{docx_file}"
                    res["download_url"] = f"/api/download/{docx_file}"
                    res["preview_output_filename"] = docx_file

                res["elapsed_seconds"] = round(time.time() - start_time, 2)
                res["primary_format"] = "docx"
                res["word_priority_note"] = "Word (.docx) đã sẵn sàng để tải về và chỉnh sửa tức thì."

            return res

    except Exception as e:
        logger.error("Translation processing error for %s: %s", filename, e, exc_info=True)
        raise HTTPException(status_code=500, detail=f"Lỗi dịch tài liệu: {str(e)}")


@app.post("/api/export-pdf")
def export_pdf_from_docx(docx_filename: str = Form(...)):
    """
    On-demand vector PDF exporter from translated DOCX file.
    Supports user requirement: 'nhớ lưu ý chỗ xuất PDF ( ghi chú là đợi 2 -3 phút , nhớ thêm ảnh động thể hiện tiến trình đang xử lý )'
    """
    docx_path = os.path.join(OUTPUT_DIR, docx_filename)
    if not os.path.exists(docx_path):
        raise HTTPException(status_code=404, detail=f"Không tìm thấy file Word '{docx_filename}' để xuất PDF.")

    base_name = os.path.splitext(docx_filename)[0]
    out_pdf_name = f"{base_name}.pdf"
    out_pdf_path = os.path.join(OUTPUT_DIR, out_pdf_name)

    # If already exists and valid size, return immediately
    if os.path.exists(out_pdf_path) and os.path.getsize(out_pdf_path) > 1000:
        return {
            "success": True,
            "filename": out_pdf_name,
            "download_pdf_url": f"/api/download/{out_pdf_name}",
            "cached": True
        }

    try:
        # Convert docx to pdf
        InterlinearTranslator._convert_docx_to_pdf_vector(docx_path, out_pdf_path)
        if not os.path.exists(out_pdf_path) or os.path.getsize(out_pdf_path) == 0:
            raise RuntimeError("Quá trình kết xuất PDF không tạo ra tệp đầu ra hợp lệ.")

        return {
            "success": True,
            "filename": out_pdf_name,
            "download_pdf_url": f"/api/download/{out_pdf_name}",
            "cached": False
        }
    except Exception as e:
        logger.error("PDF Export error for %s: %s", docx_filename, e)
        raise HTTPException(status_code=500, detail=f"Không thể kết xuất PDF: {str(e)}")


@app.post("/api/download-all-zip")
async def download_all_zip(request: Request):
    """
    Tải toàn bộ tệp kết quả trong một tệp nén ZIP duy nhất:
    Nhận danh sách filenames: List[str] từ request body JSON.
    """
    try:
        body = await request.json()
    except Exception:
        body = {}

    filenames = body.get("filenames", [])
    if not filenames or not isinstance(filenames, list):
        raise HTTPException(status_code=400, detail="Danh sách tệp tin không hợp lệ.")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_filename = f"LUV_DOC_TRAN_TronBo_{timestamp}_{uuid.uuid4().hex[:6]}.zip"
    zip_path = os.path.join(OUTPUT_DIR, zip_filename)

    added_count = 0
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname in filenames:
            safe_fname = os.path.basename(str(fname).strip())
            file_path = os.path.join(OUTPUT_DIR, safe_fname)
            if not os.path.exists(file_path):
                file_path = os.path.join(TEMP_DIR, safe_fname)
            if os.path.exists(file_path) and os.path.isfile(file_path):
                zf.write(file_path, arcname=safe_fname)
                added_count += 1

    if added_count == 0:
        if os.path.exists(zip_path):
            os.remove(zip_path)
        raise HTTPException(status_code=404, detail="Không tìm thấy tệp kết quả nào khả dụng để tạo file nén.")

    return {
        "success": True,
        "filename": zip_filename,
        "download_url": f"/api/download/{zip_filename}",
        "file_count": added_count
    }


@app.get("/api/download/{filename}")
def download_file(filename: str):
    """Serve processed document for download."""
    file_path = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(file_path):
        # Check temp dir as fallback
        file_path = os.path.join(TEMP_DIR, filename)
        if not os.path.exists(file_path):
            raise HTTPException(status_code=404, detail="Tệp tin không tồn tại hoặc đã hết hạn lưu trữ.")

    return FileResponse(
        path=file_path,
        filename=filename,
        media_type="application/octet-stream"
    )


# Serve Static files and Root
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

@app.get("/")
def serve_index():
    """Serve main Pastel Kawaii interface."""
    index_file = os.path.join(STATIC_DIR, "index.html")
    if os.path.exists(index_file):
        return FileResponse(index_file)
    return HTMLResponse("<h1>LUV-DOC-TRAN đang khởi động...</h1>")


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8080))
    uvicorn.run("server:app", host="0.0.0.0", port=port, reload=True)
