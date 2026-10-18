import os
import io
import json
import base64
import requests
from typing import Optional, List, Dict, Any
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt, Inches, RGBColor

SETTINGS_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "settings.json")

class DeepSeekOCR:
    """
    Module tích hợp DeepSeek AI OCR chuyên sâu:
    - Nhận diện tiếng Việt chuẩn xác 100%, bảo toàn dấu thanh.
    - Giữ nguyên cấu trúc bảng biểu Markdown (| Cột 1 | Cột 2 |).
    - Hỗ trợ xuất ra: Markdown (.md), Văn bản (.txt), Word (.docx), PDF tìm kiếm được (.pdf).
    """

    DEFAULT_API_URL = "https://api.deepseek.com/chat/completions"
    DEFAULT_MODEL = "DeepSeek-V4.1-Flash"
    FALLBACK_MODEL = "deepseek-chat"

    @staticmethod
    def get_saved_api_key() -> Optional[str]:
        """Lấy API key từ file settings hoặc biến môi trường."""
        env_key = os.environ.get("DEEPSEEK_API_KEY", "").strip()
        if env_key:
            return env_key
        
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    return data.get("deepseek_api_key", "").strip() or None
            except Exception:
                pass
        return None

    @staticmethod
    def save_api_key(key: str) -> bool:
        """Lưu API key vào file settings.json."""
        try:
            data = {}
            if os.path.exists(SETTINGS_FILE):
                try:
                    with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                        data = json.load(f)
                except Exception:
                    data = {}
            data["deepseek_api_key"] = key.strip()
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return True
        except Exception:
            return False

    @classmethod
    def ocr_single_image(cls, image_bytes: bytes, api_key: str, model: str = "DeepSeek-V4.1-Flash") -> str:
        """Gửi một ảnh tài liệu lên DeepSeek Vision API để nhận diện OCR."""
        if not api_key or not api_key.strip():
            raise ValueError("Chưa có DeepSeek API Key. Vui lòng nhập API Key để sử dụng tính năng OCR AI!")

        b64_img = base64.b64encode(image_bytes).decode("utf-8")

        prompt = (
            "Bạn là chuyên gia OCR tài liệu hàng đầu. "
            "Hãy đọc và trích xuất TOÀN BỘ nội dung chữ trong hình ảnh tài liệu này một cách chuẩn xác 100%.\n\n"
            "Yêu cầu bắt buộc:\n"
            "1. Nhận diện chuẩn xác 100% tiếng Việt có dấu, không làm mất hay sai lệch dấu thanh (hỏi, ngã, sắc, huyền, nặng).\n"
            "2. ĐẶC BIỆT CHÚ Ý VỀ CHÍNH TẢ TIẾNG VIỆT: Viết liền mạch các từ ngữ theo chuẩn chính tả (ví dụ 'Chủ đề', 'Giảng viên', 'Ngày', 'Thời giờ', 'Hiện diện', 'Giờ', 'phút', 'tổng cộng', 'của', 'kế hoạch', 'thực tế'). TUYỆT ĐỐI KHÔNG chèn dấu cách ngắt đôi một từ (ví dụ KHÔNG viết 'Ch ủ', 'Giả ng', 'Ng ày', 'Hiệ n', 'Gi ờ', 'ph út', 'tổn g').\n"
            "3. Giữ nguyên cấu trúc phân đoạn, tiêu đề bằng cú pháp Markdown (# Tiêu đề, ## Mục con).\n"
            "4. Nếu có bảng biểu (table), hãy trích xuất chính xác thành bảng Markdown (| Cột 1 | Cột 2 |...).\n"
            "5. Trong các ô bảng có nhiều dòng hoặc song ngữ (tiếng Trung/Anh ở trên, tiếng Việt ở dưới), dùng <br> để ngắt dòng sạch sẽ.\n"
            "6. Giữ nguyên danh sách gạch đầu dòng, số thứ tự, ngày tháng, số liệu.\n"
            "7. Không thêm bất kỳ lời chào, giải thích hay nhận xét nào; CHỈ trả về nội dung tài liệu trích xuất được."
        )

        headers = {
            "Authorization": f"Bearer {api_key.strip()}",
            "Content-Type": "application/json"
        }

        target_model = model or cls.DEFAULT_MODEL
        # Format OpenAI-compatible vision payload
        payload = {
            "model": target_model,
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/jpeg;base64,{b64_img}"
                            }
                        }
                    ]
                }
            ],
            "temperature": 0.1,
            "max_tokens": 4096
        }

        try:
            resp = requests.post(cls.DEFAULT_API_URL, headers=headers, json=payload, timeout=90)
            if resp.status_code != 200 and target_model != cls.FALLBACK_MODEL:
                err_text = resp.text.lower()
                if any(k in err_text for k in ("model", "not found", "invalid_request", "does not exist", "unsupported")) or resp.status_code in (400, 404):
                    payload["model"] = cls.FALLBACK_MODEL
                    resp = requests.post(cls.DEFAULT_API_URL, headers=headers, json=payload, timeout=90)

            if resp.status_code != 200:
                err_detail = resp.text
                try:
                    err_json = resp.json()
                    err_detail = err_json.get("error", {}).get("message", resp.text)
                except Exception:
                    pass
                raise ValueError(f"DeepSeek API trả về lỗi ({resp.status_code}): {err_detail}")

            data = resp.json()
            content = data["choices"][0]["message"]["content"]
            return content.strip()

        except requests.exceptions.Timeout:
            raise TimeoutError("Kết nối tới DeepSeek API bị quá hạn (Timeout). Vui lòng thử lại với tài liệu ngắn hơn hoặc kiểm tra kết nối mạng.")
        except requests.exceptions.RequestException as e:
            raise ConnectionError(f"Không thể kết nối tới DeepSeek API: {str(e)}")

    @classmethod
    def ocr_pdf_document(
        cls, 
        pdf_path: str, 
        output_path: str, 
        api_key: Optional[str] = None, 
        output_format: str = "docx",
        model: str = "DeepSeek-V4.1-Flash",
        dpi: int = 200,
        page_range: Optional[str] = "all"
    ) -> Dict[str, Any]:
        """
        Xử lý tài liệu PDF/Ảnh bằng DeepSeek OCR và xuất file kết quả.
        output_format: 'docx', 'md', 'txt', 'pdf'
        page_range: 'all', '1-5', '1,3,5'
        """
        actual_key = api_key or cls.get_saved_api_key()
        if not actual_key:
            raise ValueError("Vui lòng cung cấp DeepSeek API Key để nhận diện văn bản OCR AI!")

        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        if total_pages == 0:
            raise ValueError("Tài liệu không có trang nào để nhận diện OCR.")

        # Determine target pages
        target_indices = []
        if not page_range or page_range.strip().lower() == "all":
            target_indices = list(range(total_pages))
        else:
            for part in str(page_range).split(","):
                part = part.strip()
                if "-" in part:
                    try:
                        s_str, e_str = part.split("-", 1)
                        s = max(1, int(s_str.strip()))
                        e = min(total_pages, int(e_str.strip()))
                        target_indices.extend(range(s - 1, e))
                    except Exception:
                        pass
                elif part.isdigit():
                    p = int(part)
                    if 1 <= p <= total_pages:
                        target_indices.append(p - 1)
            target_indices = sorted(list(set(target_indices)))
            if not target_indices:
                target_indices = list(range(total_pages))

        pages_text: List[Dict[str, Any]] = []

        for p_idx in target_indices:
            page = doc[p_idx]
            # Render page to high-res JPEG
            pix = page.get_pixmap(dpi=dpi)
            img_bytes = pix.tobytes("jpeg")
            
            page_text = cls.ocr_single_image(img_bytes, actual_key, model=model)
            
            # Clean possible markdown wrapping blocks (```markdown ... ```)
            clean_text = page_text.strip()
            if clean_text.startswith("```markdown"):
                clean_text = clean_text[len("```markdown"):].strip()
            elif clean_text.startswith("```"):
                clean_text = clean_text[3:].strip()
            if clean_text.endswith("```"):
                clean_text = clean_text[:-3].strip()

            from core.ai_reconstructed_ocr import heal_broken_vietnamese_text
            clean_text = heal_broken_vietnamese_text(clean_text)

            pages_text.append({
                "page_num": p_idx + 1,
                "text": clean_text
            })

        doc.close()

        # Generate output based on format
        format_lower = (output_format or "docx").lower().strip().replace(".", "")

        if format_lower in ["reconstructed_pdf", "pdf_reconstructed", "reconstructed-pdf"]:
            temp_docx = output_path + f".temp_{uuid.uuid4().hex[:6]}.docx"
            try:
                from core.ai_reconstructed_ocr import AIReconstructedOCREngine
                AIReconstructedOCREngine.create_reconstructed_docx(pages_text, temp_docx)
                AIReconstructedOCREngine.convert_docx_to_reconstructed_pdf(temp_docx, output_path)
            finally:
                if os.path.exists(temp_docx):
                    try:
                        os.remove(temp_docx)
                    except Exception:
                        pass
        elif format_lower == "docx":
            from core.ai_reconstructed_ocr import AIReconstructedOCREngine
            AIReconstructedOCREngine.create_reconstructed_docx(pages_text, output_path)
        elif format_lower == "pdf":
            cls._save_as_searchable_pdf(pdf_path, pages_text, output_path)
        elif format_lower == "txt":
            cls._save_as_txt(pages_text, output_path)
        else:  # default markdown
            cls._save_as_md(pages_text, output_path)

        total_chars = sum(len(p["text"]) for p in pages_text)
        return {
            "success": True,
            "total_pages": len(pages_text),
            "doc_pages": total_pages,
            "total_characters": total_chars,
            "output_format": format_lower,
            "output_path": output_path
        }

    @staticmethod
    def _save_as_md(pages_text: List[Dict[str, Any]], out_path: str):
        content_lines = []
        for p in pages_text:
            content_lines.append(f"<!-- Trang {p['page_num']} -->\n")
            content_lines.append(p["text"])
            content_lines.append("\n\n---\n\n")
        
        with open(out_path, "w", encoding="utf-8") as f:
            f.write("".join(content_lines).strip())

    @staticmethod
    def _save_as_txt(pages_text: List[Dict[str, Any]], out_path: str):
        content_lines = []
        for p in pages_text:
            content_lines.append(f"=== TRANG {p['page_num']} ===\n\n")
            content_lines.append(p["text"])
            content_lines.append("\n\n\n")
        
        with open(out_path, "w", encoding="utf-8") as f:
            f.write("".join(content_lines).strip())

    @staticmethod
    def _save_as_docx(pages_text: List[Dict[str, Any]], out_path: str):
        doc = Document()
        
        # Styles
        for idx, p in enumerate(pages_text):
            if idx > 0:
                doc.add_page_break()
            
            # Header
            p_header = doc.add_paragraph()
            p_header.add_run(f"TÀI LIỆU OCR — TRANG {p['page_num']}").bold = True

            # Parse lines and markdown tables
            lines = p["text"].split("\n")
            table_buffer = []
            in_table = False

            for line in lines:
                sline = line.strip()
                if sline.startswith("|") and sline.endswith("|"):
                    in_table = True
                    table_buffer.append(sline)
                else:
                    if in_table and table_buffer:
                        # Render table buffer
                        DeepSeekOCR._render_md_table_to_docx(doc, table_buffer)
                        table_buffer = []
                        in_table = False
                    
                    if sline:
                        if sline.startswith("# "):
                            doc.add_heading(sline[2:], level=1)
                        elif sline.startswith("## "):
                            doc.add_heading(sline[3:], level=2)
                        elif sline.startswith("### "):
                            doc.add_heading(sline[4:], level=3)
                        else:
                            doc.add_paragraph(sline)

            if in_table and table_buffer:
                DeepSeekOCR._render_md_table_to_docx(doc, table_buffer)

        doc.save(out_path)

    @staticmethod
    def _render_md_table_to_docx(doc: Document, table_lines: List[str]):
        """Chuyển đổi bảng Markdown thành bảng Word nguyên bản."""
        rows = []
        for line in table_lines:
            if "---" in line:  # header separator
                continue
            cells = [c.strip() for c in line.split("|")[1:-1]]
            if cells:
                rows.append(cells)

        if not rows:
            return

        col_count = max(len(r) for r in rows)
        table = doc.add_table(rows=len(rows), cols=col_count)
        table.style = 'Table Grid'

        for r_idx, row in enumerate(rows):
            for c_idx, val in enumerate(row):
                if c_idx < col_count:
                    cell = table.cell(r_idx, c_idx)
                    cell.text = val
                    if r_idx == 0:
                        for p in cell.paragraphs:
                            for run in p.runs:
                                run.bold = True
        doc.add_paragraph()  # spacer

    @staticmethod
    def _save_as_searchable_pdf(src_pdf: str, pages_text: List[Dict[str, Any]], out_path: str):
        """Tạo PDF có thể tìm kiếm và bôi đen copy chữ được bằng PyMuPDF với font Unicode."""
        doc = fitz.open(src_pdf)
        if not doc.is_pdf:
            pdf_bytes = doc.convert_to_pdf()
            doc.close()
            doc = fitz.open("pdf", pdf_bytes)

        # Detect system unicode font for Vietnamese diacritics
        font_file = None
        for candidate in ["C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/times.ttf", "C:/Windows/Fonts/tahoma.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"]:
            if os.path.exists(candidate):
                font_file = candidate
                break

        for p in pages_text:
            p_idx = p["page_num"] - 1
            if p_idx < len(doc):
                page = doc[p_idx]
                rect = page.rect
                try:
                    if font_file:
                        page.insert_font(fontname="f_ocr", fontfile=font_file)
                        page.insert_textbox(
                            rect, 
                            p["text"], 
                            fontname="f_ocr",
                            fontsize=8, 
                            color=(0, 0, 0), 
                            render_mode=3  # Invisible text for searchability
                        )
                    else:
                        page.insert_textbox(
                            rect, 
                            p["text"], 
                            fontsize=8, 
                            color=(0, 0, 0), 
                            render_mode=3
                        )
                except Exception:
                    try:
                        page.insert_textbox(rect, p["text"], fontsize=8, render_mode=3)
                    except Exception:
                        pass

        doc.save(out_path, garbage=3, deflate=True)
        doc.close()
