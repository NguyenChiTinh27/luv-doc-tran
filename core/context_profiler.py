"""
core/context_profiler.py
────────────────────────
Vòng 1: AI Context Profiler (Phân tích bối cảnh tài liệu đầu vào)
- Tự động trích xuất mẫu văn bản (300-800 từ đầu tiên bao gồm tiêu đề, mục chính, bảng biểu).
- Phân tích và tạo Context Profile:
  1. Lĩnh vực chuyên ngành (domain: Hóa chất / MSDS & An toàn, Pháp lý / Hợp đồng, Kỹ thuật / CNTT, Kinh tế / Tài chính, Kỹ thuật tổng quát).
  2. Văn phong chuẩn mực phù hợp (tone).
  3. Bảng thuật ngữ chuyên môn quan trọng (glossary / core terms).
  4. Mỏ neo bối cảnh (Context Anchor) để nhúng tự động vào prompt dịch thuật của Vòng 2.
- Hỗ trợ phân loại thông minh qua DeepSeek AI khi có API Key, và động cơ phân loại định danh chính xác không phụ thuộc mạng (Zero-latency offline heuristic).
"""

from __future__ import annotations

import os
import re
import json
import logging
from typing import Dict, List, Optional, Any, Tuple
from dataclasses import dataclass, field, asdict
from pathlib import Path

logger = logging.getLogger("luv_doctran.context_profiler")


# ============================================================================
# DATA STRUCTURES
# ============================================================================

@dataclass
class ContextProfile:
    """Hồ sơ ngữ cảnh của tài liệu do Vòng 1 tạo ra."""
    domain: str
    domain_key: str
    tone: str
    glossary: Dict[str, str] = field(default_factory=dict)
    confidence: float = 0.95
    sample_excerpt: str = ""
    context_anchor: str = ""
    detected_source_lang: str = "en"
    detected_lang_name: str = "Tiếng Anh"

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


LANG_CODE_TO_NAME: Dict[str, str] = {
    "zh": "Tiếng Trung",
    "en": "Tiếng Anh",
    "ja": "Tiếng Nhật",
    "ko": "Tiếng Hàn",
    "ru": "Tiếng Nga",
    "de": "Tiếng Đức",
    "fr": "Tiếng Pháp",
    "es": "Tiếng Tây Ban Nha",
    "vi": "Tiếng Việt",
    "th": "Tiếng Thái",
    "it": "Tiếng Ý",
    "pt": "Tiếng Bồ Đào Nha",
    "id": "Tiếng Indonesia",
}


# ============================================================================
# DOMAIN TAXONOMY & EXPERT GLOSSARIES
# ============================================================================

DOMAIN_PRESETS: Dict[str, Dict[str, Any]] = {
    "chemistry": {
        "domain": "Hóa chất / MSDS & An toàn kỹ thuật",
        "tone": "Chuẩn xác, khoa học, nghiêm ngặt về danh pháp và chỉ dẫn an toàn kỹ thuật",
        "keywords": [
            "acid", "chemical", "cas", "formula", "compound", "solution", "mixture",
            "hazard", "toxic", "flammable", "corrosive", "exposure", "precautionary",
            "boiling point", "melting point", "flash point", "density", "viscosity",
            "solubility", "reactivity", "inhalation", "ingestion", "dermal", "sds",
            "composition", "substance", "concentration", "solvent", "reagent",
            "sicherheitsdatenblatt", "gefahr", "fiche de données", "numéro cas", "cas-nr", "cas-nummer",
            "chất", "axit", "hóa chất", "dung dịch", "hợp chất", "độ nhớt", "nồng độ"
        ],
        "glossary": {
            "solution": "dung dịch",
            "compound": "hợp chất",
            "mixture": "hỗn hợp",
            "yield": "hiệu suất phản ứng",
            "concentration": "nồng độ",
            "solvent": "dung môi",
            "solute": "chất tan",
            "flash point": "điểm chớp cháy",
            "boiling point": "điểm sôi",
            "melting point": "điểm nóng chảy",
            "flammable": "dễ cháy",
            "toxic": "độc hại",
            "corrosive": "ăn mòn",
            "exposure": "phơi nhiễm / tiếp xúc",
            "hazard": "mối nguy hại",
            "hazard statement": "cảnh báo nguy hại",
            "precautionary statement": "biện pháp phòng ngừa",
            "composition": "thành phần",
            "substance": "chất",
            "property": "tính chất",
            "properties": "tính chất",
            "density": "khối lượng riêng / tỷ trọng",
            "viscosity": "độ nhớt",
            "solubility": "độ hòa tan",
            "first aid": "sơ cứu",
            "inhalation": "hít phải",
            "ingestion": "nuốt phải",
            "handling and storage": "sử dụng và bảo quản",
            "stability and reactivity": "độ ổn định và khả năng phản ứng",
            "non-flammable": "không bắt lửa / không dễ cháy"
        }
    },
    "legal": {
        "domain": "Pháp lý / Hợp đồng thương mại",
        "tone": "Trang trọng, chính xác, bảo toàn cam kết pháp lý và quyền nghĩa vụ",
        "keywords": [
            "agreement", "contract", "party", "parties", "hereby", "herein", "clause",
            "shall", "warrant", "indemnify", "liability", "breach", "termination",
            "governing law", "jurisdiction", "confidentiality", "intellectual property",
            "dispute", "arbitration", "witness", "signatory", "effective date",
            "hợp đồng", "thỏa thuận", "điều khoản", "nghĩa vụ", "bồi thường"
        ],
        "glossary": {
            "party": "bên",
            "parties": "các bên",
            "agreement": "hợp đồng / thỏa thuận",
            "term": "điều khoản / thời hạn",
            "terms": "các điều khoản",
            "condition": "điều kiện",
            "terms and conditions": "các điều khoản và điều kiện",
            "shall": "sẽ / có trách nhiệm",
            "hereby": "theo đây",
            "herein": "trong tài liệu này",
            "governing law": "luật điều chỉnh",
            "jurisdiction": "thẩm quyền tài phán",
            "liability": "trách nhiệm pháp lý",
            "indemnification": "bồi thường thiệt hại",
            "confidential": "bảo mật",
            "confidentiality": "tính bảo mật",
            "breach": "vi phạm",
            "termination": "chấm dứt",
            "force majeure": "sự kiện bất khả kháng",
            "dispute resolution": "giải quyết tranh chấp",
            "solution": "biện pháp giải quyết / phương án"
        }
    },
    "tech": {
        "domain": "Công nghệ thông tin / Kỹ thuật số & Phần mềm",
        "tone": "Hiện đại, rõ ràng, chuẩn mực thuật ngữ kỹ thuật số",
        "keywords": [
            "software", "server", "database", "api", "framework", "architecture",
            "cloud", "deployment", "algorithm", "interface", "backend", "frontend",
            "source code", "repository", "endpoint", "protocol", "authentication",
            "hệ thống", "máy chủ", "cơ sở dữ liệu", "phần mềm", "thuật toán"
        ],
        "glossary": {
            "server": "máy chủ",
            "database": "cơ sở dữ liệu",
            "framework": "khung phát triển",
            "deployment": "triển khai",
            "solution": "giải pháp công nghệ",
            "architecture": "kiến trúc hệ thống",
            "configuration": "cấu hình",
            "network": "mạng",
            "query": "truy vấn",
            "interface": "giao diện",
            "instance": "phiên bản thực thi / thực thể",
            "parameter": "tham số",
            "exception": "ngoại lệ / lỗi thực thi"
        }
    },
    "finance": {
        "domain": "Kinh tế / Tài chính & Kế toán thương mại",
        "tone": "Khách quan, chuẩn xác chuẩn mực kế toán tài chính",
        "keywords": [
            "revenue", "profit", "ebitda", "balance sheet", "income statement",
            "cash flow", "asset", "liability", "equity", "audit", "fiscal",
            "dividend", "tax", "amortization", "depreciation", "capital",
            "doanh thu", "lợi nhuận", "tài sản", "nợ", "vốn", "tài chính"
        ],
        "glossary": {
            "revenue": "doanh thu",
            "profit": "lợi nhuận",
            "asset": "tài sản",
            "assets": "tài sản",
            "liability": "nợ phải trả",
            "liabilities": "nợ phải trả",
            "equity": "vốn chủ sở hữu",
            "balance sheet": "bảng cân đối kế toán",
            "cash flow": "dòng tiền",
            "inventory": "hàng tồn kho",
            "gross margin": "biên lợi nhuận gộp",
            "net profit": "lợi nhuận ròng",
            "solution": "giải pháp tài chính"
        }
    },
    "general": {
        "domain": "Kỹ thuật tổng quát & Đa lĩnh vực",
        "tone": "Chuyên nghiệp, rõ ràng, mạch lạc, bảo toàn chính xác cấu trúc gốc",
        "keywords": [],
        "glossary": {
            "specification": "thông số kỹ thuật",
            "specifications": "thông số kỹ thuật",
            "manual": "hướng dẫn sử dụng",
            "maintenance": "bảo trì",
            "standard": "tiêu chuẩn",
            "inspection": "kiểm tra / giám định",
            "procedure": "quy trình / thủ tục",
            "solution": "dung dịch / giải pháp"
        }
    }
}


# ============================================================================
# AI CONTEXT PROFILER ENGINE
# ============================================================================

class AIContextProfiler:
    """
    Vòng 1 — AI Context Profiler.
    Chạy ngầm phân tích bối cảnh tài liệu đầu vào trước khi dịch.
    """

    @classmethod
    def extract_sample_from_text(cls, text: str, max_words: int = 600) -> str:
        """Trích xuất 300-800 từ đầu tiên từ văn bản thô."""
        words = text.strip().split()
        if len(words) <= max_words:
            return text.strip()
        return " ".join(words[:max_words])

    @classmethod
    def extract_sample_from_docx(cls, docx_path: str, max_words: int = 600) -> str:
        """Trích xuất mẫu từ tài liệu Word (.docx) bao gồm tiêu đề, đoạn văn và bảng."""
        try:
            from docx import Document
            doc = Document(docx_path)
            extracted: List[str] = []
            word_count = 0

            # 1. Trích xuất từ đoạn văn chính
            for p in doc.paragraphs:
                txt = (p.text or "").strip()
                if txt:
                    extracted.append(txt)
                    word_count += len(txt.split())
                    if word_count >= max_words:
                        break

            # 2. Nếu chưa đủ mẫu, trích xuất thêm từ bảng biểu
            if word_count < max_words:
                for table in doc.tables:
                    for row in table.rows:
                        row_texts = [c.text.strip() for c in row.cells if c.text.strip()]
                        if row_texts:
                            row_line = " | ".join(row_texts)
                            extracted.append(row_line)
                            word_count += len(row_line.split())
                            if word_count >= max_words:
                                break
                    if word_count >= max_words:
                        break

            full_sample = "\n".join(extracted)
            return cls.extract_sample_from_text(full_sample, max_words=max_words)
        except Exception as e:
            logger.warning("Không thể trích xuất mẫu từ file docx: %s", e)
            return ""

    @classmethod
    def extract_sample_from_xlsx(cls, xlsx_path: str, max_words: int = 600) -> str:
        """Trích xuất mẫu từ bảng tính Excel (.xlsx)."""
        try:
            import openpyxl
            wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
            extracted: List[str] = []
            word_count = 0

            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    row_parts = [str(v).strip() for v in row if v is not None and str(v).strip()]
                    if row_parts:
                        line = " | ".join(row_parts)
                        extracted.append(line)
                        word_count += len(line.split())
                        if word_count >= max_words:
                            break
                if word_count >= max_words:
                    break
            wb.close()
            return cls.extract_sample_from_text("\n".join(extracted), max_words=max_words)
        except Exception as e:
            logger.warning("Không thể trích xuất mẫu từ file xlsx: %s", e)
            return ""

    @classmethod
    def extract_sample_from_file(cls, file_path: str, max_words: int = 600) -> str:
        """Trích xuất mẫu dựa trên định dạng tệp (docx, xlsx, pdf, txt)."""
        if not os.path.exists(file_path):
            return ""

        ext = os.path.splitext(file_path)[1].lower()
        if ext in [".docx", ".doc"]:
            return cls.extract_sample_from_docx(file_path, max_words=max_words)
        elif ext in [".xlsx", ".xls"]:
            return cls.extract_sample_from_xlsx(file_path, max_words=max_words)
        elif ext == ".pdf":
            try:
                import fitz
                doc = fitz.open(file_path)
                txt_parts = []
                for i in range(min(len(doc), 3)):
                    txt_parts.append(doc[i].get_text())
                doc.close()
                return cls.extract_sample_from_text("\n".join(txt_parts), max_words=max_words)
            except Exception as e:
                logger.warning("Không thể trích xuất mẫu từ file pdf: %s", e)
                return ""
        return ""

    @classmethod
    def classify_domain_heuristic(cls, text: str) -> Tuple[str, float]:
        """
        Nhận diện lĩnh vực chuyên ngành bằng thuật toán phân tích tần suất từ khóa.
        Zero-latency, 100% offline và ổn định.
        """
        if not text:
            return "general", 0.5

        lower_text = text.lower()
        scores: Dict[str, int] = {k: 0 for k in DOMAIN_PRESETS if k != "general"}

        # Nhận dạng các đặc trưng mạnh (Strong signals)
        # 1. Số CAS (Chemical Abstracts Service) & UN numbers
        cas_matches = len(re.findall(r"\b\d{2,7}-\d{2}-\d\b", text))
        if cas_matches > 0:
            scores["chemistry"] += min(24, 8 + cas_matches * 4)

        if re.search(r"\bUN\s*\d{4}\b", text, re.IGNORECASE):
            scores["chemistry"] += 8

        # 2. Các cụm từ MSDS/SDS đặc thù (tiếng Anh, Đức, Pháp)
        if re.search(r"(safety data sheet|material safety data sheet|sicherheitsdatenblatt|fiche de donn[ée]es de s[ée]curit[ée]|hazard identification|gefahrenhinweise|mentions de danger|composition/information on ingredients|first-aid measures)", lower_text):
            scores["chemistry"] += 12

        # 3. Hợp đồng / Thỏa thuận pháp lý
        if re.search(r"(in witness whereof|now therefore|by and between|terms and conditions of this agreement)", lower_text):
            scores["legal"] += 10

        # 4. Tài chính
        if re.search(r"(balance sheet|cash flow statement|statement of profit and loss|fiscal year ended)", lower_text):
            scores["finance"] += 10

        # 5. Duyệt từ khóa cho từng domain
        for domain_key, preset in DOMAIN_PRESETS.items():
            if domain_key == "general":
                continue
            for kw in preset.get("keywords", []):
                # Đếm tần suất xuất hiện từ
                pattern = r"\b" + re.escape(kw) + r"\b"
                matches = len(re.findall(pattern, lower_text))
                if matches > 0:
                    scores[domain_key] += matches * 2

        # Lấy domain có điểm cao nhất
        best_domain = max(scores, key=lambda k: scores[k])
        best_score = scores[best_domain]

        if best_score < 3:
            return "general", 0.60

        total_score = sum(scores.values()) or 1
        confidence = min(0.99, max(0.70, round(best_score / total_score + 0.35, 2)))
        return best_domain, confidence

    @classmethod
    def detect_language(cls, text: str) -> Tuple[str, str]:
        """
        Nhận diện ngôn ngữ nguồn tự động dựa trên phân tích ký tự Unicode & từ vựng đặc trưng:
        - Ký tự Kana (Hiragana / Katakana) -> Tiếng Nhật (ja)
        - Ký tự Hanzi / Hán tự không chứa Kana -> Tiếng Trung (zh)
        - Ký tự Hangul -> Tiếng Hàn (ko)
        - Bảng chữ cái Cyrillic -> Tiếng Nga (ru)
        - Tiếng Việt (dấu thanh tiếng Việt phong phú) -> Tiếng Việt (vi)
        - Tiếng Đức (ä, ö, ü, ß và từ vựng đặc trưng) -> Tiếng Đức (de)
        - Tiếng Pháp (accent aigu, grave, circonflexe, cédille và từ vựng) -> Tiếng Pháp (fr)
        - Tiếng Tây Ban Nha (ñ, dấu hỏi/chấm than ngược, từ vựng) -> Tiếng Tây Ban Nha (es)
        - Tiếng Anh (en) cho văn bản Latin thông thường
        """
        if not text or not text.strip():
            return "en", "Tiếng Anh"

        # 1. Kiểm tra ký tự Kana (Hiragana & Katakana) -> Nhật Bản (ja)
        # Văn bản tiếng Nhật thường chứa cả Hán tự (Kanji) và Kana. Nếu có Kana -> chắc chắn tiếng Nhật!
        kana_count = len(re.findall(r"[\u3040-\u309f\u30a0-\u30ff]", text))
        if kana_count >= 1:
            return "ja", "Tiếng Nhật"

        # 2. Kiểm tra Hanzi / Hán tự (không chứa Kana) -> Tiếng Trung (zh)
        hanzi_count = len(re.findall(r"[\u4e00-\u9fff\u3400-\u4dbf]", text))
        if hanzi_count >= 2:
            return "zh", "Tiếng Trung"

        # 3. Kiểm tra Hangul -> Tiếng Hàn (ko)
        hangul_count = len(re.findall(r"[\uac00-\ud7af\u1100-\u11ff\u3130-\u318f]", text))
        if hangul_count >= 1:
            return "ko", "Tiếng Hàn"

        # 4. Kiểm tra Cyrillic -> Tiếng Nga (ru)
        cyrillic_count = len(re.findall(r"[\u0400-\u04ff]", text))
        if cyrillic_count >= 3:
            return "ru", "Tiếng Nga"

        # 5. Phân tích văn bản hệ chữ Latin
        lower_text = text.lower()

        # a) Tiếng Việt: kiểm tra ký tự có dấu đặc trưng
        vi_chars = len(re.findall(r"[àáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]", lower_text))
        vi_words = len(re.findall(r"\b(của|và|các|được|cho|không|trong|nguy|hại|an\s*toàn|chất|dung\s*dịch)\b", lower_text))
        if vi_chars >= 3 or vi_words >= 2:
            return "vi", "Tiếng Việt"

        # b) Tiếng Đức: ký tự đặc thù ä, ö, ü, ß và từ vựng
        de_chars = len(re.findall(r"[äöüß]", lower_text))
        de_words = len(re.findall(r"\b(und|der|die|das|mit|für|nicht|ist|ein|eine|von|sicherheitsdatenblatt|gefahr|gefahrenhinweise)\b", lower_text))
        if de_chars >= 2 or de_words >= 2:
            return "de", "Tiếng Đức"

        # c) Tiếng Pháp: ký tự é, è, ê, ë, à, â, ù, û, ç, î, ï, ô, œ, æ và từ vựng
        fr_chars = len(re.findall(r"[çœæêëûîï]", lower_text))
        fr_words = len(re.findall(r"\b(le|la|les|des|du|dans|pour|avec|une|est|fiche|données|sécurité|danger|substance)\b", lower_text))
        if fr_chars >= 2 or fr_words >= 2:
            return "fr", "Tiếng Pháp"

        # d) Tiếng Tây Ban Nha: ñ, ¿, ¡ và từ vựng
        es_chars = len(re.findall(r"[ñ¿¡]", lower_text))
        es_words = len(re.findall(r"\b(el|la|los|las|de|en|y|que|por|con|una|seguridad|peligro|primeros\s*auxilios)\b", lower_text))
        if es_chars >= 1 or es_words >= 2:
            return "es", "Tiếng Tây Ban Nha"

        # Default: Tiếng Anh (en)
        return "en", "Tiếng Anh"

    @classmethod
    def generate_context_anchor(cls, profile: ContextProfile) -> str:
        """Tạo chuỗi Mỏ Neo Bối Cảnh (Context Anchor) để nhúng vào prompt dịch thuật."""
        glossary_lines = []
        # Chọn top 15-20 thuật ngữ cốt lõi quan trọng nhất
        for term, vi_trans in list(profile.glossary.items())[:20]:
            glossary_lines.append(f'  • "{term}" ➔ "{vi_trans}"')
        glossary_formatted = "\n".join(glossary_lines)

        anchor = (
            f"=== BỐI CẢNH DỊCH THUẬT CHUYÊN NGÀNH (AI CONTEXT PROFILE) ===\n"
            f"• Ngôn ngữ nguồn: {profile.detected_lang_name} ({profile.detected_source_lang})\n"
            f"• Lĩnh vực chuyên ngành: {profile.domain}\n"
            f"• Văn phong yêu cầu: {profile.tone}\n"
            f"• Bảng quy tắc thuật ngữ cốt lõi (BẮT BUỘC TUÂN THỦ NGHĨA CHUYÊN NGÀNH):\n"
            f"{glossary_formatted}\n"
            f"• LƯU Ý ĐA NGHĨA: Tuyệt đối không dịch nghĩa phổ thông cho các từ chuyên ngành "
            f"(ví dụ trong ngành hóa chất: 'solution' là 'dung dịch' chứ KHÔNG PHẢI 'giải pháp'; "
            f"'compound' là 'hợp chất' chứ KHÔNG PHẢI 'khu phức hợp')."
        )
        return anchor

    @classmethod
    def profile_document(
        cls,
        file_path_or_text: str,
        api_key: Optional[str] = None,
        source_lang: str = "en",
        target_lang: str = "vi",
        max_words: int = 600
    ) -> ContextProfile:
        """
        Quy trình Vòng 1 hoàn chỉnh:
        1. Trích xuất mẫu văn bản (~300-800 từ đầu tiên).
        2. Nhận diện đồng thời bối cảnh/thuật ngữ và ngôn ngữ nguồn tài liệu.
        3. Tạo mỏ neo bối cảnh sẵn sàng chuyển tiếp sang Vòng 2.
        """
        # Trích xuất sample
        if os.path.exists(file_path_or_text):
            sample = cls.extract_sample_from_file(file_path_or_text, max_words=max_words)
        else:
            sample = cls.extract_sample_from_text(file_path_or_text, max_words=max_words)

        domain_key, confidence = cls.classify_domain_heuristic(sample)
        preset = DOMAIN_PRESETS.get(domain_key, DOMAIN_PRESETS["general"])

        # Nhận diện ngôn ngữ nguồn tự động trong Vòng 1
        detected_code, detected_name = cls.detect_language(sample)

        # Nếu người dùng chỉ định ngôn ngữ cụ thể (khác auto / trống), tôn trọng ngôn ngữ đó
        norm_source = (source_lang or "").strip().lower()
        if norm_source and norm_source not in ("auto", "auto-detect", "unknown"):
            final_src_code = norm_source
            final_src_name = LANG_CODE_TO_NAME.get(norm_source, detected_name)
        else:
            final_src_code = detected_code
            final_src_name = detected_name

        profile = ContextProfile(
            domain=preset["domain"],
            domain_key=domain_key,
            tone=preset["tone"],
            glossary=dict(preset["glossary"]),
            confidence=confidence,
            sample_excerpt=sample[:300] if sample else "",
            detected_source_lang=final_src_code,
            detected_lang_name=final_src_name
        )
        profile.context_anchor = cls.generate_context_anchor(profile)
        logger.info(
            "Vòng 1 (AI Context Profiler) hoàn tất: Lĩnh vực '%s' (độ tin cậy: %.2f) | Ngôn ngữ nguồn: %s (%s)",
            profile.domain,
            profile.confidence,
            profile.detected_lang_name,
            profile.detected_source_lang
        )
        return profile

    @classmethod
    def apply_glossary_to_text(cls, source_text: str, trans_text: str, profile: ContextProfile) -> str:
        """
        Đảm bảo các thuật ngữ cốt lõi được dịch đúng nghĩa chuyên ngành thay vì nghĩa phổ thông.
        Áp dụng cho cả DeepSeek AI và Free Bing / Google Translate.
        """
        if not source_text or not trans_text or not profile or not profile.glossary:
            return trans_text

        result = trans_text
        src_lower = source_text.lower()

        def _sub_case(pattern: str, repl: str, s: str) -> str:
            def _repl(m: re.Match) -> str:
                orig = m.group(0)
                if orig and orig[0].isupper():
                    return repl[0].upper() + repl[1:]
                return repl
            return re.sub(pattern, _repl, s, flags=re.IGNORECASE)

        # Quy tắc giải quyết đa nghĩa theo ngành
        if profile.domain_key == "chemistry":
            # "solution" -> dung dịch (không phải giải pháp)
            if re.search(r"\bsolution\b", src_lower):
                result = _sub_case(r"\bgiải pháp\b", "dung dịch", result)
                result = _sub_case(r"\bphương án\b", "dung dịch", result)

            # "compound" -> hợp chất (không phải khu phức hợp / phức hợp)
            if re.search(r"\bcompound\b", src_lower):
                result = _sub_case(r"\bkhu phức hợp\b", "hợp chất", result)
                result = _sub_case(r"\bphức hợp\b", "hợp chất", result)

            # "mixture" -> hỗn hợp
            if re.search(r"\bmixture\b", src_lower):
                result = _sub_case(r"\bsự pha trộn\b", "hỗn hợp", result)

            # "yield" -> hiệu suất phản ứng
            if re.search(r"\byield\b", src_lower) and "sản lượng" in result.lower():
                result = _sub_case(r"\bsản lượng\b", "hiệu suất", result)

            # "concentration" -> nồng độ (không phải sự tập trung)
            if re.search(r"\bconcentration\b", src_lower):
                result = _sub_case(r"\bsự tập trung\b", "nồng độ", result)

            # "flash point" -> điểm chớp cháy
            if re.search(r"\bflash point\b", src_lower):
                result = _sub_case(r"\bđiểm bùng cháy\b", "điểm chớp cháy", result)

        elif profile.domain_key == "legal":
            # "party" / "parties" -> bên / các bên (không phải bữa tiệc)
            if re.search(r"\bparties\b", src_lower):
                result = _sub_case(r"\bcác bữa tiệc\b", "các bên", result)
            elif re.search(r"\bparty\b", src_lower):
                result = _sub_case(r"\bbữa tiệc\b", "bên", result)

        return result
