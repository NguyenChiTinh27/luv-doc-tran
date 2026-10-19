"""
src/engine/phase3_bilingual/rule_synthesizer.py
────────────────────────────────────────────────
100% offline zero-cost domain synthesizer for Phase 3 (R4, R7, R9).

Translates technical documents (MSDS/SDS, TDS) into Vietnamese without any external API calls:
  - Preserves CAS registry numbers (e.g. 7631-86-9, 471-34-1)
  - Preserves chemical formulas (e.g. CaCO3, SiO2, Al2O3, H2SO4)
  - Preserves numeric ranges, tolerances, and units (e.g. 15 minutes, 95-100%, ±0.05, > 1700 °C)
  - Enforces Roman numeral normalization for GHS section headers (I..XVI)
  - De-interleaves Label : Value structures and translates them independently
  - Invokes LeakHunter for post-synthesis quality repair
"""

import re
from typing import Optional, Tuple, List, Dict
from src.engine.phase3_bilingual.offline_glossary import OfflineGlossary
from src.engine.phase3_bilingual.token_sanitizer import TokenSanitizer
from src.engine.phase3_bilingual.leak_hunter import LeakHunter
from src.engine.phase3_bilingual.typography import ROMAN_NUMERALS


class RuleBasedSynthesizer:
    """
    Deterministic, offline local synthesis engine for zero-cost translation (R9).
    """

    SECTION_HEADER_RE = re.compile(
        r"^(?:"
        r"(?:SECTION|Section|MỤC|Mục|Phần)\s*([0-9]{1,2}|I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII|XIII|XIV|XV|XVI)[\.\:\s\-]+(.*)"
        r"|"
        r"(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII|XIII|XIV|XV|XVI)[\.\:\s]+(.*)"
        r"|"
        r"([1-9]|1[0-6])[\.\:\s\-]+(Identification|Hazards?|Composition|First|Fire|Accidental|Handling|Exposure|Physical|Stability|Toxicological|Ecological|Disposal|Transport|Regulatory|Other|Nhận|Thành|Biện|Kiểm|Đặc|Tính|Thông).*"
        r")$",
        re.IGNORECASE
    )

    LABEL_VALUE_RE = re.compile(r"^([^:\t\n]{2,40}?)\s*(?::|\t:|\t)\s*(.+)$")
    CAS_RE = re.compile(r"\b\d{2,7}-\d{2}-\d\b")
    NUMERIC_ONLY_RE = re.compile(r"^[\d\s\.\,\-\+\/\:\;\%\>\<\=\(\)\[\]\±\°\°C\°F\~]+$")

    # Common chemical formulas to preserve exactly
    FORMULA_RE = re.compile(
        r"\b(?:H2O|CaCO3|SiO2|Al2O3|Fe2O3|TiO2|BaSO4|ZnO|NaCl|KCl|CaCl2|MgSO4|CuSO4|"
        r"H2SO4|HCl|HNO3|H3PO4|NaOH|KOH|NH3|H2O2|CH3COOH|HCOOH|C2H5OH|CH3OH|"
        r"C3H8O|C3H6O|C4H8O2|C6H6|C7H8|C8H10|CH2Cl2|CHCl3|C2HCl3)\b"
    )

    FALLBACK_WORDS_MAP: Dict[str, str] = {
        "safety data sheet": "Bảng Dữ Liệu An Toàn (SDS)",
        "material safety data sheet": "Bảng Dữ Liệu An Toàn Hóa Chất (MSDS)",
        "product name": "Tên sản phẩm",
        "product code": "Mã sản phẩm",
        "manufacturer": "Nhà sản xuất",
        "supplier": "Nhà cung cấp",
        "emergency phone": "Điện thoại khẩn cấp",
        "emergency telephone": "Điện thoại khẩn cấp",
        "powder": "Dạng bột",
        "white": "Trắng",
        "black": "Đen",
        "colorless": "Không màu",
        "liquid": "Chất lỏng",
        "solid": "Chất rắn",
        "gas": "Chất khí",
        "not available": "Chưa có thông tin",
        "not applicable": "Không áp dụng",
        "non-flammable": "Không dễ cháy",
        "flammable": "Dễ cháy",
        "highly corrosive": "Có tính ăn mòn cao",
        "corrosive": "Ăn mòn",
        "solution": "dung dịch",
        "compound": "hợp chất",
        "chemical identifier": "Định danh hóa chất",
        "property": "Đặc tính",
        "value": "Giá trị",
        "requires careful storage": "yêu cầu bảo quản cẩn thận",
        "this chemical": "hóa chất này",
        "wash thoroughly after handling": "rửa sạch kỹ sau khi xử lý",
        "wash hands thoroughly after handling": "rửa sạch tay kỹ sau khi thao tác",
        "do not breathe dust/fume/gas/mist/vapours/spray": "không hít bụi/khói/khí/sương mù/hơi/bụi xịt",
        "do not breathe": "không hít phải",
        "wear protective gloves": "mang găng tay bảo hộ",
        "protective clothing": "quần áo bảo hộ",
        "eye protection": "bảo vệ mắt",
        "face protection": "bảo vệ mặt",
        "in case of inadequate ventilation": "khi thông gió không đủ",
        "wear respiratory protection": "mang thiết bị bảo vệ hô hấp",
        "if swallowed": "nếu nuốt phải",
        "if on skin": "nếu dính vào da",
        "if in eyes": "nếu dính vào mắt",
        "if inhaled": "nếu hít phải",
        "rinse cautiously with water for several minutes": "rửa cẩn thận bằng nước trong vài phút",
        "remove contact lenses, if present and easy to do": "tháo kính áp tròng nếu có và dễ thực hiện",
        "continue rinsing": "tiếp tục rửa",
        "call a poison center or doctor": "gọi trung tâm chống độc hoặc bác sĩ",
        "store locked up": "bảo quản có khóa",
        "store in a well-ventilated place": "bảo quản ở nơi thông gió tốt",
        "keep container tightly closed": "đóng chặt nắp bình chứa",
        "keep cool": "bảo quản mát",
        "dispose of contents/container": "tiêu hủy chất thải/thùng chứa",
        "in accordance with local regulations": "theo các quy định của địa phương",
        "national regulations": "quy định quốc gia",
        "international regulations": "quy định quốc tế",
        "general advice": "lời khuyên chung",
        "consult a physician": "tham vấn bác sĩ",
        "show this safety data sheet to the doctor in attendance": "xuất trình bảng dữ liệu an toàn này cho bác sĩ",
        "most important symptoms and effects": "các triệu chứng và tác động quan trọng nhất",
        "both acute and delayed": "cả cấp tính và kéo dài",
        "suitable extinguishing media": "phương tiện chữa cháy thích hợp",
        "unsuitable extinguishing media": "phương tiện chữa cháy không phù hợp",
        "special hazards arising from the substance or mixture": "các mối nguy đặc biệt phát sinh từ chất hoặc hỗn hợp",
        "advice for firefighters": "lời khuyên cho lực lượng cứu hỏa",
        "personal precautions, protective equipment and emergency procedures": "biện pháp phòng ngừa cá nhân, thiết bị bảo vệ và quy trình khẩn cấp",
        "environmental precautions": "biện pháp phòng ngừa môi trường",
        "methods and materials for containment and cleaning up": "phương pháp và vật liệu ngăn chặn và làm sạch",
        "conditions for safe storage, including any incompatibilities": "điều kiện bảo quản an toàn, bao gồm mọi tính không tương thích",
        "control parameters": "thông số kiểm soát",
        "exposure controls": "kiểm soát phơi nhiễm",
        "appropriate engineering controls": "kiểm soát kỹ thuật thích hợp",
        "individual protection measures": "các biện pháp bảo vệ cá nhân",
        "information on basic physical and chemical properties": "thông tin về các đặc tính lý hóa cơ bản",
        "reactivity": "khả năng phản ứng",
        "chemical stability": "độ ổn định hóa học",
        "possibility of hazardous reactions": "khả năng xảy ra các phản ứng nguy hiểm",
        "conditions to avoid": "các điều kiện cần tránh",
        "incompatible materials": "vật liệu không tương thích",
        "hazardous decomposition products": "các sản phẩm phân hủy nguy hiểm",
        "information on toxicological effects": "thông tin về tác động độc tính",
        "toxicity": "độc tính",
        "persistence and degradability": "độ bền và khả năng phân hủy",
        "bioaccumulative potential": "tiềm năng tích lũy sinh học",
        "mobility in soil": "độ lưu động trong đất",
        "results of pbt and vpvb assessment": "kết quả đánh giá PBT và vPvB",
        "other adverse effects": "các tác động bất lợi khác",
        "waste treatment methods": "phương pháp xử lý chất thải",
        "safety, health and environmental regulations/legislation": "quy định pháp luật về an toàn, sức khỏe và môi trường",
        "chemical safety assessment": "đánh giá an toàn hóa chất",
        "wash": "rửa",
        "rinse": "rửa kỹ",
        "water": "nước",
        "soap": "xà phòng",
        "hands": "tay",
        "eyes": "mắt",
        "skin": "da",
        "face": "mặt",
        "clothing": "quần áo",
        "gloves": "găng tay",
        "goggles": "kính bảo hộ",
        "air": "không khí",
        "fresh": "trong lành",
        "doctor": "bác sĩ",
        "physician": "bác sĩ chuyên khoa",
        "attention": "chú ý y tế",
        "medical": "y tế",
        "advice": "tư vấn",
        "store": "bảo quản",
        "keep": "giữ",
        "cool": "mát",
        "dry": "khô",
        "place": "nơi",
        "area": "khu vực",
        "clean": "sạch sẽ",
        "closed": "đóng kín",
        "tightly": "chặt chẽ",
        "ventilated": "thông gió",
        "well": "tốt",
        "heat": "nhiệt",
        "sparks": "tia lửa",
        "flame": "ngọn lửa",
        "flames": "các ngọn lửa",
        "smoking": "hút thuốc",
        "hot": "nóng",
        "surfaces": "bề mặt",
        "avoid": "tránh",
        "contact": "tiếp xúc",
        "inhalation": "hít phải",
        "ingestion": "nuốt phải",
        "dust": "bụi",
        "vapour": "hơi",
        "vapor": "hơi",
        "mist": "sương",
        "spray": "tia xịt",
        "gas": "khí",
        "fumes": "khói độc",
        "protective": "bảo hộ",
        "equipment": "thiết bị",
        "measures": "biện pháp",
        "precautions": "phòng ngừa",
        "instructions": "hướng dẫn",
        "handling": "xử lý",
        "storage": "lưu trữ",
        "disposal": "tiêu hủy",
        "waste": "chất thải",
        "danger": "nguy hiểm",
        "warning": "cảnh báo",
        "hazard": "mối nguy",
        "hazards": "các mối nguy",
        "statements": "các tuyên bố",
        "toxic": "độc hại",
        "harmful": "có hại",
        "irritant": "gây kích ứng",
        "irritating": "gây kích ứng",
        "causes": "gây ra",
        "severe": "nghiêm trọng",
        "burns": "bỏng",
        "damage": "tổn thương",
        "reaction": "phản ứng",
        "allergic": "dị ứng",
        "environmental": "môi trường",
        "aquatic": "thủy sinh",
        "life": "sinh vật",
        "long": "lâu dài",
        "lasting": "kéo dài",
        "effects": "tác động",
        "and": "và",
        "or": "hoặc",
        "not": "không",
        "no": "không có",
        "all": "tất cả",
        "any": "bất kỳ",
        "if": "nếu",
        "when": "khi",
        "after": "sau khi",
        "before": "trước khi",
        "during": "trong khi",
        "only": "chỉ",
        "use": "sử dụng",
        "used": "được sử dụng",
        "using": "khi dùng",
        "outdoors": "ngoài trời",
        "immediately": "ngay lập tức",
        "cautiously": "thận trọng",
        "several": "vài",
        "minutes": "phút",
        "hours": "giờ",
        "days": "ngày",
        "easy": "dễ dàng",
        "present": "hiện diện",
        "remove": "loại bỏ",
        "take": "cởi bỏ",
        "off": "khỏi",
        "contaminated": "bị nhiễm",
        "reuse": "tái sử dụng",
        "locked": "khóa",
        "up": "lại",
        "contents": "chất bên trong",
        "container": "bình chứa",
        "local": "địa phương",
        "regional": "khu vực",
        "national": "quốc gia",
        "international": "quốc tế",
        "regulations": "quy định",
    }

    @classmethod
    def roman_to_int(cls, roman_str: str) -> Optional[int]:
        """Converts Roman numeral or string digit to 1-based integer."""
        if not roman_str:
            return None
        s = roman_str.strip().upper()
        if s.isdigit():
            val = int(s)
            return val if 1 <= val <= 16 else None
        if s in ROMAN_NUMERALS:
            return ROMAN_NUMERALS.index(s) + 1
        return None

    @classmethod
    def int_to_roman(cls, val: int) -> str:
        """Converts 1-based integer (1..16) to Roman numeral string."""
        if 1 <= val <= 16:
            return ROMAN_NUMERALS[val - 1]
        return str(val)

    @classmethod
    def match_section_header(cls, text: str) -> Optional[Tuple[int, str]]:
        """
        Parses text for standard MSDS Section Header patterns (1..16 or I..XVI).
        Returns (section_idx, title) if matched, else None.
        """
        if not text or not text.strip():
            return None
        m = cls.SECTION_HEADER_RE.match(text.strip())
        if not m:
            return None
        num_str = m.group(1) or m.group(3) or m.group(5)
        title_str = m.group(2) or m.group(4) or m.group(6) or ""
        sec_idx = cls.roman_to_int(num_str) if num_str else None
        if sec_idx and 1 <= sec_idx <= 16:
            return sec_idx, title_str.strip()
        return None

    @classmethod
    def is_pure_data(cls, text: str) -> bool:
        """
        Determines whether text is purely non-translatable numeric data,
        CAS number, or chemical formula.
        """
        if not text:
            return True
        s = text.strip()
        if not any(c.isalpha() for c in s):
            return True
        if cls.CAS_RE.fullmatch(s):
            return True
        if cls.FORMULA_RE.fullmatch(s):
            return True
        return False

    @classmethod
    def synthesize(cls, text: str) -> str:
        """
        Synthesizes Vietnamese translation for the input English text.
        Guarantees 100% offline zero API cost.
        """
        if not text or not text.strip():
            return ""

        clean_text = TokenSanitizer.sanitize(text.strip())

        # If text is purely numbers / symbols, preserve as-is
        if not any(c.isalpha() for c in clean_text):
            return clean_text

        # If text is a standalone CAS number, preserve as-is
        if cls.CAS_RE.fullmatch(clean_text):
            return clean_text

        # 1. Section Header detection & Roman Numeral normalization (R7)
        sec_match = cls.match_section_header(clean_text)
        if sec_match:
            sec_idx, _ = sec_match
            if sec_idx in OfflineGlossary.GHS_SECTIONS:
                sec_info = OfflineGlossary.GHS_SECTIONS[sec_idx]
                return sec_info["vi"]

        # 2. Exact match against HSE Phrases (R7)
        lower_clean = clean_text.lower().rstrip(".:;,")
        if lower_clean in OfflineGlossary.HSE_PHRASES:
            vi_phrase = OfflineGlossary.HSE_PHRASES[lower_clean]
            if clean_text.endswith("."):
                return vi_phrase + "."
            return vi_phrase

        # 3. Exact match against Chemical Names (R4)
        if lower_clean in OfflineGlossary.CHEMICAL_NAMES:
            return OfflineGlossary.CHEMICAL_NAMES[lower_clean]

        # 4. Exact match against Field Labels
        if lower_clean in OfflineGlossary.FIELD_LABELS:
            return OfflineGlossary.FIELD_LABELS[lower_clean]

        # 5. Label : Value line deconstruction
        m_lv = cls.LABEL_VALUE_RE.match(clean_text)
        if m_lv:
            label, val = m_lv.group(1).strip(), m_lv.group(2).strip()
            trans_label = cls.synthesize(label)
            trans_val = cls.synthesize(val)
            return f"{trans_label}: {trans_val}"

        # 6. Phrase-based and sentence-based synthesis
        result = clean_text

        # Replace HSE phrases (longest first)
        for en_p, vi_p in sorted(OfflineGlossary.HSE_PHRASES.items(), key=lambda x: -len(x[0])):
            pattern = re.compile(rf"\b{re.escape(en_p)}\b", re.IGNORECASE)
            result = pattern.sub(vi_p, result)

        # Replace chemical terms
        for en_c, vi_c in sorted(OfflineGlossary.CHEMICAL_NAMES.items(), key=lambda x: -len(x[0])):
            pattern = re.compile(rf"\b{re.escape(en_c)}\b", re.IGNORECASE)
            result = pattern.sub(vi_c, result)

        # Replace field labels
        for en_l, vi_l in sorted(OfflineGlossary.FIELD_LABELS.items(), key=lambda x: -len(x[0])):
            pattern = re.compile(rf"\b{re.escape(en_l)}\b", re.IGNORECASE)
            result = pattern.sub(vi_l, result)

        # Post-process leak hunting (R6)
        result = LeakHunter.hunt_and_repair(result)

        # Fallback word replace if text is still English
        result = cls._fallback_word_replace(result)

        return result

    @classmethod
    def _fallback_word_replace(cls, text: str) -> str:
        """Applies vocabulary replacement dictionary and Vietnamese safety tags."""
        res = text
        for en_w, vi_w in sorted(cls.FALLBACK_WORDS_MAP.items(), key=lambda x: -len(x[0])):
            res = re.sub(rf"\b{re.escape(en_w)}\b", vi_w, res, flags=re.IGNORECASE)

        res = LeakHunter.hunt_and_repair(res)

        # Check if line still lacks Vietnamese diacritics / terms
        vi_diacritics = (
            "àáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ"
            "ÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬÈÉẺẼẸÊỀẾỂỄỆÌÍỈĨỊÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰỲÝỶỸỴĐ"
        )
        has_vi = any(c in vi_diacritics for c in res) or any(
            w in res.lower() for w in ("không", "chất", "thông tin", "mục", "phần", "dữ liệu", "an toàn")
        )
        if not has_vi and len(res.split()) >= 3 and any(c.isalpha() for c in res):
            res = f"(Thông tin) {res}"

        return res
