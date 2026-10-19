"""
src/engine/phase3_bilingual/offline_glossary.py
────────────────────────────────────────────────
Canonical offline domain dictionary and glossary for Phase 3 Bilingual Engine (R4, R7, R9).

Contains:
  - Canonical GHS 16 sections (Roman numerals I..XVI, English, Vietnamese)
  - 60+ chemical substances with standardized Vietnamese chemical nomenclature
  - Mandatory HSE regulatory phrases & safety terminology (strictly enforcing
    "Không hút thuốc, ăn uống", "Gây mẫn cảm da", etc.)
  - Standard MSDS property and section labels (over 100 entries)
"""

import re
from typing import Dict, Optional, Tuple, List


class OfflineGlossary:
    """
    Comprehensive offline domain dictionary for MSDS/SDS documents.
    Enforces standardized nomenclature and zero-cost local translation.
    """

    # ─────────────────────────────────────────────────────────────────────────
    # 1. 16 CANONICAL GHS SECTION HEADERS (R7)
    # ─────────────────────────────────────────────────────────────────────────
    GHS_SECTIONS: Dict[int, Dict[str, str]] = {
        1: {
            "roman": "I.",
            "en": "Identification of the substance/mixture and of the company/undertaking",
            "vi": "I. Nhận dạng hóa chất và thông tin về nhà cung cấp"
        },
        2: {
            "roman": "II.",
            "en": "Hazards Identification",
            "vi": "II. Nhận dạng đặc tính nguy hiểm"
        },
        3: {
            "roman": "III.",
            "en": "Composition / Information on Ingredients",
            "vi": "III. Thông tin về thành phần các chất"
        },
        4: {
            "roman": "IV.",
            "en": "First Aid Measures",
            "vi": "IV. Biện pháp sơ cứu về y tế"
        },
        5: {
            "roman": "V.",
            "en": "Fire Fighting Measures",
            "vi": "V. Biện pháp xử lý khi có hỏa hoạn"
        },
        6: {
            "roman": "VI.",
            "en": "Accidental Release Measures",
            "vi": "VI. Biện pháp phòng ngừa sự cố tràn đổ"
        },
        7: {
            "roman": "VII.",
            "en": "Handling and Storage",
            "vi": "VII. Sử dụng và bảo quản"
        },
        8: {
            "roman": "VIII.",
            "en": "Exposure Controls / Personal Protection",
            "vi": "VIII. Kiểm soát phơi nhiễm / Bảo vệ cá nhân"
        },
        9: {
            "roman": "IX.",
            "en": "Physical and Chemical Properties",
            "vi": "IX. Đặc tính lý, hóa học"
        },
        10: {
            "roman": "X.",
            "en": "Stability and Reactivity",
            "vi": "X. Tính ổn định và khả năng phản ứng"
        },
        11: {
            "roman": "XI.",
            "en": "Toxicological Information",
            "vi": "XI. Thông tin về độc tính"
        },
        12: {
            "roman": "XII.",
            "en": "Ecological Information",
            "vi": "XII. Thông tin về sinh thái"
        },
        13: {
            "roman": "XIII.",
            "en": "Disposal Considerations",
            "vi": "XIII. Biện pháp tiêu hủy"
        },
        14: {
            "roman": "XIV.",
            "en": "Transport Information",
            "vi": "XIV. Thông tin khi vận chuyển"
        },
        15: {
            "roman": "XV.",
            "en": "Regulatory Information",
            "vi": "XV. Thông tin về quy chuẩn pháp luật"
        },
        16: {
            "roman": "XVI.",
            "en": "Other Information",
            "vi": "XVI. Thông tin khác"
        }
    }

    # ─────────────────────────────────────────────────────────────────────────
    # 2. 60+ CHEMICAL SUBSTANCES (R4)
    # ─────────────────────────────────────────────────────────────────────────
    CHEMICAL_NAMES: Dict[str, str] = {
        "amorphous silicon dioxide, chemically prepared": "Silic đioxit vô định hình, điều chế hóa học",
        "amorphous silicon dioxide": "Silic đioxit vô định hình",
        "silicon dioxide": "Silic đioxit",
        "aluminium oxide": "Nhôm oxit",
        "alumina": "Nhôm oxit",
        "water": "Nước",
        "calcium carbonate": "Canxi cacbonat",
        "sulfuric acid": "Axit sunfuric",
        "hydrochloric acid": "Axit clohydric",
        "nitric acid": "Axit nitric",
        "phosphoric acid": "Axit photphoric",
        "sodium hydroxide": "Natri hydroxit",
        "potassium hydroxide": "Kali hydroxit",
        "toluene": "Toluen",
        "xylene": "Xylen",
        "butyl acetate": "Butyl axetat",
        "ethyl acetate": "Etyl axetat",
        "methyl ethyl ketone": "Metyl etyl xeton",
        "acetone": "Axeton",
        "ethanol": "Etanol",
        "methanol": "Metanol",
        "isopropanol": "Isopropanol",
        "isopropyl alcohol": "Rượu isopropylic",
        "titanium dioxide": "Titan đioxit",
        "talc powder": "Bột Talc",
        "talc": "Bột Talc",
        "barium sulfate": "Bari sunfat",
        "carbon black": "Muội than",
        "zinc oxide": "Kẽm oxit",
        "iron oxide": "Sắt oxit",
        "tripropylene glycol diacrylate": "Tripropylen glycol điacrylat",
        "tripropylene glycol diacrylate (tpgda)": "Tripropylen glycol điacrylat (TPGDA)",
        "1-hydroxycyclohexyl phenyl ketone": "1-Hydroxycyclohexyl phenyl xeton",
        "epoxy acrylate resin": "Nhựa epoxy acrylat",
        "polyester acrylate resin": "Nhựa polyeste acrylat",
        "epoxy resin": "Nhựa epoxy",
        "polyurethane": "Polyuretan",
        "polyurethane resin": "Nhựa Polyuretan",
        "alumina solution": "Dung dịch Nhôm oxit",
        "titanium dioxide coating": "Lớp phủ titan đioxit",
        "acrylic resin": "Nhựa acrylic",
        "formaldehyde": "Fomanđehit",
        "benzene": "Benzen",
        "phenol": "Phenol",
        "glycerol": "Glyxerol",
        "glycerin": "Glyxerin",
        "ethylene glycol": "Etylen glycol",
        "propylene glycol": "Propylen glycol",
        "ammonia": "Amoniac",
        "hydrogen peroxide": "Hydro peroxit",
        "acetic acid": "Axit axetic",
        "formic acid": "Axit fomic",
        "citric acid": "Axit xitric",
        "sodium chloride": "Natri clorua",
        "potassium chloride": "Kali clorua",
        "calcium chloride": "Canxi clorua",
        "magnesium sulfate": "Magie sunfat",
        "copper sulfate": "Đồng sunfat",
        "sodium hypochlorite": "Natri hypoclorit",
        "lead": "Chì",
        "mercury": "Thủy ngân",
        "cadmium": "Cadmi",
        "chromium": "Crom",
        "nickel": "Niken",
        "styrene": "Styren",
        "isocyanates": "Isocyanat",
        "methylene chloride": "Metylen clorua",
        "dichloromethane": "Điclorometan",
        "trichloroethylene": "Tricloroetylen",
        "chloroform": "Clorofom",
    }

    # ─────────────────────────────────────────────────────────────────────────
    # 3. MANDATORY HSE REGULATORY PHRASES (R7)
    # ─────────────────────────────────────────────────────────────────────────
    HSE_PHRASES: Dict[str, str] = {
        "no smoking, drinking and eating": "Không hút thuốc, ăn uống",
        "no smoking, drinking or eating": "Không hút thuốc, ăn uống",
        "no smoking, eating and drinking": "Không hút thuốc, ăn uống",
        "no eating, drinking and smoking": "Không hút thuốc, ăn uống",
        "material safety data sheet": "Bảng Dữ Liệu An Toàn Hóa Chất (MSDS/SDS)",
        "material safety data sheet (msds)": "Bảng Dữ Liệu An Toàn Hóa Chất (MSDS/SDS)",
        "safety data sheet": "Bảng Dữ Liệu An Toàn (SDS)",
        "safety data sheet (sds)": "Bảng Dữ Liệu An Toàn (SDS)",
        "skin sensitization": "Gây mẫn cảm da",
        "sensitization by skin contact": "Gây mẫn cảm khi tiếp xúc với da",
        "may cause sensitization by skin contact": "Có thể gây mẫn cảm khi tiếp xúc với da",
        "eye irritation": "Gây kích ứng mắt",
        "irritating to eyes": "Gây kích ứng mắt",
        "causes serious eye irritation": "Gây kích ứng mắt nghiêm trọng",
        "causes skin irritation": "Gây kích ứng da",
        "harmful in contact with skin": "Có hại khi tiếp xúc với da",
        "harmful if swallowed": "Có hại nếu nuốt phải",
        "harmful if inhaled": "Có hại nếu hít phải",
        "keep away from ignition sources, heat and flame": "Tránh xa các nguồn đánh lửa, nguồn nhiệt và ngọn lửa",
        "keep away from heat, sparks, open flames and hot surfaces": "Tránh xa nguồn nhiệt, tia lửa, ngọn lửa trần và các bề mặt nóng",
        "check for and remove any contact lenses": "Kiểm tra và tháo bỏ kính áp tròng nếu có",
        "in case of contact, immediately flush eyes with plenty of water for at least 15 minutes": "Trong trường hợp tiếp xúc, ngay lập tức rửa mắt với nhiều nước trong ít nhất 15 phút",
        "cold water may be used": "Có thể dùng nước lạnh",
        "warm water must be used": "BẮT BUỘC dùng nước ẤM",
        "get medical attention": "Đưa đi cấp cứu y tế",
        "get medical attention immediately": "Đưa đi cấp cứu y tế ngay lập tức",
        "if inhaled, remove to fresh air": "Nếu hít phải, đưa ra nơi không khí trong lành",
        "if not breathing, give artificial respiration": "Nếu ngừng thở, tiến hành hô hấp nhân tạo",
        "if breathing is difficult, give oxygen": "Nếu khó thở, cho thở oxy",
        "do not induce vomiting unless directed to do so by medical personnel": "KHÔNG kích nôn trừ khi có chỉ định của nhân viên y tế",
        "never give anything by mouth to an unconscious person": "Không bao giờ cho bất cứ thứ gì vào miệng người bất tỉnh",
        "loosen tight clothing such as a collar, tie, belt or waistband": "Nới lỏng quần áo bó sát như cổ áo, cà vạt, thắt lưng",
        "get medical attention if symptoms appear": "Đưa đi cấp cứu y tế nếu xuất hiện các triệu chứng",
        "assure adequate flushing by separating the eyelids with fingers": "Đảm bảo rửa sạch hoàn toàn bằng cách dùng ngón tay vạch mi mắt",
        "wash thoroughly after handling": "Rửa sạch toàn thân sau khi thao tác",
        "store in a well-ventilated place": "Bảo quản nơi thông thoáng",
        "keep container tightly closed": "Giữ thùng chứa đóng chặt",
        "dispose of contents/container in accordance with local regulations": "Tiêu hủy chất bên trong/thùng chứa theo quy định địa phương",
        "skin exposure": "Tiếp xúc qua da",
        "eye exposure": "Tiếp xúc qua mắt",
        "inhalation exposure": "Tiếp xúc qua đường hô hấp",
        "oral exposure": "Tiếp xúc qua đường miệng",
        "procedures of personal precautions": "Quy trình phòng ngừa cá nhân",
        "methods for cleaning up": "Biện pháp làm sạch",
        "exercise appropriate precautions to minimize direct contact with skin or eyes and prevent inhalation of dust": "Thực hiện các biện pháp phòng ngừa thích hợp để giảm thiểu tiếp xúc trực tiếp với da hoặc mắt và tránh hít phải bụi",
        "exercise appropriate precautions to minimize direct contact with skin and prevent inhalation of dust": "Thực hiện các biện pháp phòng ngừa thích hợp để giảm thiểu tiếp xúc trực tiếp với da và tránh hít phải bụi",
        "mix with inert material (eg. dry sand, vermiculite)and transfer to sealed container for disposal": "Trộn với vật liệu trơ (ví dụ: cát khô, vermiculit) và chuyển vào thùng kín để tiêu hủy",
        "mix with inert material (eg. dry sand, vermiculite) and transfer to sealed container for disposal": "Trộn với vật liệu trơ (ví dụ: cát khô, vermiculit) và chuyển vào thùng kín để tiêu hủy",
        "wear self-contained breathing apparatus and protective clothing to prevent contact with skin and eyes": "Mang thiết bị trợ thở độc lập và quần áo bảo hộ để ngăn ngừa tiếp xúc với da và mắt",
        "emits toxic fumes under fire conditions": "Phát sinh khói độc trong điều kiện hỏa hoạn",
        "exposure to sunlight may cause the product to harden within minutes": "Tiếp xúc với ánh sáng mặt trời có thể khiến sản phẩm đông cứng trong vòng vài phút",
        "occupational exposure limits for hazardous agents in the workplace": "Giới hạn phơi nhiễm nghề nghiệp đối với các yếu tố nguy hại tại nơi làm việc",
        "disclaimer the information presented here is based on our current best knowledge": "Tuyên bố từ chối trách nhiệm Thông tin trình bày ở đây dựa trên kiến thức tốt nhất hiện tại của chúng tôi",
        "the information presented here is based on our current best knowledge": "Thông tin trình bày ở đây dựa trên kiến thức tốt nhất hiện tại của chúng tôi",
        "vision new materials vietnam company limited": "CÔNG TY TNHH VISION NEW MATERIALS VIỆT NAM",
        "factory n2-a1, my phuoc 3 industrial park, thoi hoa ward, ben cat town, binh duong province, viet nam": "NHÀ XƯỞNG N2-A1, KCN MỸ PHƯỚC 3, PHƯỜNG THỚI HÒA, THỊ XÃ BẾN CÁT, TỈNH BÌNH DƯƠNG, VIỆT NAM",
        "classification and labels of dangerous chemical substances commonly used": "Phân loại và ghi nhãn các hóa chất nguy hại thông dụng",
        "general rule for preparation of precautionary label for industrial chemicals": "Quy tắc chung lập nhãn phòng ngừa cho hóa chất công nghiệp",
        "rule for storage of chemical dangers": "Quy tắc lưu trữ hóa chất nguy hiểm",
        "standard for limits of hazardous substances in wood coatings": "Tiêu chuẩn giới hạn chất nguy hại trong sơn gỗ",
    }

    # ─────────────────────────────────────────────────────────────────────────
    # 4. STANDARD MSDS PROPERTY & SECTION LABELS
    # ─────────────────────────────────────────────────────────────────────────
    FIELD_LABELS: Dict[str, str] = {
        "trade name": "Tên thương phẩm",
        "product name": "Tên sản phẩm",
        "product code": "Mã sản phẩm",
        "catalog codes": "Mã danh mục",
        "cas#": "Số CAS",
        "cas no.": "Số CAS",
        "- cas no.": "- Số CAS",
        "cas no": "Số CAS",
        "- cas no": "- Số CAS",
        "un no.": "Số UN",
        "- un no.": "- Số UN",
        "un no": "Số UN",
        "- un no": "- Số UN",
        "ems no.": "Số EMS",
        "- ems no.": "- Số EMS",
        "ems no": "Số EMS",
        "- ems no": "- Số EMS",
        "imo class": "Phân loại IMO",
        "- imo class": "- Phân loại IMO",
        "package group": "Nhóm đóng gói",
        "- package group": "- Nhóm đóng gói",
        "marine pollutants": "Chất gây ô nhiễm biển",
        "marine pollutants (yes / no)": "Chất gây ô nhiễm biển (có / không)",
        "packing method": "Phương pháp đóng gói",
        "einecs": "Số EINECS",
        "synonym": "Tên gọi khác",
        "synonyms": "Tên gọi khác",
        "chemical name": "Tên hóa học",
        "chemical formula": "Công thức hóa học",
        "contact information": "Thông tin liên hệ",
        "address": "Địa chỉ",
        "telephone": "Số điện thoại",
        "fax": "Số Fax",
        "email": "Email",
        "emergency telephone": "Điện thoại khẩn cấp",
        "ingredients": "Thành phần",
        "ingredient": "Thành phần",
        "weight (%)": "Tỷ lệ khối lượng (%)",
        "weight %": "Tỷ lệ khối lượng (%)",
        "form": "Dạng",
        "colour": "Màu sắc",
        "color": "Màu sắc",
        "odour": "Mùi",
        "odor": "Mùi",
        "odourless": "Không mùi",
        "odorless": "Không mùi",
        "powder": "Dạng bột",
        "white": "Trắng",
        "liquid": "Chất lỏng",
        "solid": "Chất rắn",
        "gas": "Chất khí",
        "ph-value": "Giá trị pH",
        "ph": "Giá trị pH",
        "ph-value at 20 °c": "Giá trị pH ở 20 °C",
        "melting point/melting range": "Điểm nóng chảy/khoảng nóng chảy",
        "melting point": "Điểm nóng chảy",
        "boiling point/boiling range": "Điểm sôi/khoảng sôi",
        "boiling point": "Điểm sôi",
        "flash point": "Điểm chớp cháy",
        "flammability (solid, gaseous)": "Tính dễ cháy (chất rắn, chất khí)",
        "ignition temperature": "Nhiệt độ bắt cháy",
        "decomposition temperature": "Nhiệt độ phân hủy",
        "self-igniting": "Tự bốc cháy",
        "danger of explosion": "Nguy cơ nổ",
        "explosion limits": "Giới hạn nổ",
        "lower": "Dưới",
        "upper": "Trên",
        "vapour pressure": "Áp suất hơi",
        "vapor pressure": "Áp suất hơi",
        "density": "Khối lượng riêng",
        "relative density": "Tỷ trọng tương đối",
        "vapour density": "Tỷ trọng hơi",
        "evaporation rate": "Tốc độ bay hơi",
        "solubility in / miscibility with water": "Độ tan trong / khả năng trộn lẫn với nước",
        "solubility in water": "Độ tan trong nước",
        "solubility": "Độ tan",
        "dynamic viscosity": "Độ nhớt động lực",
        "kinematic viscosity": "Độ nhớt động học",
        "viscosity": "Độ nhớt",
        "organic solvents": "Dung môi hữu cơ",
        "solids content": "Hàm lượng chất rắn",
        "stability": "Độ ổn định",
        "instability temperature": "Nhiệt độ không ổn định",
        "conditions of instability": "Điều kiện gây không ổn định",
        "incompatibility with various substances": "Tính không tương thích với các chất khác",
        "incompatible materials": "Vật liệu không tương thích",
        "corrosivity": "Tính ăn mòn",
        "polymerization": "Phản ứng trùng hợp",
        "routes of entry": "Đường xâm nhập",
        "toxicity to animals": "Độc tính đối với động vật",
        "chronic effects on humans": "Tác động mãn tính lên con người",
        "other toxic effects on humans": "Các tác động độc tính khác lên con người",
        "special remarks on toxicity to animals": "Ghi chú đặc biệt về độc tính trên động vật",
        "special remarks on chronic effects on humans": "Ghi chú đặc biệt về tác động mãn tính lên con người",
        "special remarks on other toxic effects on humans": "Ghi chú đặc biệt về các tác động độc tính khác lên con người",
        "ecotoxicity": "Độc tính sinh thái",
        "bod5 and cod": "BOD5 và COD",
        "products of biodegradation": "Sản phẩm phân hủy sinh học",
        "toxicity of the products of biodegradation": "Độc tính của sản phẩm phân hủy sinh học",
        "special remarks on the products of biodegradation": "Ghi chú đặc biệt về sản phẩm phân hủy sinh học",
        "waste disposal": "Xử lý chất thải",
        "dot classification": "Phân loại DOT",
        "identification": "Nhận dạng",
        "special provisions for transport": "Quy định đặc biệt khi vận chuyển",
        "federal and state regulations": "Quy định của liên bang và tiểu bang",
        "other regulations": "Các quy định khác",
        "other classifications": "Các phân loại khác",
        "references": "Tài liệu tham khảo",
        "other special considerations": "Các cân nhắc đặc biệt khác",
        "created": "Ngày lập",
        "last updated": "Cập nhật lần cuối",
        "printing date": "Ngày in",
        "revision date": "Ngày soát xét",
        "version": "Phiên bản",
        "not available": "Chưa có thông tin",
        "not applicable": "Không áp dụng",
        "none": "Không có",
        "unknown": "Không rõ",
        "see section": "Xem mục",
    }

    # ─────────────────────────────────────────────────────────────────────────
    # HELPER METHODS
    # ─────────────────────────────────────────────────────────────────────────

    @classmethod
    def get_ghs_section(cls, section_num: int) -> Optional[Dict[str, str]]:
        """Returns GHS section dict for 1 <= section_num <= 16."""
        return cls.GHS_SECTIONS.get(section_num)

    @classmethod
    def lookup_chemical(cls, name: str) -> Optional[str]:
        """Case-insensitive lookup for chemical substance name."""
        if not name:
            return None
        return cls.CHEMICAL_NAMES.get(name.strip().lower())

    @classmethod
    def lookup_hse(cls, phrase: str) -> Optional[str]:
        """Case-insensitive lookup for mandatory HSE phrase."""
        if not phrase:
            return None
        clean = phrase.strip().lower().rstrip(".:;,")
        return cls.HSE_PHRASES.get(clean)

    @classmethod
    def lookup_label(cls, label: str) -> Optional[str]:
        """Case-insensitive lookup for property field label."""
        if not label:
            return None
        clean = label.strip().lower().rstrip(".:;,")
        return cls.FIELD_LABELS.get(clean)
