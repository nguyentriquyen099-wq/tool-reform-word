"""
Cấu hình chuẩn hóa báo cáo Word (.docx)
Hỗ trợ Nghị định 30/2020/NĐ-CP và Chuẩn Báo cáo Luận văn / Đồ án / Kỹ thuật
"""

import os
from pathlib import Path
from dotenv import load_dotenv

# Tải cấu hình từ .env
load_dotenv(override=True)

def get_config_value(key: str, default: str = "") -> str:
    """
    Đọc cấu hình theo thứ tự ưu tiên:
    1. st.secrets (dành cho Streamlit Cloud hoặc .streamlit/secrets.toml)
       Hỗ trợ cả UPPERCASE, lowercase, và cấu hình lồng trong [payment] / [vietqr]
    2. os.getenv (từ file .env hoặc biến môi trường hệ thống)
    3. Giá trị mặc định (default)
    """
    try:
        import streamlit as st
        if hasattr(st, "secrets"):
            keys_to_try = [key, key.lower(), key.upper()]
            # 1.1 Kiểm tra top-level keys
            for k in keys_to_try:
                if k in st.secrets:
                    val = st.secrets[k]
                    if val is not None:
                        return str(val).strip()
            # 1.2 Kiểm tra nested section [payment] hoặc [vietqr] hoặc [sepay]
            for section_name in ["payment", "PAYMENT", "vietqr", "VIETQR", "sepay", "SEPAY"]:
                if section_name in st.secrets:
                    sec = st.secrets[section_name]
                    if isinstance(sec, dict) or hasattr(sec, "__getitem__"):
                        for k in keys_to_try:
                            if k in sec:
                                val = sec[k]
                                if val is not None:
                                    return str(val).strip()
    except Exception:
        pass

    for k in [key, key.lower(), key.upper()]:
        val = os.getenv(k)
        if val is not None:
            return str(val).strip()

    return default

def parse_payment_amount(val_str, default: int = 5000) -> int:
    """Chuyển đổi chuỗi số tiền sang số nguyên an toàn, hỗ trợ dấu phẩy và chấm nghìn"""
    if val_str is None:
        return default
    if isinstance(val_str, (int, float)):
        return max(0, int(val_str))
    cleaned = str(val_str).strip().replace(" ", "")
    if not cleaned:
        return default
    try:
        # Nếu có dấu phẩy ngăn cách hàng nghìn (vd 10,000)
        if "," in cleaned and "." not in cleaned:
            cleaned = cleaned.replace(",", "")
        # Nếu có dạng 10.000 (dấu chấm ngăn cách nghìn kiểu Việt Nam)
        elif cleaned.count(".") == 1 and len(cleaned.split(".")[1]) == 3 and "," not in cleaned:
            cleaned = cleaned.replace(".", "")
        return max(0, int(float(cleaned)))
    except (ValueError, TypeError):
        return default

# Danh mục các ngân hàng phổ biến tại Việt Nam (chuẩn mã Napas / VietQR)
POPULAR_BANKS = {
    "MB": "MBBank - Ngân hàng Quân Đội",
    "VCB": "Vietcombank - Ngân hàng Ngoại Thương",
    "TCB": "Techcombank - Ngân hàng Kỹ Thương",
    "BIDV": "BIDV - Ngân hàng Đầu tư & Phát triển",
    "ICB": "VietinBank - Ngân hàng Công Thương",
    "VBA": "Agribank - Ngân hàng Nông nghiệp & PT Nông thôn",
    "ACB": "ACB - Ngân hàng Á Châu",
    "VPB": "VPBank - Ngân hàng Việt Nam Thịnh Vượng",
    "TPB": "TPBank - Ngân hàng Tiên Phong",
    "STB": "Sacombank - Ngân hàng Sài Gòn Thương Tín",
    "VIB": "VIB - Ngân hàng Quốc tế",
    "SHB": "SHB - Ngân hàng Sài Gòn Hà Nội",
    "HDB": "HDBank - Ngân hàng Phát triển TP.HCM",
    "LPB": "LPBank - Ngân hàng Lộc Phát",
    "MSB": "MSB - Ngân hàng Hàng Hải",
    "OCB": "OCB - Ngân hàng Phương Đông",
    "SEAB": "SeABank - Ngân hàng Đông Nam Á",
    "NAB": "Nam A Bank - Ngân hàng Nam Á",
    "BAB": "Bac A Bank - Ngân hàng Bắc Á",
    "BVB": "BVBank - Ngân hàng Bản Việt",
    "PVCB": "PVcomBank - Ngân hàng Đại Chúng",
    "NCB": "NCB - Ngân hàng Quốc Dân",
    "VAB": "VietABank - Ngân hàng Việt Á",
    "VIETBANK": "Vietbank - Ngân hàng Việt Nam Thương Tín",
    "SCB": "SCB - Ngân hàng Sài Gòn",
    "PGB": "PGBank - Ngân hàng Thịnh Vượng và Phát Triển",
    "CAKE": "CAKE by VPBank",
    "TIMO": "Timo by BVBank",
}

# Ánh xạ tên gọi / bí danh thường dùng sang mã Napas chuẩn
BANK_ALIASES = {
    "AGRIBANK": "VBA",
    "VIETCOMBANK": "VCB",
    "MBBANK": "MB",
    "TECHCOMBANK": "TCB",
    "VIETINBANK": "ICB",
    "SACOMBANK": "STB",
    "TIENPHONGBANK": "TPB",
    "VPBANK": "VPB",
    "ACBBANK": "ACB",
    "VCCB": "BVB",
    "NAMABANK": "NAB",
    "BACABANK": "BAB",
    "PVCOMBANK": "PVCB",
}

def normalize_bank_id(raw_bank_id: str) -> str:
    """Chuẩn hóa mã định danh ngân hàng sang chuẩn Napas/VietQR"""
    if not raw_bank_id:
        return ""
    clean = str(raw_bank_id).strip().replace(" ", "").upper()
    return BANK_ALIASES.get(clean, clean)

class Config:
    # DeepSeek AI Configuration
    DEEPSEEK_API_KEY = get_config_value("DEEPSEEK_API_KEY", "").strip()
    DEEPSEEK_BASE_URL = get_config_value("DEEPSEEK_BASE_URL", "https://api.deepseek.com").strip()
    DEEPSEEK_MODEL = get_config_value("DEEPSEEK_MODEL", "deepseek-chat").strip()

    # Cấu hình Cổng Thanh toán QR Banking (VietQR & Local QR)
    ENABLE_PAYMENT_GATEWAY = get_config_value("ENABLE_PAYMENT_GATEWAY", "True").lower() in ("true", "1", "yes")
    BANK_ID = normalize_bank_id(get_config_value("BANK_ID", "MB"))
    ACCOUNT_NO = get_config_value("ACCOUNT_NO", "").strip()
    ACCOUNT_NAME = get_config_value("ACCOUNT_NAME", "").strip()
    PAYMENT_AMOUNT = parse_payment_amount(get_config_value("PAYMENT_AMOUNT", "5000"), 5000)

    # Cấu hình Thanh toán Tự động SePay (Cấp độ 2 - Tự động check biến động số dư)
    PAYMENT_MODE = get_config_value("PAYMENT_MODE", "auto").strip().lower()
    SEPAY_API_TOKEN = get_config_value("SEPAY_API_TOKEN", "").strip()
    _init_chk_int = get_config_value("PAYMENT_CHECK_INTERVAL_SECONDS", "3").strip()
    PAYMENT_CHECK_INTERVAL_SECONDS = max(1, int(_init_chk_int if _init_chk_int.isdigit() else 3))
    SEPAY_TEST_MODE = get_config_value("SEPAY_TEST_MODE", "False").lower() in ("true", "1", "yes")
    SEPAY_API_URL = get_config_value("SEPAY_API_URL", "https://my.sepay.vn/userapi/transactions/list").strip()

    # Font và kiểu dáng mặc định (Chuẩn hóa 100% theo Mẫu form)
    FONT_NAME = get_config_value("FONT_NAME", "Times New Roman")
    BODY_FONT_SIZE = float(get_config_value("BODY_FONT_SIZE", "14"))
    HEADING_1_FONT_SIZE = float(get_config_value("HEADING_1_FONT_SIZE", "16"))
    HEADING_2_FONT_SIZE = float(get_config_value("HEADING_2_FONT_SIZE", "14"))
    HEADING_3_FONT_SIZE = float(get_config_value("HEADING_3_FONT_SIZE", "14"))
    HEADING_4_FONT_SIZE = float(get_config_value("HEADING_4_FONT_SIZE", "14"))
    CAPTION_FONT_SIZE = float(get_config_value("CAPTION_FONT_SIZE", "14.0"))
    TABLE_FONT_SIZE = float(get_config_value("TABLE_FONT_SIZE", "14.0"))

    # Căn lề trang (Standard Page Margins theo Mẫu form chuẩn: Top 2.5cm, Bottom 2.0cm, Left 3.5cm, Right 2.0cm)
    MARGIN_TOP_CM = float(get_config_value("MARGIN_TOP_CM", "2.5"))
    MARGIN_BOTTOM_CM = float(get_config_value("MARGIN_BOTTOM_CM", "2.0"))
    MARGIN_LEFT_CM = float(get_config_value("MARGIN_LEFT_CM", "3.5"))
    MARGIN_RIGHT_CM = float(get_config_value("MARGIN_RIGHT_CM", "2.0"))

    # Thụt lề đầu dòng Body Text (cm) - Mẫu form chuẩn 1.0cm (567 dxa)
    FIRST_LINE_INDENT_CM = float(get_config_value("FIRST_LINE_INDENT_CM", "1.0"))

    # Thụt lề phân cấp cho Heading (cm mỗi cấp) - Mẫu form chuẩn kịch lề trái 0cm cho mọi cấp
    HEADING_INDENT_PER_LEVEL_CM = float(get_config_value("HEADING_INDENT_PER_LEVEL_CM", "0.0"))

    # Khoảng cách dòng & đoạn cho Body Text (Normal) - Mẫu form chuẩn 1.5 lines (360 dxa)
    LINE_SPACING = float(get_config_value("LINE_SPACING", "1.5"))
    SPACE_BEFORE_PT = float(get_config_value("SPACE_BEFORE_PT", "0"))
    SPACE_AFTER_PT = float(get_config_value("SPACE_AFTER_PT", "6"))

    # Khoảng cách đoạn cho Headings
    HEADING_1_SPACE_BEFORE_PT = float(get_config_value("HEADING_1_SPACE_BEFORE_PT", "12"))
    HEADING_1_SPACE_AFTER_PT = float(get_config_value("HEADING_1_SPACE_AFTER_PT", "6"))
    HEADING_SPACE_BEFORE_PT = float(get_config_value("HEADING_SPACE_BEFORE_PT", "6"))
    HEADING_SPACE_AFTER_PT = float(get_config_value("HEADING_SPACE_AFTER_PT", "3"))

    # Whitelist các tiêu đề không đánh số (Unnumbered Headings)
    # Nhóm A: Hiển thị trong Mục lục (TOC) - outlineLevel = 0
    TOC_UNNUMBERED_WHITELIST = {
        "LỜI MỞ ĐẦU",
        "LỜI NÓI ĐẦU",
        "MỞ ĐẦU",
        "KẾT LUẬN",
        "KẾT LUẬN VÀ KIẾN NGHỊ",
        "TÀI LIỆU THAM KHẢO",
        "PHỤ LỤC",
        "DANH MỤC TÀI LIỆU THAM KHẢO"
    }

    # Nhóm B: Định dạng giống Heading 1 nhưng KHÔNG đưa vào Mục lục (outlineLevel = None)
    NON_TOC_WHITELIST = {
        "LỜI CẢM ƠN",
        "MỤC LỤC",
        "DANH MỤC HÌNH",
        "DANH MỤC HÌNH ẢNH",
        "DANH MỤC HÌNH VẼ",
        "DANH MỤC BẢNG",
        "DANH MỤC BẢNG BIỂU",
        "DANH MỤC TỪ VIẾT TẮT",
        "DANH MỤC CÁC TỪ VIẾT TẮT",
        "DANH MỤC VIẾT TẮT",
        "DANH MỤC KÝ HIỆU VÀ CHỮ VIẾT TẮT",
        "TÓM TẮT",
        "TÓM TẮT NỘI DUNG",
        "ABSTRACT"
    }

    # Các từ khóa nhận diện ranh giới kết thúc Trang Bìa (Document Zoning)
    COVER_BOUNDARY_KEYWORDS = [
        "LỜI CẢM ƠN",
        "LỜI NÓI ĐẦU",
        "LỜI MỞ ĐẦU",
        "MỞ ĐẦU",
        "MỤC LỤC",
        "CHƯƠNG 1",
        "CHƯƠNG I",
        "PHẦN 1",
        "PHẦN I",
        "DANH MỤC HÌNH",
        "DANH MỤC HÌNH VẼ",
        "DANH MỤC BẢNG",
        "DANH MỤC BẢNG BIỂU",
        "DANH MỤC TỪ VIẾT TẮT",
        "DANH MỤC CÁC TỪ VIẾT TẮT",
        "TÓM TẮT"
    ]

    # Các ký hiệu Bullet cần chuẩn hóa
    BULLET_SYMBOLS = ["•", "●", "○", "▪", "❖", "➢", "", "", "★", "–", "—", "*", "►", ""]

    # Chuẩn hóa số Chương sang số La Mã - Mặc định False theo Form mẫu (dùng CHƯƠNG 1, 2, 3 đồng bộ 1.1, 1.2)
    NORMALIZE_CHAPTER_TO_ROMAN = os.getenv("NORMALIZE_CHAPTER_TO_ROMAN", "False").lower() in ("true", "1", "yes")

    # Chuẩn hóa số Chương sang số La Mã - Mặc định False theo Form mẫu (dùng CHƯƠNG 1, 2, 3 đồng bộ 1.1, 1.2)
    NORMALIZE_CHAPTER_TO_ROMAN = get_config_value("NORMALIZE_CHAPTER_TO_ROMAN", "False").lower() in ("true", "1", "yes")

    # Chuẩn hóa chữ hoa cho Heading cấp con (Heading 2, 3...) sang dạng Sentence Case chuẩn
    NORMALIZE_HEADING_CAPS = get_config_value("NORMALIZE_HEADING_CAPS", "True").lower() in ("true", "1", "yes")

    # Tên Style tùy biến
    STYLE_HEADING_1_UNNUMBERED = "Heading 1 Unnumbered"
    STYLE_NON_TOC_HEADING = "Non-TOC Heading"

    @classmethod
    def get_bank_display_name(cls, bank_id: str = None) -> str:
        """Lấy tên hiển thị chi tiết của ngân hàng"""
        raw = bank_id if bank_id is not None else cls.BANK_ID
        b_id = normalize_bank_id(raw)
        if not b_id:
            return "Ngân hàng"
        return POPULAR_BANKS.get(b_id, f"Ngân hàng {b_id}")

    @classmethod
    def find_local_qr_image(cls) -> str | None:
        """
        Tìm kiếm file ảnh QR có sẵn trong thư mục gốc dự án (qr.png, qr.jpg, ...)
        Hỗ trợ Cơ chế 2: Hiển thị file ảnh QR cá nhân của người dùng.
        Chỉ chấp nhận file tồn tại và có dung lượng > 0 bytes để chống lỗi file rỗng.
        """
        base_dir = Path(__file__).resolve().parent
        candidate_names = [
            "qr.png", "qr.jpg", "qr.jpeg",
            "qr_code.png", "qr_code.jpg", "qr_code.jpeg",
            "qrcode.png", "qrcode.jpg", "qrcode.jpeg",
            "bank_qr.png", "bank_qr.jpg",
            "QR.png", "QR.jpg", "QR.PNG", "QR.jpeg",
            "QR_CODE.png", "QR_CODE.jpg"
        ]
        for fname in candidate_names:
            p = base_dir / fname
            if p.exists() and p.is_file() and p.stat().st_size > 0:
                return str(p)
        return None

    @classmethod
    def is_vietqr_configured(cls, bank_id: str = None, account_no: str = None) -> bool:
        """Kiểm tra thông tin tài khoản VietQR đã được cấu hình hay chưa an toàn với None/rỗng"""
        b_raw = bank_id if bank_id is not None else cls.BANK_ID
        b_id = normalize_bank_id(b_raw)
        acc_raw = account_no if account_no is not None else cls.ACCOUNT_NO
        acc = str(acc_raw or "").strip().replace(" ", "")
        return bool(b_id and acc)

    @classmethod
    def get_vietqr_url(
        cls,
        bank_id: str = None,
        account_no: str = None,
        account_name: str = None,
        amount: int = None,
        add_info: str = None
    ) -> str:
        """
        Tạo mã VietQR chuẩn quốc tế với số tiền và nội dung chuyển khoản tự sinh qua URL API VietQR miễn phí:
        https://img.vietqr.io/image/{BANK_ID}-{ACCOUNT_NO}-compact2.png?amount={AMOUNT}&addInfo={ADD_INFO}&accountName={ACCOUNT_NAME}
        """
        import urllib.parse
        b_raw = bank_id if bank_id is not None else cls.BANK_ID
        b_id = normalize_bank_id(b_raw)
        acc_raw = account_no if account_no is not None else cls.ACCOUNT_NO
        acc = str(acc_raw or "").strip().replace(" ", "")
        name = str(account_name if account_name is not None else cls.ACCOUNT_NAME or "").strip()
        amt = parse_payment_amount(amount if amount is not None else cls.PAYMENT_AMOUNT, 0)
        info = str(add_info if add_info is not None else "CH BAO CAO").strip()

        params = []
        if amt > 0:
            params.append(f"amount={amt}")
        if info:
            params.append(f"addInfo={urllib.parse.quote(info)}")
        if name:
            params.append(f"accountName={urllib.parse.quote(name)}")

        query = f"?{'&'.join(params)}" if params else ""
        return f"https://img.vietqr.io/image/{b_id}-{acc}-compact2.png{query}"

    @classmethod
    def reload(cls, load_env: bool = True):
        """Tải lại toàn bộ cấu hình từ file .env hoặc biến môi trường / st.secrets"""
        if load_env:
            load_dotenv(override=True)
        cls.DEEPSEEK_API_KEY = get_config_value("DEEPSEEK_API_KEY", "").strip()
        cls.DEEPSEEK_BASE_URL = get_config_value("DEEPSEEK_BASE_URL", "https://api.deepseek.com").strip()
        cls.DEEPSEEK_MODEL = get_config_value("DEEPSEEK_MODEL", "deepseek-chat").strip()

        cls.ENABLE_PAYMENT_GATEWAY = get_config_value("ENABLE_PAYMENT_GATEWAY", "True").lower() in ("true", "1", "yes")
        cls.BANK_ID = normalize_bank_id(get_config_value("BANK_ID", "MB"))
        cls.ACCOUNT_NO = get_config_value("ACCOUNT_NO", "").strip()
        cls.ACCOUNT_NAME = get_config_value("ACCOUNT_NAME", "").strip()
        cls.PAYMENT_AMOUNT = parse_payment_amount(get_config_value("PAYMENT_AMOUNT", "5000"), 5000)

        cls.PAYMENT_MODE = get_config_value("PAYMENT_MODE", "auto").strip().lower()
        cls.SEPAY_API_TOKEN = get_config_value("SEPAY_API_TOKEN", "").strip()
        _chk_int = get_config_value("PAYMENT_CHECK_INTERVAL_SECONDS", "3").strip()
        cls.PAYMENT_CHECK_INTERVAL_SECONDS = max(1, int(_chk_int if _chk_int.isdigit() else 3))
        cls.SEPAY_TEST_MODE = get_config_value("SEPAY_TEST_MODE", "False").lower() in ("true", "1", "yes")
        cls.SEPAY_API_URL = get_config_value("SEPAY_API_URL", "https://my.sepay.vn/userapi/transactions/list").strip()

        cls.FONT_NAME = get_config_value("FONT_NAME", "Times New Roman")
        cls.BODY_FONT_SIZE = float(get_config_value("BODY_FONT_SIZE", "14"))
        cls.HEADING_1_FONT_SIZE = float(get_config_value("HEADING_1_FONT_SIZE", "16"))
        cls.HEADING_2_FONT_SIZE = float(get_config_value("HEADING_2_FONT_SIZE", "14"))
        cls.HEADING_3_FONT_SIZE = float(get_config_value("HEADING_3_FONT_SIZE", "14"))
        cls.HEADING_4_FONT_SIZE = float(get_config_value("HEADING_4_FONT_SIZE", "14"))
        cls.CAPTION_FONT_SIZE = float(get_config_value("CAPTION_FONT_SIZE", "14.0"))
        cls.TABLE_FONT_SIZE = float(get_config_value("TABLE_FONT_SIZE", "14.0"))

        cls.MARGIN_TOP_CM = float(get_config_value("MARGIN_TOP_CM", "2.5"))
        cls.MARGIN_BOTTOM_CM = float(get_config_value("MARGIN_BOTTOM_CM", "2.0"))
        cls.MARGIN_LEFT_CM = float(get_config_value("MARGIN_LEFT_CM", "3.5"))
        cls.MARGIN_RIGHT_CM = float(get_config_value("MARGIN_RIGHT_CM", "2.0"))

        cls.FIRST_LINE_INDENT_CM = float(get_config_value("FIRST_LINE_INDENT_CM", "1.0"))
        cls.HEADING_INDENT_PER_LEVEL_CM = float(get_config_value("HEADING_INDENT_PER_LEVEL_CM", "0.0"))

        cls.LINE_SPACING = float(get_config_value("LINE_SPACING", "1.5"))
        cls.SPACE_BEFORE_PT = float(get_config_value("SPACE_BEFORE_PT", "0"))
        cls.SPACE_AFTER_PT = float(get_config_value("SPACE_AFTER_PT", "6"))

        cls.HEADING_1_SPACE_BEFORE_PT = float(get_config_value("HEADING_1_SPACE_BEFORE_PT", "12"))
        cls.HEADING_1_SPACE_AFTER_PT = float(get_config_value("HEADING_1_SPACE_AFTER_PT", "6"))
        cls.HEADING_SPACE_BEFORE_PT = float(get_config_value("HEADING_SPACE_BEFORE_PT", "6"))
        cls.HEADING_SPACE_AFTER_PT = float(get_config_value("HEADING_SPACE_AFTER_PT", "3"))

        cls.NORMALIZE_CHAPTER_TO_ROMAN = get_config_value("NORMALIZE_CHAPTER_TO_ROMAN", "False").lower() in ("true", "1", "yes")
        cls.NORMALIZE_HEADING_CAPS = get_config_value("NORMALIZE_HEADING_CAPS", "True").lower() in ("true", "1", "yes")
