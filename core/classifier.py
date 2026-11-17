"""
Phân hệ 2: Waterfall Classifier (Phân loại thác nước O(N))
Duyệt qua các paragraphs trong 1 vòng lặp O(N):
- Tầng 1: Media/Math (<w:drawing>, <w:pict>, <m:oMathPara>, <m:oMath>) -> Center, 0 indent.
- Tầng 2: Caption (Regex ^(Hình|Bảng|Sơ đồ|Biểu đồ)\s+\d+..., <150 chars) -> Center, Italic, 0 indent.
- Tầng 3: Heading -> Chuyển HSM & Whitelist & Auto-Numbering.
- Tầng 4: List items (<w:numPr>, bullet symbols •, ●, ○, ▪, ❖, ➢, -, +, a), b)) -> List Normalizer.
- Tầng 5: Body Text -> Strip tab/spaces thừa, Justify, first_line_indent = 1.27cm.
"""

import re
import unicodedata
import logging
from typing import List, Dict, Tuple, Optional, Any
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from config import Config
from core.run_manipulator import RunManipulator, int_to_roman
from core.list_normalizer import ListNormalizer
from core.hsm import HierarchicalStateMachine, AIFallbackClassifier
from core.numbering_parser import NumberingParser

logger = logging.getLogger(__name__)

# Math XML Namespace
MATH_NS = {'m': 'http://schemas.openxmlformats.org/officeDocument/2006/math'}

def roman_to_int(s: str) -> int:
    """Chuyển đổi số La Mã sang số nguyên (I -> 1, II -> 2, III -> 3, IV -> 4,...)"""
    roman_map = {'I': 1, 'V': 5, 'X': 10, 'L': 50, 'C': 100, 'D': 500, 'M': 1000}
    s = s.upper().strip()
    total = 0
    prev = 0
    for char in reversed(s):
        val = roman_map.get(char, 0)
        if val >= prev:
            total += val
        else:
            total -= val
        prev = val
    return total if total > 0 else 1

# Regex nhận diện Caption (hỗ trợ cả caption không có số thứ tự như "Hình : ...")
CAPTION_PATTERN = re.compile(
    r'^(hình|bảng|sơ đồ|biểu đồ|đồ thị|ảnh)(\s*[:\.\-\–—]\s*.+|\s+[0-9IVXLCDM]+[a-zA-Z0-9\.\-–—]*(\s*[:\.\-\–—\s].*|$))',
    re.IGNORECASE
)

# Regex nhận diện Heading số thông thường (1., 1.1, 1.1.1, 1.1.1.1) và đề mục dở dang (3.1.)
NUMBERED_HEADING_PATTERN = re.compile(
    r'^(\d+(\.\d+)*)\.?(\s+(.*))?$'
)

# Regex nhận diện đề mục số La Mã: I. PHÂN CÔNG CÔNG VIỆC
ROMAN_HEADING_PATTERN = re.compile(r'^[IVXLCDM]+\.\s+.+$')

# Regex nhận diện Chương / Phần / Bài
CHAPTER_HEADING_PATTERN = re.compile(
    r'^(chương|phần|bài)\s+([0-9IVXLCDM]+)[\s:.\-–—]*(.*)$',
    re.IGNORECASE
)

# Regex nhận diện List bullet hoặc ký tự đánh dấu danh sách
LIST_BULLET_PATTERN = re.compile(
    r'^[\s\t]*([•●○▪❖➢★–—\-\+\*►]|\([a-z0-9]\)|[a-z]\)|\d+\))\s+',
    re.IGNORECASE
)

class ParagraphType:
    MEDIA_MATH = "MEDIA_MATH"
    CAPTION = "CAPTION"
    HEADING_1 = "HEADING_1"
    HEADING_2 = "HEADING_2"
    HEADING_3 = "HEADING_3"
    HEADING_4 = "HEADING_4"
    HEADING_UNNUMBERED = "HEADING_UNNUMBERED"
    HEADING_NON_TOC = "HEADING_NON_TOC"
    LIST_ITEM = "LIST_ITEM"
    BODY = "BODY"
    AMBIGUOUS = "AMBIGUOUS"

class WaterfallClassifier:
    def __init__(
        self,
        doc,
        numbering_parser: NumberingParser,
        ai_classifier: Optional[AIFallbackClassifier] = None,
        normalize_chapter_to_roman: bool = False
    ):
        self.doc = doc
        self.numbering_parser = numbering_parser
        self.hsm = HierarchicalStateMachine()
        self.ai_classifier = ai_classifier or AIFallbackClassifier()
        self.normalize_chapter_to_roman = normalize_chapter_to_roman

    @staticmethod
    def normalize_chapter_to_roman(paragraph) -> bool:
        """
        Chuẩn hóa số Chương từ số Ả Rập sang số La Mã trên paragraph (thao tác trên Run).
        Ví dụ: 'CHƯƠNG 1: ...' -> 'CHƯƠNG I: ...'
        """
        return RunManipulator.convert_chapter_number_to_roman(paragraph)

    def is_tier1_media_or_math(self, paragraph) -> bool:
        """
        Tầng 1: Kiểm tra xem đoạn văn có chứa Media/Drawing hoặc Công thức Toán học đứng độc lập không
        """
        elem = paragraph._element

        # Kiểm tra Media: w:drawing, w:pict, w:object
        drawings = elem.xpath('.//w:drawing | .//w:pict | .//w:object')
        if drawings:
            # Nếu có hình ảnh và đoạn văn rất ngắn (< 30 ký tự text đi kèm)
            if len(paragraph.text.strip()) < 35:
                return True

        # Kiểm tra Math: m:oMathPara (khối công thức toán độc lập)
        math_paras = elem.xpath('.//m:oMathPara')
        if math_paras:
            return True

        # Kiểm tra m:oMath khi nội dung chữ hầu như chỉ là công thức
        maths = elem.xpath('.//m:oMath')
        if maths and len(paragraph.text.strip()) < 15:
            return True

        return False

    def is_tier2_caption(self, paragraph) -> bool:
        """
        Tầng 2: Kiểm tra Caption hình ảnh, bảng biểu (hỗ trợ cả style 'hinhanh', 'bang', 'Caption')
        """
        text = unicodedata.normalize('NFC', paragraph.text.strip())
        if not text or len(text) > 160:
            return False

        style_name = (paragraph.style.name if paragraph.style else "").lower()
        if style_name in ['caption', 'hinhanh', 'bang']:
            return True

        return bool(CAPTION_PATTERN.match(text))

    def evaluate_tier3_heading(self, paragraph) -> Tuple[bool, Optional[str], Optional[int]]:
        """
        Tầng 3: Kiểm tra xem đoạn văn có phải Heading hay không.
        Trả về: (is_heading: bool, heading_type: Optional[str], heading_level: Optional[int])
        """
        text = unicodedata.normalize('NFC', paragraph.text.strip())
        if not text:
            return False, None, None

        upper_text = text.upper()
        # Loại bỏ dấu kết thúc để so khớp whitelist nếu có
        clean_upper = upper_text.rstrip(':. ')

        # 1. Kiểm tra Whitelist A: Heading 1 Unnumbered (Vào TOC)
        if clean_upper in Config.TOC_UNNUMBERED_WHITELIST:
            return True, ParagraphType.HEADING_UNNUMBERED, 1

        # 2. Kiểm tra Whitelist B: Non-TOC Heading (Không vào TOC)
        if clean_upper in Config.NON_TOC_WHITELIST:
            return True, ParagraphType.HEADING_NON_TOC, 1

        # 3. Kiểm tra Chương / Phần / Bài (Chương 1, Chương I,...) -> Luôn là Heading 1
        chap_match = CHAPTER_HEADING_PATTERN.match(text)
        if chap_match:
            # Cập nhật HSM
            chap_num_str = chap_match.group(2)
            try:
                chap_num = int(chap_num_str)
            except ValueError:
                chap_num = roman_to_int(chap_num_str)
            self.hsm.validate_and_update([chap_num], text, is_bold=True)
            return True, ParagraphType.HEADING_1, 1

        # 3b. Kiểm tra Đề mục số La Mã (I. PHÂN CÔNG CÔNG VIỆC) -> Luôn là Heading 1
        if ROMAN_HEADING_PATTERN.match(text):
            roman_str = text.split('.')[0].strip()
            val = roman_to_int(roman_str)
            self.hsm.validate_and_update([val], text, is_bold=True)
            return True, ParagraphType.HEADING_1, 1

        # 4. Kiểm tra Auto-Numbering từ numbering.xml
        num_id, ilvl = self.numbering_parser.extract_num_pr(paragraph)
        if num_id is not None:
            num_info = self.numbering_parser.get_info(num_id, ilvl)
            style_name = (paragraph.style.name if paragraph.style else "").lower()
            p_style = (num_info.get("p_style") or "").lower()
            is_bold = any(r.bold for r in paragraph.runs if r.text.strip())
            is_heading_style = "heading" in style_name or "heading" in p_style

            if num_info["is_decimal"] and num_info["heading_level"] is not None and (is_heading_style or is_bold):
                h_lvl = num_info["heading_level"]
                dummy_path = [1] * h_lvl
                valid, valid_lvl = self.hsm.validate_and_update(dummy_path, text, is_bold=is_bold)
                if valid:
                    type_str = f"HEADING_{min(valid_lvl, 4)}"
                    return True, getattr(ParagraphType, type_str, ParagraphType.HEADING_4), valid_lvl

        # 5. Kiểm tra Regex số phân cấp thủ công: "1.", "1.1", "1.1.1", "1.1.1.1", đề mục dở dang "3.1."
        num_match = NUMBERED_HEADING_PATTERN.match(text)
        if num_match:
            num_part = num_match.group(1)
            title_part = (num_match.group(4) or "").strip()
            try:
                candidate_path = [int(p) for p in num_part.split('.') if p]
            except ValueError:
                candidate_path = []

            # Đề mục dở dang cấp 1 chỉ có số ("1.") mà không có tiêu đề thì không phải heading
            if len(candidate_path) == 1 and not title_part:
                candidate_path = []

            if candidate_path:
                is_bold = any(r.bold for r in paragraph.runs if r.text.strip())
                valid, valid_lvl = self.hsm.validate_and_update(candidate_path, text, is_bold=is_bold)
                if valid:
                    type_str = f"HEADING_{min(valid_lvl, 4)}"
                    return True, getattr(ParagraphType, type_str, ParagraphType.HEADING_4), valid_lvl

        # 6. Kiểm tra Style gốc có tên chứa "heading" sẵn không
        style_name = (paragraph.style.name if paragraph.style else "").lower()
        if "heading" in style_name:
            for lvl in range(1, 5):
                if str(lvl) in style_name:
                    return True, getattr(ParagraphType, f"HEADING_{lvl}"), lvl

        return False, None, None

    def is_tier4_list(self, paragraph) -> Tuple[bool, int]:
        """
        Tầng 4: Kiểm tra danh sách (List item)
        Trả về: (is_list: bool, level: int)
        """
        # Kiểm tra w:numPr (bất kỳ đoạn nào có numId mà không phải heading thì là list item)
        num_id, ilvl = self.numbering_parser.extract_num_pr(paragraph)
        if num_id is not None:
            return True, ListNormalizer.determine_list_level(paragraph, ilvl)

        # Kiểm tra ký tự text bắt đầu bằng ký hiệu bullet hoặc marker a), b), 1)
        text = unicodedata.normalize('NFC', paragraph.text)
        if LIST_BULLET_PATTERN.match(text):
            return True, ListNormalizer.determine_list_level(paragraph, None)

        return False, 1

    def is_ambiguous_candidate(self, paragraph) -> bool:
        """
        Kiểm tra xem paragraph có phải là ứng viên nghi vấn cần chuyển cho AI Fallback không:
        - Văn bản ngắn (< 110 ký tự)
        - Có in đậm (is_bold = True)
        - Không có số mục rõ ràng
        - Không kết thúc bằng dấu chấm câu đầy đủ
        """
        text = unicodedata.normalize('NFC', paragraph.text.strip())
        if not text or len(text) > 110 or len(text) < 4:
            return False

        # Nếu kết thúc bằng dấu chấm câu bình thường hoặc dấu hai chấm
        if text.endswith((".", ":", ";", ",")):
            return False

        is_bold = any(r.bold for r in paragraph.runs if r.text.strip())
        return is_bold
