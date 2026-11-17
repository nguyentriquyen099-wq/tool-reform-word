"""
Phân hệ 4: List Normalizer (Chuẩn hóa danh sách)
- Xóa bullet rác, ký tự Symbol ẩn (<w:sym>)
- Xóa thẻ <w:numPr> để tránh xung đột bullet gốc của Word
- Quy đổi về dấu chuẩn:
  + Cấp 1: "- " (first_line_indent = 1.27cm, left_indent = 0)
  + Cấp 2: "+ " (first_line_indent = 2.0cm, left_indent = 0, thụt lề sâu hơn)
- Sử dụng RunManipulator để bảo tồn tuyệt đối bold/italic và không mất dữ liệu
"""

import re
import logging
from typing import Optional
from docx.shared import Cm, Pt
from docx.oxml.ns import qn
from config import Config
from core.run_manipulator import RunManipulator

logger = logging.getLogger(__name__)

# Pattern nhận diện các danh sách con dạng chữ cái hoặc dấu cộng: a), b), +, o, -
LEVEL_2_TRIGGER_PATTERN = re.compile(r'^(\+|\*|○|o|\([a-z0-9]\)|[a-z]\))\s*', re.IGNORECASE)

class ListNormalizer:
    @staticmethod
    def normalize_list_item(paragraph, level: int = 1):
        """
        Chuẩn hóa một đoạn văn danh sách:
        1. Xóa ký tự bullet rác & symbol ẩn ở các Run đầu
        2. Gỡ bỏ thẻ w:numPr trong XML
        3. Gán dấu đầu dòng chuẩn ("- " cho cấp 1, "+ " cho cấp 2)
        4. Thiết lập thụt lề chuẩn:
           - Cấp 1: first_line_indent = 1.27cm, left_indent = 0
           - Cấp 2: first_line_indent = 2.0cm, left_indent = 0
        5. Đảm bảo giãn dòng gọn gàng, reset thừa style
        """
        # 1. Làm sạch các ký tự bullet rác ở Run đầu tiên
        RunManipulator.clean_leading_junk_and_symbols(paragraph, remove_bullets=True)

        # 2. Gỡ bỏ thẻ w:numPr khỏi paragraph để không bị double-bullet trong Word
        pPr = paragraph._element.find(qn('w:pPr'))
        if pPr is not None:
            numPr = pPr.find(qn('w:numPr'))
            if numPr is not None:
                pPr.remove(numPr)

        # 3. Xác định tiền tố marker theo cấp độ
        marker = "- " if level == 1 else "+ "

        # 4. Thêm marker vào đầu Run đầu tiên
        RunManipulator.prepend_text_to_paragraph(paragraph, marker)

        # Đặt paragraph style về Normal để không bị ảnh hưởng style bullet cũ của Word
        try:
            paragraph.style = "Normal"
        except Exception:
            try:
                paragraph.style = paragraph.part.document.styles['Normal']
            except Exception:
                pass

        # 5. Căn chỉnh thụt lề (Paragraph Format) - Dùng Hanging Indent chống thụt ngược dòng 2
        pf = paragraph.paragraph_format
        if level == 1:
            pf.left_indent = Cm(Config.FIRST_LINE_INDENT_CM)
            pf.first_line_indent = -Cm(0.63)
        else:
            # Cấp 2 thụt lề sâu hơn (2.0cm)
            pf.left_indent = Cm(2.0)
            pf.first_line_indent = -Cm(0.63)

        pf.right_indent = Cm(0)
        pf.line_spacing = Config.LINE_SPACING
        pf.space_before = Pt(Config.SPACE_BEFORE_PT)
        pf.space_after = Pt(Config.SPACE_AFTER_PT)

        # 6. Reset style override của run (giữ bold/italic)
        RunManipulator.reset_run_style_inheritance(paragraph, keep_bold=True, keep_italic=True)

    @staticmethod
    def determine_list_level(paragraph, ilvl: Optional[int] = None) -> int:
        """
        Xác định cấp độ danh sách (1 hoặc 2):
        - Nếu có ilvl từ numbering.xml: ilvl == 0 -> cấp 1; ilvl >= 1 -> cấp 2
        - Nếu là văn bản thường: kiểm tra left_indent và ký tự bắt đầu
        """
        if ilvl is not None:
            return 2 if ilvl >= 1 else 1

        # Kiểm tra lề trái hiện tại
        left_indent = paragraph.paragraph_format.left_indent
        if left_indent and left_indent.pt > 20: # Lớn hơn ~0.7cm
            return 2

        # Kiểm tra văn bản bắt đầu bằng a), b), +, o
        text = paragraph.text.strip()
        if LEVEL_2_TRIGGER_PATTERN.match(text):
            return 2

        return 1
