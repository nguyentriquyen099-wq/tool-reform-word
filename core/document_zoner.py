"""
Phân hệ 1: Document Zoning (Phân vùng tài liệu)
- Khóa vùng Trang bìa (xác định section break <w:sectPr> hoặc mốc "MỞ ĐẦU", "LỜI MỞ ĐẦU", "CHƯƠNG 1",...)
- Tách riêng Bảng biểu (document.tables): đảm bảo không áp dụng first_line_indent = 1.27cm vào ô bảng,
  căn chỉnh bảng khoa học (giữa trang, giãn dòng gọn, cỡ chữ 12pt).
"""

import re
import logging
from typing import List, Tuple, Set
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from config import Config

logger = logging.getLogger(__name__)

class DocumentZoner:
    def __init__(self, doc):
        self.doc = doc
        self.cover_end_index = -1
        self.cover_paragraphs: Set[int] = set()

    def identify_zones(self) -> Tuple[int, Set[int]]:
        """
        Xác định ranh giới giữa Trang bìa (Cover Zone) và Nội dung chính (Body Zone).
        Trả về: (cover_end_index, tập_hợp_index_trang_bìa)
        """
        paragraphs = self.doc.paragraphs
        total = len(paragraphs)
        if total == 0:
            return -1, set()

        boundary_idx = -1

        # 1. Quét tìm section break <w:sectPr> trong các đoạn đầu tiên (tối đa 30 đoạn)
        for i in range(min(total, 35)):
            p = paragraphs[i]
            # Kiểm tra section break
            sect_pr = p._element.xpath('./w:pPr/w:sectPr')
            if sect_pr:
                # Đoạn có sectPr là đoạn cuối cùng của Section 1 (Trang bìa)
                boundary_idx = i + 1
                logger.info(f"Phát hiện Section Break trang bìa tại đoạn {i}")
                break

        # 2. Nếu không có Section Break rõ ràng, tìm theo Mốc từ khóa nội dung đầu tiên
        if boundary_idx == -1:
            for i in range(min(total, 40)):
                p = paragraphs[i]
                text = p.text.strip().upper()
                if not text:
                    continue

                for kw in Config.COVER_BOUNDARY_KEYWORDS:
                    # Kiểm tra trùng khớp từ khóa (chính xác hoặc ở đầu câu)
                    if kw == "TÓM TẮT":
                        if text == "TÓM TẮT" or text in ["TÓM TẮT NỘI DUNG", "TÓM TẮT ĐỀ TÀI", "TÓM TẮT BÁO CÁO"]:
                            boundary_idx = i
                            logger.info(f"Phát hiện ranh giới Trang bìa bằng từ khóa '{kw}' tại đoạn {i}")
                            break
                    elif text == kw or text.startswith(kw + " ") or text.startswith(kw + ":") or text.startswith(kw + " -"):
                        # Mốc nội dung bắt đầu từ đoạn i -> Trang bìa là các đoạn trước i
                        boundary_idx = i
                        logger.info(f"Phát hiện ranh giới Trang bìa bằng từ khóa '{kw}' tại đoạn {i}")
                        break
                if boundary_idx != -1:
                    break

        # 3. Nếu vẫn không tìm thấy, kiểm tra ngắt trang thủ công (w:br type="page") trong 25 đoạn đầu
        if boundary_idx == -1:
            for i in range(min(total, 25)):
                p = paragraphs[i]
                page_br = p._element.xpath('.//w:br[@w:type="page"]')
                if page_br:
                    boundary_idx = i + 1
                    logger.info(f"Phát hiện ngắt trang thủ công tại đoạn {i}")
                    break

        self.cover_end_index = boundary_idx
        if boundary_idx > 0:
            self.cover_paragraphs = set(range(boundary_idx))
        else:
            self.cover_paragraphs = set()

        return self.cover_end_index, self.cover_paragraphs

    @staticmethod
    def _set_run_font(run, font_name: str, font_size_pt: float = None, bold: bool = None):
        """Thiết lập font và đồng bộ đầy đủ thuộc tính rFonts cho tiếng Việt Unicode"""
        run.font.name = font_name
        if font_size_pt is not None:
            run.font.size = Pt(font_size_pt)
        if bold is not None:
            run.bold = bold
        rPr = run._element.get_or_add_rPr()
        rfonts = rPr.find(qn('w:rFonts'))
        if rfonts is None:
            rfonts = OxmlElement('w:rFonts')
            rPr.insert(0, rfonts)
        rfonts.set(qn('w:ascii'), font_name)
        rfonts.set(qn('w:hAnsi'), font_name)
        rfonts.set(qn('w:cs'), font_name)
        rfonts.set(qn('w:eastAsia'), font_name)

    def format_cover_page(self):
        """
        Khóa vùng trang bìa:
        - Giữ nguyên bố cục căn giữa / thụt lề ban đầu của trang bìa
        - Chỉ đảm bảo Font là Times New Roman và màu đen (không bị lỗi font)
        - TUYỆT ĐỐI KHÔNG áp dụng thụt lề 1.27cm hay justify
        - Đảm bảo ngắt trang (Page Break) giữa Trang bìa và Nội dung chính
        """
        if not self.cover_paragraphs:
            return

        for idx in self.cover_paragraphs:
            p = self.doc.paragraphs[idx]
            for run in p.runs:
                # Đảm bảo font Times New Roman nhưng giữ nguyên size & bold do người dùng thiết kế bìa
                self._set_run_font(run, Config.FONT_NAME)

        self.ensure_cover_page_break()

    def ensure_cover_page_break(self):
        """
        Tự động đảm bảo có ngắt trang (Page Break) giữa Trang bìa và phần nội dung:
        Nếu chưa có Section Break hoặc Page Break, gán page_break_before = True
        cho đoạn đầu tiên của vùng nội dung.
        """
        if self.cover_end_index <= 0 or self.cover_end_index >= len(self.doc.paragraphs):
            return

        # Kiểm tra xem các đoạn trong trang bìa đã có Section Break hoặc Page Break chưa
        has_break = False
        for idx in self.cover_paragraphs:
            p = self.doc.paragraphs[idx]
            if p._element.xpath('.//w:sectPr') or p._element.xpath('.//w:br[@w:type="page"]'):
                has_break = True
                break

        first_body_p = self.doc.paragraphs[self.cover_end_index]
        if first_body_p.paragraph_format.page_break_before:
            has_break = True

        if not has_break:
            first_body_p.paragraph_format.page_break_before = True
            logger.info(f"Đã đảm bảo ngắt trang tại đoạn {self.cover_end_index} ('{first_body_p.text[:30]}')")

    def _is_table_in_cover(self, table) -> bool:
        """Kiểm tra xem bảng có nằm trong Vùng Trang Bìa (trước boundary_idx) không"""
        if self.cover_end_index <= 0:
            return False
        if self.cover_end_index >= len(self.doc.paragraphs):
            # Toàn bộ tài liệu là trang bìa -> Bảng chắc chắn thuộc trang bìa
            return True
        try:
            boundary_elem = self.doc.paragraphs[self.cover_end_index]._element
            body = self.doc._element.body
            elem = table._element
            while elem is not None and elem.getparent() != body:
                elem = elem.getparent()
            if elem is None:
                return False
            tbl_idx = body.index(elem)
            boundary_idx = body.index(boundary_elem)
            return tbl_idx < boundary_idx
        except Exception:
            return False

    def _format_single_table(self, table, is_nested: bool = False, is_cover_table: bool = False):
        """Định dạng 1 bảng (hỗ trợ cả bảng lồng nhau)"""
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        for row_idx, row in enumerate(table.rows):
            for cell in row.cells:
                for p in cell.paragraphs:
                    pf = p.paragraph_format
                    # Khóa thụt lề = 0
                    pf.first_line_indent = Inches(0)
                    pf.left_indent = Inches(0)
                    pf.right_indent = Inches(0)
                    pf.line_spacing = 1.15
                    pf.space_before = Pt(2)
                    pf.space_after = Pt(2)

                    # Dòng đầu của bảng chính (Header row) -> in đậm, căn giữa (trừ bảng trang bìa)
                    if row_idx == 0 and not is_nested and not is_cover_table:
                        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for run in p.runs:
                            self._set_run_font(run, Config.FONT_NAME, Config.TABLE_FONT_SIZE, bold=True)
                    else:
                        for run in p.runs:
                            self._set_run_font(run, Config.FONT_NAME, Config.TABLE_FONT_SIZE)
                            if run.font.color and run.font.color.rgb:
                                run.font.color.rgb = RGBColor(0, 0, 0)

                # Đệ quy format toàn bộ bảng lồng nhau bên trong ô (nested tables)
                for nested_tbl in cell.tables:
                    self._format_single_table(nested_tbl, is_nested=True, is_cover_table=is_cover_table)

    def format_tables(self):
        """
        Chuẩn hóa toàn bộ bảng biểu (document.tables bao gồm nested tables):
        - TUYỆT ĐỐI KHÔNG áp dụng first_line_indent = 1.27cm
        - Căn bảng ra giữa trang (WD_TABLE_ALIGNMENT.CENTER)
        - Định dạng chữ trong ô: Times New Roman, 12pt, giãn dòng 1.15, lề gọn
        - Bảng trang bìa: Bỏ qua logic format Header row in đậm/căn giữa
        """
        for table in self.doc.tables:
            is_cover = self._is_table_in_cover(table)
            self._format_single_table(table, is_nested=False, is_cover_table=is_cover)
