"""
Unit tests cho Phân hệ 1: Document Zoning (Trang bìa & Bảng biểu)
"""

import unittest
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.table import WD_TABLE_ALIGNMENT

from core.document_zoner import DocumentZoner

class TestDocumentZoner(unittest.TestCase):
    def test_identify_cover_by_keyword(self):
        doc = Document()
        # Tạo trang bìa
        doc.add_paragraph("BỘ GIÁO DỤC VÀ ĐÀO TẠO")
        doc.add_paragraph("TRƯỜNG ĐẠI HỌC BÁCH KHOA")
        doc.add_paragraph("LUẬN VĂN TỐT NGHIỆP")
        doc.add_paragraph("Hà Nội, 2026")
        # Mốc nội dung
        doc.add_paragraph("LỜI MỞ ĐẦU")
        doc.add_paragraph("Nội dung mở đầu của luận văn...")

        zoner = DocumentZoner(doc)
        cover_end, cover_set = zoner.identify_zones()

        self.assertEqual(cover_end, 4)
        self.assertEqual(len(cover_set), 4)
        self.assertIn(0, cover_set)
        self.assertIn(3, cover_set)
        self.assertNotIn(4, cover_set)

    def test_identify_cover_by_section_break(self):
        doc = Document()
        p0 = doc.add_paragraph("TRANG BÌA 1")
        p1 = doc.add_paragraph("TRANG BÌA 2")
        # Thêm sectPr vào p1
        pPr = p1._element.get_or_add_pPr()
        sectPr = OxmlElement('w:sectPr')
        pPr.append(sectPr)

        doc.add_paragraph("CHƯƠNG 1: TỔNG QUAN")
        doc.add_paragraph("Nội dung chương 1...")

        zoner = DocumentZoner(doc)
        cover_end, cover_set = zoner.identify_zones()

        self.assertEqual(cover_end, 2)
        self.assertEqual(cover_set, {0, 1})

    def test_format_tables_separately(self):
        doc = Document()
        table = doc.add_table(rows=2, cols=2)
        cell = table.cell(0, 0)
        p = cell.paragraphs[0]
        p.text = "Tiêu đề cột 1"
        # Giả lập lề bị sai lệch trước đó
        p.paragraph_format.first_line_indent = Inches(0.5)

        zoner = DocumentZoner(doc)
        zoner.format_tables()

        # Bảng phải được căn giữa
        self.assertEqual(table.alignment, WD_TABLE_ALIGNMENT.CENTER)
        # Ô bảng TUYỆT ĐỐI KHÔNG có first_line_indent = 1.27cm (phải bằng 0)
        self.assertEqual(p.paragraph_format.first_line_indent.pt, 0)

if __name__ == '__main__':
    unittest.main()
