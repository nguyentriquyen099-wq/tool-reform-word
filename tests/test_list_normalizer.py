"""
Unit tests cho Phân hệ 4: List Normalizer
"""

import unittest
from docx import Document
from docx.shared import Cm, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from config import Config
from core.list_normalizer import ListNormalizer

class TestListNormalizer(unittest.TestCase):
    def test_normalize_level_1_bullet(self):
        doc = Document()
        p = doc.add_paragraph("•  Mục danh sách cấp 1")
        # Thêm w:numPr giả lập Word auto bullet
        pPr = p._element.get_or_add_pPr()
        numPr = OxmlElement('w:numPr')
        pPr.append(numPr)

        ListNormalizer.normalize_list_item(p, level=1)

        # 1. Văn bản phải bắt đầu bằng "- "
        self.assertTrue(p.text.startswith("- Mục danh sách cấp 1"))
        # 2. Thẻ numPr phải bị gỡ
        self.assertIsNone(p._element.find(qn('w:pPr')).find(qn('w:numPr')))
        # 3. Thụt lề cấp 1 (Hanging Indent): left_indent = Config.FIRST_LINE_INDENT_CM, first_line_indent = -0.63cm
        self.assertAlmostEqual(p.paragraph_format.left_indent.cm, Config.FIRST_LINE_INDENT_CM, places=2)
        self.assertAlmostEqual(p.paragraph_format.first_line_indent.cm, -0.63, places=2)

    def test_normalize_level_2_bullet(self):
        doc = Document()
        p = doc.add_paragraph("❖  Mục danh sách cấp 2 sâu hơn")

        ListNormalizer.normalize_list_item(p, level=2)

        # 1. Văn bản phải bắt đầu bằng "+ "
        self.assertTrue(p.text.startswith("+ Mục danh sách cấp 2"))
        # 2. Thụt lề cấp 2 (Hanging Indent): left_indent = 2.0cm, first_line_indent = -0.63cm
        self.assertAlmostEqual(p.paragraph_format.left_indent.cm, 2.0, places=2)
        self.assertAlmostEqual(p.paragraph_format.first_line_indent.cm, -0.63, places=2)

if __name__ == '__main__':
    unittest.main()
