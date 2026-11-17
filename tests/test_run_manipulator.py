"""
Unit tests cho Phân hệ 5 & Lưu ý 1: Run-level Manipulation & Style Reset
"""

import unittest
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from core.run_manipulator import RunManipulator

class TestRunManipulator(unittest.TestCase):
    def test_clean_fragmented_runs_without_overwriting_runs(self):
        doc = Document()
        p = doc.add_paragraph()
        r0 = p.add_run("\t   ")
        r1 = p.add_run("•  ")
        r2 = p.add_run("Đây là đoạn văn có ")
        r3 = p.add_run("từ khóa in đậm")
        r3.bold = True
        r4 = p.add_run(" và ")
        r5 = p.add_run("in nghiêng.")
        r5.italic = True

        initial_run_count = len(p.runs)
        self.assertEqual(initial_run_count, 6)

        # Làm sạch leading junk
        removed = RunManipulator.clean_leading_junk_and_symbols(p, remove_bullets=True)

        # Số run vẫn được giữ nguyên để bảo tồn formatting
        self.assertEqual(len(p.runs), initial_run_count)
        # Run in đậm và in nghiêng vẫn giữ nguyên thuộc tính
        self.assertTrue(p.runs[3].bold)
        self.assertEqual(p.runs[3].text, "từ khóa in đậm")
        self.assertTrue(p.runs[5].italic)
        self.assertEqual(p.runs[5].text, "in nghiêng.")

        # Văn bản nội dung kết hợp không còn bullet
        self.assertTrue(p.text.startswith("Đây là đoạn văn"))

    def test_remove_w_sym_element(self):
        doc = Document()
        p = doc.add_paragraph()
        r0 = p.add_run()
        # Chèn thẻ w:sym giả lập Wingdings bullet
        sym = OxmlElement('w:sym')
        sym.set(qn('w:font'), 'Wingdings')
        sym.set(qn('w:char'), 'F0B7')
        r0._element.append(sym)

        r1 = p.add_run("Nội dung danh sách có symbol ẩn")

        self.assertEqual(len(p._element.xpath('.//w:sym')), 1)

        RunManipulator.clean_leading_junk_and_symbols(p, remove_bullets=True)

        # Thẻ w:sym phải bị xóa hoàn toàn
        self.assertEqual(len(p._element.xpath('.//w:sym')), 0)
        self.assertIn("Nội dung danh sách", p.text)

    def test_reset_run_style_inheritance_preserves_bold_italic(self):
        doc = Document()
        p = doc.add_paragraph()
        r1 = p.add_run("Bình thường ")
        r2 = p.add_run("Rất quan trọng ")
        r2.bold = True
        r2.font.name = "Arial" # font rác
        r2.font.size = Pt(20)  # size rác
        r3 = p.add_run("Ngoại lệ")
        r3.italic = True

        RunManipulator.reset_run_style_inheritance(p, keep_bold=True, keep_italic=True)

        # Override font và size đã bị gỡ để kế thừa style gốc
        self.assertIsNone(r2.font.name)
        self.assertIsNone(r2.font.size)
        # Nhưng bold và italic vẫn được bảo toàn tuyệt đối!
        self.assertTrue(r2.bold)
        self.assertTrue(r3.italic)

if __name__ == '__main__':
    unittest.main()
