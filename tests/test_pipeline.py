"""
Integration test toàn diện cho Pipeline chuẩn hóa tài liệu Word (.docx)
"""

import os
import unittest
from docx import Document
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from config import Config
from core.pipeline import DocumentStandardizer
from tests.create_sample_docx import create_sample_messy_docx

class TestPipelineIntegration(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.sample_input = "test_input_sample.docx"
        cls.sample_output = "test_output_standardized.docx"
        create_sample_messy_docx(cls.sample_input)

    @classmethod
    def tearDownClass(cls):
        for f in [cls.sample_input, cls.sample_output]:
            if os.path.exists(f):
                try:
                    os.remove(f)
                except Exception:
                    pass

    def test_full_pipeline_processing(self):
        standardizer = DocumentStandardizer()
        result = standardizer.process(self.sample_input, output_target=self.sample_output)

        self.assertIn("stats", result)
        stats = result["stats"]

        # Kiểm tra thống kê cơ bản
        self.assertGreater(stats["cover_paragraphs"], 0)
        self.assertGreater(stats["heading_1_count"], 0)
        self.assertGreater(stats["heading_2_count"], 0)
        self.assertGreater(stats["heading_3_count"], 0)
        self.assertGreater(stats["heading_unnumbered_count"], 0)
        self.assertGreater(stats["list_items_count"], 0)
        self.assertGreater(stats["body_paragraphs_count"], 0)
        self.assertGreater(stats["caption_count"], 0)
        self.assertTrue(stats["toc_flag_enabled"])

        # Mở lại file output đã chuẩn hóa để kiểm tra cấu trúc thực tế
        out_doc = Document(self.sample_output)

        # 1. Kiểm tra lề trang chuẩn theo Form mẫu
        for sec in out_doc.sections:
            self.assertAlmostEqual(sec.top_margin.cm, Config.MARGIN_TOP_CM, places=2)
            self.assertAlmostEqual(sec.bottom_margin.cm, Config.MARGIN_BOTTOM_CM, places=2)
            self.assertAlmostEqual(sec.left_margin.cm, Config.MARGIN_LEFT_CM, places=2)
            self.assertAlmostEqual(sec.right_margin.cm, Config.MARGIN_RIGHT_CM, places=2)

        # 1.5. Kiểm tra Normal style đạt chuẩn 14pt đồng bộ
        self.assertEqual(out_doc.styles['Normal'].font.size.pt, 14.0)
        self.assertEqual(out_doc.styles['Normal'].font.name, Config.FONT_NAME)

        # 2. Kiểm tra updateFields trong settings.xml
        update_fields = out_doc.settings.element.find(qn('w:updateFields'))
        self.assertIsNotNone(update_fields)
        self.assertEqual(update_fields.get(qn('w:val')), 'true')

        # 3. Kiểm tra bảo tồn bold / italic trong các run
        bold_found = False
        italic_found = False
        for p in out_doc.paragraphs:
            for r in p.runs:
                if r.bold:
                    bold_found = True
                if r.italic:
                    italic_found = True
        self.assertTrue(bold_found, "Phải bảo tồn các đoạn in đậm nội tại!")
        self.assertTrue(italic_found, "Phải bảo tồn các đoạn in nghiêng nội tại!")

        # 4. Kiểm tra danh sách bullet đã chuẩn hóa về "- "
        list_marker_found = False
        for p in out_doc.paragraphs:
            if p.text.strip().startswith("- "):
                list_marker_found = True
                break
        self.assertTrue(list_marker_found, "Danh sách bullet phải được chuẩn hóa về tiền tố '- '")

        # 5. Kiểm tra bảng biểu: ô bảng first_line_indent phải bằng 0
        self.assertGreater(len(out_doc.tables), 0)
        for tbl in out_doc.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        self.assertEqual(p.paragraph_format.first_line_indent.pt, 0)

        # 6. Kiểm tra bẫy numbered list: câu "1. Bước một..." và "2. Bước hai..." không được là Heading 1
        trap1_found = False
        trap2_found = False
        for p in out_doc.paragraphs:
            if "Bước một: Tiếp nhận file" in p.text:
                self.assertNotEqual(p.style.name, 'Heading 1', "Bước một không được là Heading 1")
                trap1_found = True
            if "Bước hai: Giải mã XML" in p.text:
                self.assertNotEqual(p.style.name, 'Heading 1', "Bước hai không được là Heading 1")
                trap2_found = True
        self.assertTrue(trap1_found, "Phải tìm thấy Bước một!")
        self.assertTrue(trap2_found, "Phải tìm thấy Bước hai không bị gán nhầm làm Heading 1!")

    def test_nfd_caption_assigned_custom_style(self):
        """Kiểm thử caption với Unicode NFD (tổ hợp) được nhận diện chuẩn xác và gán style hinhanh / bang"""
        import unicodedata
        doc = Document()
        # Chèn văn bản với ký tự Unicode NFD phân rã
        nfd_fig = unicodedata.normalize('NFD', "Hình : Sơ đồ kiến trúc logic")
        nfd_tbl = unicodedata.normalize('NFD', "Bảng 1.1: Danh sách thông số")
        p_fig = doc.add_paragraph(nfd_fig)
        p_tbl = doc.add_paragraph(nfd_tbl)

        standardizer = DocumentStandardizer()
        res = standardizer.process(doc, auto_insert_toc=False, evaluate_toc=False)

        self.assertEqual(p_fig.style.name, 'hinhanh', "Caption hình ảnh Unicode NFD phải được gán style 'hinhanh'")
        self.assertEqual(p_tbl.style.name, 'bang', "Caption bảng biểu Unicode NFD phải được gán style 'bang'")
        self.assertEqual(p_fig.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.CENTER)
        self.assertEqual(p_tbl.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.CENTER)

    def test_header_footer_distance_standardization(self):
        """Kiểm thử chuẩn hóa lề trang bao gồm khoảng cách header/footer = 1.27cm (720 dxa) theo Form mẫu"""
        doc = Document()
        standardizer = DocumentStandardizer()
        standardizer.process(doc, auto_insert_toc=False, evaluate_toc=False)

        for sec in doc.sections:
            self.assertAlmostEqual(sec.header_distance.cm, 1.27, places=2)
            self.assertAlmostEqual(sec.footer_distance.cm, 1.27, places=2)

if __name__ == '__main__':
    unittest.main()
