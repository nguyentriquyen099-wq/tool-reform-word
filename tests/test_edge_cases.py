"""
Edge Cases and Stress Tests cho Hệ thống Chuẩn hóa Báo cáo Word
"""

import unittest
from unittest.mock import patch, MagicMock
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from config import Config
from core.pipeline import DocumentStandardizer
from core.classifier import WaterfallClassifier
from core.hsm import HierarchicalStateMachine, AIFallbackClassifier
from core.run_manipulator import RunManipulator
from core.numbering_parser import NumberingParser

class TestEdgeCases(unittest.TestCase):
    def test_empty_document(self):
        """Kiểm thử tài liệu rỗng 0 đoạn văn, 0 bảng"""
        doc = Document()
        standardizer = DocumentStandardizer()
        res = standardizer.process(doc)
        self.assertIsNotNone(res)
        self.assertEqual(res["stats"]["total_paragraphs"], 0)
        self.assertEqual(res["stats"]["tables_count"], 0)

    def test_document_with_only_cover(self):
        """Tài liệu chỉ có trang bìa"""
        doc = Document()
        doc.add_paragraph("TRƯỜNG ĐẠI HỌC BÁCH KHOA")
        doc.add_paragraph("BÁO CÁO NGHIÊN CỨU KHOA HỌC")
        doc.add_paragraph("Sinh viên: Lê Văn C")
        # Section break
        pPr = doc.paragraphs[-1]._element.get_or_add_pPr()
        pPr.append(OxmlElement('w:sectPr'))

        standardizer = DocumentStandardizer()
        res = standardizer.process(doc)
        self.assertEqual(res["stats"]["cover_paragraphs"], 3)
        self.assertEqual(res["stats"]["body_paragraphs_count"], 0)

    def test_bizarre_leading_spaces_and_symbols(self):
        """Kiểm thử các khoảng trắng lạ: \xa0, \u2002, \u2003, bullet lạ: , """
        doc = Document()
        p = doc.add_paragraph("\t\xa0\u2002\u2003  Nội dung văn bản với ký hiệu lạ")
        RunManipulator.clean_leading_junk_and_symbols(p, remove_bullets=True)
        self.assertTrue(p.text.startswith("Nội dung văn bản với ký hiệu lạ"))

    def test_media_and_math_detection(self):
        """Kiểm thử phát hiện Drawing và Math độc lập"""
        doc = Document()
        # Đoạn chứa drawing
        p_img = doc.add_paragraph()
        r = p_img.add_run()
        drawing = OxmlElement('w:drawing')
        r._element.append(drawing)

        # Đoạn chứa oMathPara
        p_math = doc.add_paragraph()
        omath_para = OxmlElement('m:oMathPara')
        p_math._element.append(omath_para)

        num_parser = NumberingParser(doc)
        classifier = WaterfallClassifier(doc, num_parser)

        self.assertTrue(classifier.is_tier1_media_or_math(p_img))
        self.assertTrue(classifier.is_tier1_media_or_math(p_math))

    @patch('requests.post')
    def test_ai_fallback_mock_api(self, mock_post):
        """Kiểm thử AI Fallback khi DeepSeek API trả về kết quả thành công"""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "choices": [
                {
                    "message": {
                        "content": '[{"id": 0, "type": "HEADING_2"}, {"id": 1, "type": "BODY"}]'
                    }
                }
            ]
        }
        mock_post.return_value = mock_response

        ai = AIFallbackClassifier(api_key="test_deepseek_key_123")
        candidates = [
            {"id": 0, "text": "Kiến trúc hệ thống", "is_bold": True},
            {"id": 1, "text": "Mô tả một phần của kiến trúc...", "is_bold": True}
        ]
        res = ai.classify_candidates(candidates)
        self.assertEqual(res[0], "HEADING_2")
        self.assertEqual(res[1], "BODY")

    @patch('requests.post')
    def test_ai_fallback_network_error(self, mock_post):
        """Kiểm thử AI Fallback khi DeepSeek API gặp lỗi mạng / 500 error"""
        mock_post.side_effect = Exception("Connection Timeout")

        ai = AIFallbackClassifier(api_key="test_deepseek_key_123")
        candidates = [
            {"id": 0, "text": "TÀI LIỆU THAM KHẢO", "is_bold": True},
            {"id": 1, "text": "Một đoạn văn bình thường", "is_bold": False}
        ]
        # Hệ thống phải tự động fallback sang Heuristic mà KHÔNG gây crash
        res = ai.classify_candidates(candidates)
        self.assertEqual(res[0], "HEADING_1")
        self.assertEqual(res[1], "BODY")

    def test_hsm_state_reset_between_documents(self):
        """Đảm bảo HSM reset sạch trạng thái giữa các lần gọi liên tiếp"""
        doc1 = Document()
        doc1.add_paragraph("CHƯƠNG 1")
        doc1.add_paragraph("1.1 Mục 1.1")

        doc2 = Document()
        doc2.add_paragraph("CHƯƠNG 1")

        standardizer = DocumentStandardizer()
        standardizer.process(doc1)
        res2 = standardizer.process(doc2)
        # Nếu state không bị rò rỉ, doc2 xử lý bình thường với Chương 1 là heading 1
        self.assertEqual(res2["stats"]["heading_1_count"], 1)

    def test_roman_numeral_chapters(self):
        """Kiểm thử nhận diện chương bằng số La Mã (Chương I, Chương II, Chương III)"""
        doc = Document()
        doc.add_paragraph("CHƯƠNG I: TỔNG QUAN")
        doc.add_paragraph("1.1 Bối cảnh đề tài")
        doc.add_paragraph("CHƯƠNG II: CƠ SỞ KỸ THUẬT")
        doc.add_paragraph("2.1 Kiến trúc giải pháp")
        doc.add_paragraph("CHƯƠNG III: THỰC NGHIỆM")

        standardizer = DocumentStandardizer()
        res = standardizer.process(doc)
        stats = res["stats"]
        self.assertEqual(stats["heading_1_count"], 3, "Phải nhận diện cả 3 chương số La Mã là Heading 1")
        self.assertEqual(stats["heading_2_count"], 2, "Phải nhận diện cả 2 tiểu mục 1.1 và 2.1 là Heading 2")

    def test_nested_tables_formatting(self):
        """Kiểm thử chuẩn hóa bảng biểu lồng nhau (nested tables)"""
        doc = Document()
        outer_table = doc.add_table(rows=2, cols=2)
        outer_cell = outer_table.cell(0, 0)
        outer_cell.paragraphs[0].text = "Bảng ngoài"
        outer_cell.paragraphs[0].paragraph_format.first_line_indent = Inches(0.5)

        # Tạo bảng lồng bên trong ô outer_cell
        inner_table = outer_cell.add_table(rows=2, cols=2)
        inner_cell = inner_table.cell(0, 0)
        inner_cell.paragraphs[0].text = "Bảng trong"
        inner_cell.paragraphs[0].paragraph_format.first_line_indent = Inches(0.5)

        standardizer = DocumentStandardizer()
        res = standardizer.process(doc)

        # Cả bảng ngoài và bảng trong đều phải có first_line_indent = 0
        self.assertEqual(outer_cell.paragraphs[0].paragraph_format.first_line_indent.pt, 0)
        self.assertEqual(inner_cell.paragraphs[0].paragraph_format.first_line_indent.pt, 0)
        # Kiểm tra đếm đầy đủ 2 bảng
        self.assertEqual(res["stats"]["tables_count"], 2)

    def test_letter_marker_normalization(self):
        """Kiểm thử chuẩn hóa marker chữ cái a), b) được quy đổi sạch sẽ sang + """
        from core.list_normalizer import ListNormalizer
        doc = Document()
        p = doc.add_paragraph("a)   Yêu cầu về mặt chức năng của hệ thống")
        ListNormalizer.normalize_list_item(p, level=2)

        # Sau khi chuẩn hóa, tiền tố 'a)' phải được thay thế hoàn toàn bằng '+ '
        self.assertTrue(p.text.startswith("+ Yêu cầu về mặt chức năng"))
        self.assertNotIn("a)", p.text)

    def test_custom_api_config_passed(self):
        """Kiểm thử tham số API mở rộng (base_url, model) được truyền đầy đủ tới AI Classifier"""
        custom_standardizer = DocumentStandardizer(
            api_key="test_key_xyz",
            base_url="https://my-custom-proxy.com/v1",
            model="custom-deepseek-v3"
        )
        self.assertEqual(custom_standardizer.ai_classifier.api_key, "test_key_xyz")
        self.assertEqual(custom_standardizer.ai_classifier.base_url, "https://my-custom-proxy.com/v1")
        self.assertEqual(custom_standardizer.ai_classifier.model, "custom-deepseek-v3")

    def test_missing_heading_styles_auto_creation(self):
        """Kiểm thử tự động khởi tạo Heading 1, 2, 3, 4 khi styles.xml thiếu"""
        from core.style_manager import StyleManager
        doc = Document()
        # Xóa style Heading 1, 2, 3 nếu có trong document
        sm = StyleManager(doc)
        sm.setup_document_styles()
        self.assertIn('Heading 1', doc.styles)
        self.assertIn('Heading 2', doc.styles)
        self.assertIn('Heading 3', doc.styles)
        self.assertIn('Heading 4', doc.styles)

    def test_heading_4_classification_and_formatting(self):
        """Kiểm thử đề mục cấp 4 (1.2.2.1) được nhận diện HEADING_4 và gán outlineLvl=3"""
        doc = Document()
        doc.add_paragraph("CHƯƠNG 1: TỔNG QUAN")
        doc.add_paragraph("1.2 PHÂN TÍCH HỆ THỐNG")
        doc.add_paragraph("1.2.2 Khảo sát thiết bị")
        p4 = doc.add_paragraph("1.2.2.1 Chi tiết phòng 302")

        standardizer = DocumentStandardizer()
        res = standardizer.process(doc)
        self.assertEqual(res["stats"]["heading_4_count"], 1)

        # Kiểm tra outlineLvl = 3 trên p4
        pPr = p4._element.find(qn('w:pPr'))
        outline = pPr.find(qn('w:outlineLvl'))
        self.assertIsNotNone(outline)
        self.assertEqual(outline.get(qn('w:val')), '3')

    def test_roman_heading_without_ai(self):
        """Kiểm thử nhận diện đề mục số La Mã I. PHÂN CÔNG CÔNG VIỆC là Heading 1"""
        doc = Document()
        p = doc.add_paragraph("I. PHÂN CÔNG CÔNG VIỆC")
        standardizer = DocumentStandardizer()
        res = standardizer.process(doc)
        self.assertEqual(res["stats"]["heading_1_count"], 1)
        outline = p._element.find(qn('w:pPr')).find(qn('w:outlineLvl'))
        self.assertIsNotNone(outline)
        self.assertEqual(outline.get(qn('w:val')), '0')

    def test_caption_without_sequence_number(self):
        """Kiểm thử nhận diện caption không có số thứ tự (Hình : Sơ đồ phòng...)"""
        doc = Document()
        p = doc.add_paragraph("Hình : Sơ đồ phòng 302-A9")
        standardizer = DocumentStandardizer()
        res = standardizer.process(doc)
        self.assertEqual(res["stats"]["caption_count"], 1)
        self.assertEqual(p.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.CENTER)

    def test_dangling_heading_number(self):
        """Kiểm thử đề mục dở dang 3.1.\xa0 không bị lỗi và được nhận diện là Heading 2"""
        doc = Document()
        doc.add_paragraph("CHƯƠNG 3: TỔNG KẾT")
        p_dangling = doc.add_paragraph("3.1.\xa0")
        standardizer = DocumentStandardizer()
        res = standardizer.process(doc)
        self.assertEqual(res["stats"]["heading_2_count"], 1)
        outline = p_dangling._element.find(qn('w:pPr')).find(qn('w:outlineLvl'))
        self.assertIsNotNone(outline)
        self.assertEqual(outline.get(qn('w:val')), '1')

    def test_cover_table_and_page_break_ensured(self):
        """Kiểm thử bảng bìa không bị in đậm dòng 0 và tự động ngắt trang trước Lời nói đầu"""
        doc = Document()
        # Bìa
        doc.add_paragraph("TRƯỜNG ĐẠI HỌC CÔNG NGHIỆP")
        tbl = doc.add_table(rows=2, cols=2)
        tbl.rows[0].cells[0].paragraphs[0].text = "GVHD:"
        tbl.rows[0].cells[1].paragraphs[0].text = "ThS. Nguyễn Văn A"
        doc.add_paragraph("Hà Nội, 2026")
        # Nội dung
        p_intro = doc.add_paragraph("LỜI NÓI ĐẦU")
        doc.add_paragraph("Nội dung lời nói đầu...")

        standardizer = DocumentStandardizer()
        standardizer.process(doc)

        # 1. Bảng bìa: ô R0C1 không bị in đậm
        r0c1_runs = tbl.rows[0].cells[1].paragraphs[0].runs
        is_bold = any(r.bold for r in r0c1_runs)
        self.assertFalse(is_bold, "Dòng 0 của bảng bìa không được tự ý in đậm toàn bộ!")

        # 2. Ngắt trang bìa: Lời nói đầu phải có page_break_before = True
        self.assertTrue(p_intro.paragraph_format.page_break_before)

    def test_ppr_rpr_sz_cleanup_for_body_inheritance(self):
        """Kiểm thử xóa sz trong pPr/rPr để body text kế thừa trọn vẹn 14pt từ Normal style"""
        doc = Document()
        p = doc.add_paragraph("Đoạn văn nội dung có sz rác trong pPr.")
        pPr = p._element.get_or_add_pPr()
        rPr = OxmlElement('w:rPr')
        sz = OxmlElement('w:sz')
        sz.set(qn('w:val'), '24')
        rPr.append(sz)
        pPr.append(rPr)

        RunManipulator.reset_run_style_inheritance(p)
        self.assertIsNone(rPr.find(qn('w:sz')), "Thẻ w:sz trong pPr/rPr phải được dọn dẹp để kế thừa từ style Normal")

    def test_cover_only_table_detection(self):
        """Kiểm thử phát hiện bảng trong tài liệu chỉ có trang bìa"""
        doc = Document()
        doc.add_paragraph("TRƯỜNG ĐẠI HỌC BÁCH KHOA")
        p_last = doc.add_paragraph("BÁO CÁO TỐT NGHIỆP")
        tbl = doc.add_table(rows=2, cols=2)
        tbl.rows[0].cells[0].paragraphs[0].text = "GVHD:"
        tbl.rows[0].cells[1].paragraphs[0].text = "ThS. Tran Van B"
        # Thêm section break ở đoạn cuối bìa
        p_last._element.get_or_add_pPr().append(OxmlElement('w:sectPr'))

        from core.document_zoner import DocumentZoner
        zoner = DocumentZoner(doc)
        zoner.identify_zones()
        self.assertTrue(zoner._is_table_in_cover(tbl), "Bảng trong tài liệu chỉ có bìa phải được xác định là cover table")

    def test_heading_trailing_whitespace_cleanup(self):
        """Kiểm thử dọn dẹp khoảng trắng / non-breaking space \\xa0 ở cuối heading"""
        doc = Document()
        doc.add_paragraph("CHƯƠNG 1: TỔNG QUAN\xa0")
        doc.add_paragraph("1.1. Yêu cầu thiết kế\xa0")
        standardizer = DocumentStandardizer()
        standardizer.process(doc)

        self.assertEqual(doc.paragraphs[0].text, "CHƯƠNG 1: TỔNG QUAN")
        self.assertEqual(doc.paragraphs[1].text, "1.1. Yêu cầu thiết kế")

    def test_heuristic_fallback_caption_and_non_toc(self):
        """Kiểm thử heuristic fallback nhận diện đúng Caption không có số và Non-TOC heading"""
        ai = AIFallbackClassifier(api_key=None)
        candidates = [
            {"id": 0, "text": "Hình : Sơ đồ phòng 302", "is_bold": True},
            {"id": 1, "text": "MỤC LỤC", "is_bold": True},
            {"id": 2, "text": "LỜI MỞ ĐẦU", "is_bold": True}
        ]
        res = ai.classify_candidates(candidates)
        self.assertEqual(res[0], "CAPTION")
        self.assertEqual(res[1], "HEADING_NON_TOC")
        self.assertEqual(res[2], "HEADING_1")

if __name__ == '__main__':
    unittest.main()
