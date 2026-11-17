"""
Unit tests cho 3 tính năng mới:
1. Chuẩn hóa khoảng cách dòng & đoạn (Line spacing 1.2, Space before/after rút gọn cho Body và Headings).
2. Chuẩn hóa chữ hoa/thường cho Heading cấp con (Sentence Case, bảo tồn từ viết tắt LAN/IP/CNTT, giữ nguyên Chương ALL CAPS).
3. Thụt lề phân cấp cho Heading (Hierarchical Indentation: Cấp 1 kịch lề trái 0cm, Cấp 2 thụt 0.5cm, Cấp 3 thụt 1.0cm, Cấp 4 thụt 1.5cm).
"""

import unittest
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

from config import Config
from core.run_manipulator import RunManipulator, to_sentence_case
from core.style_manager import StyleManager
from core.pipeline import DocumentStandardizer
from core.classifier import ParagraphType

class TestHeadingFormattingAndSpacing(unittest.TestCase):

    def test_to_sentence_case_basic(self):
        """Kiểm thử to_sentence_case chuyển ALL CAPS sang Sentence Case"""
        raw = "PHÂN TÍCH NHIỆM VỤ, VỊ TRÍ LẮP ĐẶT ĐỂ ĐƯA RA MÔ HÌNH TỔNG QUÁT CỦA HỆ THỐNG VÀ LÝ DO LỰA CHỌN"
        expected = "Phân tích nhiệm vụ, vị trí lắp đặt để đưa ra mô hình tổng quát của hệ thống và lý do lựa chọn"
        self.assertEqual(to_sentence_case(raw), expected)

    def test_to_sentence_case_preserves_acronyms(self):
        """Kiểm thử to_sentence_case bảo tồn các từ viết tắt kỹ thuật (LAN, IP, TCP, CNTT, AI...)"""
        raw = "THIẾT KẾ HỆ THỐNG MẠNG LAN VÀ ĐỊA CHỈ IP DỰ PHÒNG CHO PHÒNG CNTT"
        expected = "Thiết kế hệ thống mạng LAN và địa chỉ IP dự phòng cho phòng CNTT"
        self.assertEqual(to_sentence_case(raw), expected)

    def test_to_sentence_case_preserves_alphanumeric_codes(self):
        """Kiểm thử bảo tồn mã phòng / thiết bị chứa cả chữ và số (302-A9, A9, v1)"""
        raw = "KHẢO SÁT VỊ TRÍ PHÒNG 302-A9 VÀ PHÒNG 305-A9"
        expected = "Khảo sát vị trí phòng 302-A9 và phòng 305-A9"
        self.assertEqual(to_sentence_case(raw), expected)

    def test_to_sentence_case_preserves_roman_numerals(self):
        """Kiểm thử bảo tồn số La Mã trong tiêu đề (PHẦN I, II, III...)"""
        raw = "ĐÁNH GIÁ KẾT QUẢ THỰC NGHIỆM GIAI ĐOẠN I VÀ II"
        expected = "Đánh giá kết quả thực nghiệm giai đoạn I và II"
        self.assertEqual(to_sentence_case(raw), expected)

    def test_to_sentence_case_ignores_already_mixed_case(self):
        """Tiêu đề đã viết hoa/thường bình thường thì giữ nguyên không đổi"""
        raw = "1.2.1. Phân tích nhiệm vụ thiết kế hệ thống mạng"
        self.assertEqual(to_sentence_case(raw), raw)

    def test_normalize_heading_casing_single_run(self):
        """Kiểm thử normalize_heading_casing trên đoạn 1 run"""
        doc = Document()
        p = doc.add_paragraph("1.2. PHÂN TÍCH NHIỆM VỤ, VỊ TRÍ LẮP ĐẶT")
        changed = RunManipulator.normalize_heading_casing(p)
        self.assertTrue(changed)
        self.assertEqual(p.text, "1.2. Phân tích nhiệm vụ, vị trí lắp đặt")

    def test_normalize_heading_casing_fragmented_runs_preserves_formatting(self):
        """Kiểm thử normalize_heading_casing trên nhiều Run phân mảnh, bảo tồn bold/italic"""
        doc = Document()
        p = doc.add_paragraph()
        r1 = p.add_run("1.2. ")
        r1.bold = True
        r2 = p.add_run("PHÂN TÍCH NHIỆM VỤ, ")
        r2.bold = True
        r3 = p.add_run("VỊ TRÍ LẮP ĐẶT ")
        r3.italic = True
        r4 = p.add_run("MẠNG LAN")
        r4.bold = True

        runs_count = len(p.runs)
        changed = RunManipulator.normalize_heading_casing(p)
        self.assertTrue(changed)
        self.assertEqual(p.text, "1.2. Phân tích nhiệm vụ, vị trí lắp đặt mạng LAN")
        self.assertEqual(len(p.runs), runs_count, "Phải bảo tồn toàn bộ các Run, không được xóa!")
        self.assertTrue(p.runs[1].bold)
        self.assertTrue(p.runs[2].italic)
        self.assertTrue(p.runs[3].bold)

    def test_hierarchical_heading_indentation_in_pipeline(self):
        """
        Kiểm thử thụt lề phân cấp cho Heading:
        - Cấp 1 (Chương / Heading 1): left_indent = 0cm
        - Cấp 2 (1.1. / Heading 2): left_indent = 0.5cm
        - Cấp 3 (1.2.1. / Heading 3): left_indent = 1.0cm
        - Cấp 4 (1.2.1.1. / Heading 4): left_indent = 1.5cm
        """
        doc = Document()
        doc.add_paragraph("CHƯƠNG 1: XÂY DỰNG HỆ THỐNG MẠNG")
        doc.add_paragraph("1.1. YÊU CẦU THIẾT KẾ HỆ THỐNG MẠNG")
        doc.add_paragraph("1.1.1. Yêu cầu băng thông")
        doc.add_paragraph("1.1.1.1. Băng thông cho phòng 302-A9")
        doc.add_paragraph("Nội dung thân bài...")

        standardizer = DocumentStandardizer()
        res = standardizer.process(doc, auto_insert_toc=False, evaluate_toc=False, normalize_chapter_to_roman=True)

        p_h1 = doc.paragraphs[0]
        p_h2 = doc.paragraphs[1]
        p_h3 = doc.paragraphs[2]
        p_h4 = doc.paragraphs[3]
        p_body = doc.paragraphs[4]

        # Kiểm tra thụt lề phân cấp theo cấu hình Form mẫu
        self.assertAlmostEqual(p_h1.paragraph_format.left_indent.cm, Config.HEADING_INDENT_PER_LEVEL_CM * 0, places=2)
        self.assertAlmostEqual(p_h2.paragraph_format.left_indent.cm, Config.HEADING_INDENT_PER_LEVEL_CM * 1, places=2)
        self.assertAlmostEqual(p_h3.paragraph_format.left_indent.cm, Config.HEADING_INDENT_PER_LEVEL_CM * 2, places=2)
        self.assertAlmostEqual(p_h4.paragraph_format.left_indent.cm, Config.HEADING_INDENT_PER_LEVEL_CM * 3, places=2)

        # Kiểm tra Heading 1 căn giữa theo chuẩn Form mẫu
        self.assertEqual(p_h1.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.CENTER)

        # Body text không bị áp đặt left_indent phân cấp của heading
        self.assertAlmostEqual(p_body.paragraph_format.left_indent.cm, 0.0, places=2)
        self.assertAlmostEqual(p_body.paragraph_format.first_line_indent.cm, Config.FIRST_LINE_INDENT_CM, places=2)

    def test_reduced_spacing_in_pipeline(self):
        """
        Kiểm thử khoảng cách dòng & đoạn đã được rút gọn thanh thoát:
        - Body: space_before = 0, space_after = 3pt, line_spacing = 1.2
        - Heading 1: space_before = 12pt, space_after = 4pt
        - Heading 2, 3, 4: space_before = 4pt, space_after = 2pt
        """
        doc = Document()
        doc.add_paragraph("CHƯƠNG 1: TỔNG QUAN")
        doc.add_paragraph("1.1. Khảo sát thực tế")
        doc.add_paragraph("1.1.1. Chi tiết phòng ban")
        doc.add_paragraph("Đoạn văn thân bài số 1...")

        standardizer = DocumentStandardizer()
        standardizer.process(doc, auto_insert_toc=False, evaluate_toc=False)

        p_h1 = doc.paragraphs[0]
        p_h2 = doc.paragraphs[1]
        p_h3 = doc.paragraphs[2]
        p_body = doc.paragraphs[3]

        # Heading 1 Spacing
        self.assertEqual(p_h1.paragraph_format.space_before.pt, Config.HEADING_1_SPACE_BEFORE_PT)
        self.assertEqual(p_h1.paragraph_format.space_after.pt, Config.HEADING_1_SPACE_AFTER_PT)

        # Heading 2 Spacing (gọn gàng, space_after = 2pt, space_before = 4pt)
        self.assertEqual(p_h2.paragraph_format.space_before.pt, Config.HEADING_SPACE_BEFORE_PT)
        self.assertEqual(p_h2.paragraph_format.space_after.pt, Config.HEADING_SPACE_AFTER_PT)

        # Heading 3 Spacing
        self.assertEqual(p_h3.paragraph_format.space_before.pt, Config.HEADING_SPACE_BEFORE_PT)
        self.assertEqual(p_h3.paragraph_format.space_after.pt, Config.HEADING_SPACE_AFTER_PT)

        # Body Spacing
        self.assertEqual(p_body.paragraph_format.space_before.pt, Config.SPACE_BEFORE_PT)
        self.assertEqual(p_body.paragraph_format.space_after.pt, Config.SPACE_AFTER_PT)
        self.assertAlmostEqual(p_body.paragraph_format.line_spacing, Config.LINE_SPACING, places=2)

    def test_casing_normalization_in_pipeline(self):
        """
        Kiểm thử pipeline tích hợp:
        - Heading 1 (Chương) giữ nguyên ALL CAPS
        - Heading 2 ALL CAPS tự động chuyển sang Sentence Case
        - Heading 3 giữ nguyên chữ thường
        """
        doc = Document()
        doc.add_paragraph("CHƯƠNG 1: XÂY DỰNG HỆ THỐNG MẠNG")
        doc.add_paragraph("1.2. PHÂN TÍCH NHIỆM VỤ, VỊ TRÍ LẮP ĐẶT ĐỂ ĐƯA RA MÔ HÌNH TỔNG QUÁT CỦA HỆ THỐNG VÀ LÝ DO LỰA CHỌN")
        doc.add_paragraph("1.2.1. Phân tích nhiệm vụ thiết kế hệ thống mạng")

        standardizer = DocumentStandardizer()
        res = standardizer.process(doc, auto_insert_toc=False, evaluate_toc=False, normalize_chapter_to_roman=True)

        p_h1 = doc.paragraphs[0]
        p_h2 = doc.paragraphs[1]
        p_h3 = doc.paragraphs[2]

        # Heading 1: Đổi số La Mã nhưng giữ nguyên ALL CAPS
        self.assertEqual(p_h1.text, "CHƯƠNG I: XÂY DỰNG HỆ THỐNG MẠNG")

        # Heading 2: Tự động chuyển ALL CAPS -> Sentence Case đồng bộ kích cỡ thị giác với 1.2.1.
        self.assertEqual(
            p_h2.text,
            "1.2. Phân tích nhiệm vụ, vị trí lắp đặt để đưa ra mô hình tổng quát của hệ thống và lý do lựa chọn"
        )

        # Heading 3: Giữ nguyên
        self.assertEqual(p_h3.text, "1.2.1. Phân tích nhiệm vụ thiết kế hệ thống mạng")

        # Thống kê
        self.assertGreater(res["stats"]["headings_casing_normalized_count"], 0)

    def test_disabled_casing_normalization(self):
        """Khi tắt normalize_heading_caps=False, Heading 2 ALL CAPS giữ nguyên"""
        doc = Document()
        doc.add_paragraph("CHƯƠNG 1: XÂY DỰNG HỆ THỐNG MẠNG")
        doc.add_paragraph("1.2. PHÂN TÍCH NHIỆM VỤ")

        standardizer = DocumentStandardizer(normalize_heading_caps=False)
        res = standardizer.process(doc, auto_insert_toc=False, evaluate_toc=False)

        self.assertEqual(doc.paragraphs[1].text, "1.2. PHÂN TÍCH NHIỆM VỤ")
        self.assertEqual(res["stats"]["headings_casing_normalized_count"], 0)

    def test_style_manager_hierarchical_indents(self):
        """Kiểm thử StyleManager cấu hình sẵn các style gốc với left_indent phân cấp"""
        doc = Document()
        sm = StyleManager(doc)
        sm.setup_document_styles()

        s_h1 = doc.styles['Heading 1']
        s_h2 = doc.styles['Heading 2']
        s_h3 = doc.styles['Heading 3']
        s_h4 = doc.styles['Heading 4']
        s_normal = doc.styles['Normal']

        self.assertAlmostEqual(s_h1.paragraph_format.left_indent.cm, Config.HEADING_INDENT_PER_LEVEL_CM * 0, places=2)
        self.assertAlmostEqual(s_h2.paragraph_format.left_indent.cm, Config.HEADING_INDENT_PER_LEVEL_CM * 1, places=2)
        self.assertAlmostEqual(s_h3.paragraph_format.left_indent.cm, Config.HEADING_INDENT_PER_LEVEL_CM * 2, places=2)
        self.assertAlmostEqual(s_h4.paragraph_format.left_indent.cm, Config.HEADING_INDENT_PER_LEVEL_CM * 3, places=2)

        # Căn lề Heading 1 là CENTER theo Form mẫu, Heading 2/3/4 là LEFT
        self.assertEqual(s_h1.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.CENTER)
        self.assertEqual(s_h2.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.LEFT)
        self.assertEqual(s_h3.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.LEFT)
        self.assertEqual(s_h4.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.LEFT)

        self.assertAlmostEqual(s_normal.paragraph_format.line_spacing, Config.LINE_SPACING, places=2)
        self.assertEqual(s_normal.paragraph_format.space_after.pt, Config.SPACE_AFTER_PT)

    def test_numbered_heading_without_chapter_keyword_hierarchy(self):
        """
        Kiểm thử đề mục đánh số dạng '1.', '1.1.', '1.1.1.' không có chữ 'Chương':
        - '1. TỔNG QUAN HỆ THỐNG' phải được nhận diện là Heading 1, lề trái 0cm
        - '1.1. Yêu cầu chi tiết' phải được nhận diện là Heading 2, thụt lề 0.5cm
        - '1.1.1. Băng thông mạng' phải được nhận diện là Heading 3, thụt lề 1.0cm
        """
        doc = Document()
        p1 = doc.add_paragraph("1. TỔNG QUAN HỆ THỐNG")
        p1.runs[0].bold = True
        doc.add_paragraph("1.1. Yêu cầu chi tiết")
        doc.add_paragraph("1.1.1. Băng thông mạng")
        doc.add_paragraph("Nội dung thân bài kiểm tra...")

        standardizer = DocumentStandardizer()
        res = standardizer.process(doc, auto_insert_toc=False, evaluate_toc=False, normalize_chapter_to_roman=True)

        stats = res["stats"]
        self.assertEqual(stats["heading_1_count"], 1)
        self.assertEqual(stats["heading_2_count"], 1)
        self.assertEqual(stats["heading_3_count"], 1)

        p_h1 = doc.paragraphs[0]
        p_h2 = doc.paragraphs[1]
        p_h3 = doc.paragraphs[2]

        self.assertEqual(p_h1.style.name, "Heading 1")
        self.assertEqual(p_h2.style.name, "Heading 2")
        self.assertEqual(p_h3.style.name, "Heading 3")

        self.assertAlmostEqual(p_h1.paragraph_format.left_indent.cm, Config.HEADING_INDENT_PER_LEVEL_CM * 0, places=2)
        self.assertAlmostEqual(p_h2.paragraph_format.left_indent.cm, Config.HEADING_INDENT_PER_LEVEL_CM * 1, places=2)
        self.assertAlmostEqual(p_h3.paragraph_format.left_indent.cm, Config.HEADING_INDENT_PER_LEVEL_CM * 2, places=2)

        self.assertEqual(p_h1.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.CENTER)
        self.assertEqual(p_h2.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.LEFT)

    def test_to_sentence_case_homonym_disambiguation(self):
        """Kiểm thử khử nhập nhằng ngữ cảnh giữa từ viết tắt kỹ thuật và từ tiếng Việt (lan, san, ai)"""
        # 'lây lan' vs 'mạng LAN'
        s1 = "NGUY CƠ LÂY LAN CỦA VIRUS TRONG MẠNG LAN"
        self.assertEqual(to_sentence_case(s1), "Nguy cơ lây lan của virus trong mạng LAN")

        # 'san bằng' vs 'mạng SAN'
        s2 = "GIẢI PHÁP SAN BẰNG TẢI TRỌNG CHO MẠNG SAN"
        self.assertEqual(to_sentence_case(s2), "Giải pháp san bằng tải trọng cho mạng SAN")

        # 'ai là' vs 'mô hình AI'
        s3 = "CÂU HỎI AI LÀ NGƯỜI CHỊU TRÁCH NHIỆM VỀ MÔ HÌNH AI"
        self.assertEqual(to_sentence_case(s3), "Câu hỏi ai là người chịu trách nhiệm về mô hình AI")

    def test_to_sentence_case_preserves_it_and_wifi_and_subtitles(self):
        """Kiểm thử bảo tồn IT, WIFI và viết hoa phân đoạn sau dấu hai chấm / gạch ngang"""
        s1 = "KHẢO SÁT HẠ TẦNG PHÒNG IT VÀ MẠNG WIFI"
        self.assertEqual(to_sentence_case(s1), "Khảo sát hạ tầng phòng IT và mạng WIFI")

        s2 = "TỔNG QUAN: NHIỆM VỤ THIẾT KẾ HỆ THỐNG MẠNG"
        self.assertEqual(to_sentence_case(s2), "Tổng quan: Nhiệm vụ thiết kế hệ thống mạng")

        s3 = "PHÒNG 302-A9 – PHÒNG NHÂN SỰ"
        self.assertEqual(to_sentence_case(s3), "Phòng 302-A9 – Phòng nhân sự")

if __name__ == '__main__':
    unittest.main()
