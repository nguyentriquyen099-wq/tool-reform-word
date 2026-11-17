"""
Unit Tests for Roman Numeral Chapter Normalization
Kiểm thử tính năng chuẩn hóa số Chương sang số La Mã (CHƯƠNG I, II, III...)
theo đúng quy chuẩn học thuật Việt Nam.
"""

import unittest
from docx import Document
from config import Config
from core.run_manipulator import RunManipulator, int_to_roman
from core.classifier import WaterfallClassifier, ParagraphType, roman_to_int
from core.pipeline import DocumentStandardizer

class TestChapterRomanNormalization(unittest.TestCase):

    def test_int_to_roman_basic_values(self):
        """Kiểm thử hàm int_to_roman cho các giá trị thông dụng 1 -> 20"""
        expected = {
            1: 'I', 2: 'II', 3: 'III', 4: 'IV', 5: 'V',
            6: 'VI', 7: 'VII', 8: 'VIII', 9: 'IX', 10: 'X',
            11: 'XI', 12: 'XII', 13: 'XIII', 14: 'XIV', 15: 'XV',
            16: 'XVI', 17: 'XVII', 18: 'XVIII', 19: 'XIX', 20: 'XX',
            40: 'XL', 50: 'L', 90: 'XC', 100: 'C', 400: 'CD', 500: 'D', 900: 'CM', 1000: 'M'
        }
        for num, roman in expected.items():
            self.assertEqual(int_to_roman(num), roman, f"Chuyển đổi {num} sang La Mã phải là {roman}")

    def test_int_to_roman_edge_cases(self):
        """Kiểm thử int_to_roman với các trường hợp biên (0, số âm, boolean, chuỗi số, không phải int)"""
        self.assertEqual(int_to_roman(0), "0")
        self.assertEqual(int_to_roman(-1), "-1")
        self.assertEqual(int_to_roman("abc"), "abc") # type: ignore
        self.assertEqual(int_to_roman(None), "None") # type: ignore
        self.assertEqual(int_to_roman(True), "True") # bool không được bị ép thành số 1
        self.assertEqual(int_to_roman(False), "False")
        self.assertEqual(int_to_roman("5"), "V") # chuỗi số hợp lệ
        self.assertEqual(int_to_roman("10"), "X")
        self.assertEqual(int_to_roman(4000), "4000") # vượt giới hạn La Mã tiêu chuẩn

    def test_convert_chapter_single_run(self):
        """Kiểm thử chuyển đổi số Chương trên 1 Run duy nhất"""
        doc = Document()
        p1 = doc.add_paragraph("CHƯƠNG 1: XÂY DỰNG HỆ THỐNG MẠNG")
        p2 = doc.add_paragraph("CHƯƠNG 2: PHÂN CHIA ĐỊA CHỈ SUBNET")
        p3 = doc.add_paragraph("CHƯƠNG 3: TỔNG KẾT")
        p4 = doc.add_paragraph("Chương 4: KẾT LUẬN")
        p5 = doc.add_paragraph("CHƯƠNG 10")

        self.assertTrue(RunManipulator.convert_chapter_number_to_roman(p1))
        self.assertEqual(p1.text, "CHƯƠNG I: XÂY DỰNG HỆ THỐNG MẠNG")

        self.assertTrue(RunManipulator.convert_chapter_number_to_roman(p2))
        self.assertEqual(p2.text, "CHƯƠNG II: PHÂN CHIA ĐỊA CHỈ SUBNET")

        self.assertTrue(RunManipulator.convert_chapter_number_to_roman(p3))
        self.assertEqual(p3.text, "CHƯƠNG III: TỔNG KẾT")

        self.assertTrue(RunManipulator.convert_chapter_number_to_roman(p4))
        self.assertEqual(p4.text, "CHƯƠNG IV: KẾT LUẬN")

        self.assertTrue(RunManipulator.convert_chapter_number_to_roman(p5))
        self.assertEqual(p5.text, "CHƯƠNG X")

    def test_convert_chapter_split_runs_preserves_formatting(self):
        """
        Kiểm thử Phân hệ 5: Thao tác trên Run phân mảnh (Fragmented Runs)
        Bảo tồn tuyệt đối đối tượng Run và định dạng riêng của từng Run
        """
        doc = Document()
        p = doc.add_paragraph()
        r1 = p.add_run("CHƯƠNG ")
        r1.bold = True
        r2 = p.add_run("1")
        r2.bold = True
        r3 = p.add_run(": XÂY DỰNG HỆ THỐNG")
        r3.italic = True

        runs_before_count = len(p.runs)
        changed = RunManipulator.convert_chapter_number_to_roman(p)

        self.assertTrue(changed)
        self.assertEqual(p.text, "CHƯƠNG I: XÂY DỰNG HỆ THỐNG")
        self.assertEqual(len(p.runs), runs_before_count, "Số lượng Run không được bị xóa hoặc thay thế bằng paragraph.text")
        self.assertTrue(p.runs[0].bold, "Run 0 phải giữ nguyên bold=True")
        self.assertTrue(p.runs[2].italic, "Run 2 phải giữ nguyên italic=True")

    def test_convert_chapter_with_leading_whitespace(self):
        """Kiểm thử khi đoạn văn có khoảng trắng ở đầu trên nhiều Run"""
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("   ")
        p.add_run("CHƯƠNG 1: TỔNG QUAN")

        changed = RunManipulator.convert_chapter_number_to_roman(p)
        self.assertTrue(changed)
        self.assertEqual(p.text, "   CHƯƠNG I: TỔNG QUAN")

    def test_already_roman_chapter_unchanged(self):
        """Kiểm thử tiêu đề đã là số La Mã sẵn (CHƯƠNG I, II, III...) thì giữ nguyên không đổi"""
        doc = Document()
        p1 = doc.add_paragraph("CHƯƠNG I: TỔNG QUAN")
        p2 = doc.add_paragraph("CHƯƠNG II: CƠ SỞ KỸ THUẬT")
        p3 = doc.add_paragraph("CHƯƠNG III: THỰC NGHIỆM")

        self.assertFalse(RunManipulator.convert_chapter_number_to_roman(p1))
        self.assertEqual(p1.text, "CHƯƠNG I: TỔNG QUAN")

        self.assertFalse(RunManipulator.convert_chapter_number_to_roman(p2))
        self.assertEqual(p2.text, "CHƯƠNG II: CƠ SỞ KỸ THUẬT")

        self.assertFalse(RunManipulator.convert_chapter_number_to_roman(p3))
        self.assertEqual(p3.text, "CHƯƠNG III: THỰC NGHIỆM")

    def test_subheadings_and_body_untouched(self):
        """
        Đảm bảo các tiểu mục con dạng số Ả Rập (1.1, 1.2, 1.2.1...)
        và văn bản thân (Body text) KHÔNG bị thay đổi
        """
        doc = Document()
        p_sub1 = doc.add_paragraph("1.1. YÊU CẦU THIẾT KẾ HỆ THỐNG MẠNG")
        p_sub2 = doc.add_paragraph("1.2. MÔ HÌNH HỆ THỐNG MẠNG")
        p_sub3 = doc.add_paragraph("1.2.1. Phân tích lưu lượng")
        p_num1 = doc.add_paragraph("1. GIỚI THIỆU CHUNG")
        p_body = doc.add_paragraph("Trong chương 1, chúng ta đã xem xét các thông số kỹ thuật.")

        for p in [p_sub1, p_sub2, p_sub3, p_num1, p_body]:
            orig = p.text
            changed = RunManipulator.convert_chapter_number_to_roman(p)
            self.assertFalse(changed, f"Không được sửa đổi đoạn: {orig}")
            self.assertEqual(p.text, orig)

    def test_pipeline_normalize_chapter_to_roman(self):
        """
        Kiểm thử tích hợp qua Pipeline chuẩn hóa:
        - Tiêu đề Chương chuyển sang số La Mã: CHƯƠNG I, CHƯƠNG II, CHƯƠNG III
        - Các tiểu mục con giữ nguyên dạng số Ả Rập: 1.1, 1.2, 2.1
        - HSM nhận diện đúng phân cấp Heading 1 và Heading 2
        """
        doc = Document()
        doc.add_paragraph("CHƯƠNG 1: XÂY DỰNG HỆ THỐNG MẠNG")
        doc.add_paragraph("1.1. YÊU CẦU THIẾT KẾ HỆ THỐNG MẠNG")
        doc.add_paragraph("Nội dung mô tả yêu cầu thiết kế...")
        doc.add_paragraph("1.2. MÔ HÌNH HỆ THỐNG MẠNG")
        doc.add_paragraph("Nội dung mô tả mô hình mạng...")
        doc.add_paragraph("CHƯƠNG 2: PHÂN CHIA ĐỊA CHỈ SUBNET CHO CÁC PHÒNG BAN")
        doc.add_paragraph("2.1. THIẾT LẬP SUBNET")
        doc.add_paragraph("Nội dung subnet...")
        doc.add_paragraph("CHƯƠNG 3: TỔNG KẾT")
        doc.add_paragraph("Kết luận cuối cùng...")

        standardizer = DocumentStandardizer()
        res = standardizer.process(
            doc,
            auto_insert_toc=False,
            evaluate_toc=False,
            normalize_chapter_to_roman=True
        )
        stats = res["stats"]

        self.assertEqual(stats["chapters_roman_count"], 3, "Phải ghi nhận 3 chương được chuẩn hóa sang La Mã")
        self.assertEqual(stats["heading_1_count"], 3, "Phải nhận diện đủ 3 Heading 1")
        self.assertEqual(stats["heading_2_count"], 3, "Phải nhận diện đủ 3 Heading 2 (1.1, 1.2, 2.1)")

        # Kiểm tra nội dung text sau khi xử lý (Chương sang La Mã, Heading 2 chuẩn hóa sang Sentence Case)
        paras = doc.paragraphs
        self.assertEqual(paras[0].text, "CHƯƠNG I: XÂY DỰNG HỆ THỐNG MẠNG")
        self.assertEqual(paras[1].text, "1.1. Yêu cầu thiết kế hệ thống mạng")
        self.assertEqual(paras[3].text, "1.2. Mô hình hệ thống mạng")
        self.assertEqual(paras[5].text, "CHƯƠNG II: PHÂN CHIA ĐỊA CHỈ SUBNET CHO CÁC PHÒNG BAN")
        self.assertEqual(paras[6].text, "2.1. Thiết lập subnet")
        self.assertEqual(paras[8].text, "CHƯƠNG III: TỔNG KẾT")

        # Kiểm tra Style của các đoạn
        self.assertEqual(paras[0].style.name, "Heading 1")
        self.assertEqual(paras[1].style.name, "Heading 2")
        self.assertEqual(paras[5].style.name, "Heading 1")
        self.assertEqual(paras[6].style.name, "Heading 2")
        self.assertEqual(paras[8].style.name, "Heading 1")

    def test_pipeline_disabled_option_leaves_arabic(self):
        """Kiểm thử khi tắt tùy chọn normalize_chapter_to_roman=False thì giữ nguyên số Ả Rập"""
        doc = Document()
        doc.add_paragraph("CHƯƠNG 1: TỔNG QUAN")
        doc.add_paragraph("1.1. Chi tiết")

        standardizer = DocumentStandardizer()
        res = standardizer.process(
            doc,
            auto_insert_toc=False,
            evaluate_toc=False,
            normalize_chapter_to_roman=False
        )
        stats = res["stats"]

        self.assertEqual(stats["chapters_roman_count"], 0)
        self.assertEqual(doc.paragraphs[0].text, "CHƯƠNG 1: TỔNG QUAN")

    def test_pipeline_override_instance_true_with_process_false(self):
        """
        Kiểm thử DocumentStandardizer được tạo với normalize_chapter_to_roman=True,
        nhưng khi gọi process(normalize_chapter_to_roman=False) thì bắt buộc phải giữ nguyên số Ả Rập.
        """
        doc = Document()
        doc.add_paragraph("CHƯƠNG 1: TỔNG QUAN")
        standardizer = DocumentStandardizer(normalize_chapter_to_roman=True)
        res = standardizer.process(
            doc,
            auto_insert_toc=False,
            evaluate_toc=False,
            normalize_chapter_to_roman=False
        )
        self.assertEqual(res["stats"]["chapters_roman_count"], 0)
        self.assertEqual(doc.paragraphs[0].text, "CHƯƠNG 1: TỔNG QUAN")

    def test_convert_chapter_nfd_unicode(self):
        """Kiểm thử RunManipulator xử lý tiêu đề có mã Unicode NFD (tổ hợp)"""
        import unicodedata
        doc = Document()
        nfd_text = unicodedata.normalize('NFD', "CHƯƠNG 1: XÂY DỰNG HỆ THỐNG MẠNG")
        p = doc.add_paragraph(nfd_text)
        changed = RunManipulator.convert_chapter_number_to_roman(p)
        self.assertTrue(changed)
        self.assertEqual(p.text, "CHƯƠNG I: XÂY DỰNG HỆ THỐNG MẠNG")

    def test_pipeline_nfd_unicode_chapters(self):
        """Kiểm thử toàn trình pipeline nhận diện và chuyển đổi chương có mã Unicode NFD (tổ hợp)"""
        import unicodedata
        doc = Document()
        nfd_title = unicodedata.normalize('NFD', "CHƯƠNG 1: KIẾN TRÚC MẠNG")
        doc.add_paragraph(nfd_title)
        doc.add_paragraph("1.1. Khảo sát")
        standardizer = DocumentStandardizer()
        res = standardizer.process(
            doc,
            auto_insert_toc=False,
            evaluate_toc=False,
            normalize_chapter_to_roman=True
        )
        self.assertEqual(res["stats"]["chapters_roman_count"], 1)
        self.assertEqual(doc.paragraphs[0].text, "CHƯƠNG I: KIẾN TRÚC MẠNG")
        self.assertEqual(doc.paragraphs[1].text, "1.1. Khảo sát")

if __name__ == "__main__":
    unittest.main()
