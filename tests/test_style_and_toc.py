"""
Unit tests cho StyleManager & TOCManager (Lưu ý 2, 4 & Phân hệ 6)
"""

import unittest
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from config import Config
from core.style_manager import StyleManager
from core.toc_manager import TOCManager

class TestStyleAndTOC(unittest.TestCase):
    def test_heading_1_unnumbered_has_outline_lvl_0(self):
        doc = Document()
        sm = StyleManager(doc)
        sm.setup_document_styles()

        style = doc.styles[Config.STYLE_HEADING_1_UNNUMBERED]
        pPr = style._element.get_or_add_pPr()
        outline = pPr.find(qn('w:outlineLvl'))
        self.assertIsNotNone(outline)
        # Bắt buộc outlineLevel = 0 để vào TOC
        self.assertEqual(outline.get(qn('w:val')), '0')

    def test_non_toc_heading_has_no_outline_lvl(self):
        doc = Document()
        sm = StyleManager(doc)
        sm.setup_document_styles()

        style = doc.styles[Config.STYLE_NON_TOC_HEADING]
        pPr = style._element.get_or_add_pPr()
        outline = pPr.find(qn('w:outlineLvl'))
        # Bắt buộc KHÔNG CÓ outlineLvl để không vào TOC
        self.assertIsNone(outline)

    def test_enable_update_fields_on_open(self):
        doc = Document()
        res = TOCManager.enable_update_fields_on_open(doc)
        self.assertTrue(res)

        settings_elem = doc.settings.element
        update_fields = settings_elem.find(qn('w:updateFields'))
        self.assertIsNotNone(update_fields)
        self.assertEqual(update_fields.get(qn('w:val')), 'true')

    def test_standardize_page_margins(self):
        doc = Document()
        TOCManager.standardize_page_margins(doc)
        for s in doc.sections:
            self.assertAlmostEqual(s.top_margin.cm, Config.MARGIN_TOP_CM, places=2)
            self.assertAlmostEqual(s.bottom_margin.cm, Config.MARGIN_BOTTOM_CM, places=2)
            self.assertAlmostEqual(s.left_margin.cm, Config.MARGIN_LEFT_CM, places=2)
            self.assertAlmostEqual(s.right_margin.cm, Config.MARGIN_RIGHT_CM, places=2)

    def test_paragraph_level_outline_applied(self):
        """Kiểm thử outlineLvl được gán chính xác trên paragraph XML khi chuẩn hóa"""
        from core.pipeline import DocumentStandardizer
        doc = Document()
        p_toc = doc.add_paragraph("LỜI MỞ ĐẦU")
        p_non_toc = doc.add_paragraph("LỜI CẢM ƠN")

        standardizer = DocumentStandardizer()
        standardizer.process(doc)

        # LỜI MỞ ĐẦU phải có outlineLvl = 0 trực tiếp trên paragraph
        outline_toc = p_toc._element.get_or_add_pPr().find(qn('w:outlineLvl'))
        self.assertIsNotNone(outline_toc)
        self.assertEqual(outline_toc.get(qn('w:val')), '0')

        # LỜI CẢM ƠN không được có outlineLvl trên paragraph
        outline_non_toc = p_non_toc._element.get_or_add_pPr().find(qn('w:outlineLvl'))
        self.assertIsNone(outline_non_toc)

    def test_auto_insert_toc_when_missing(self):
        """Kiểm thử tự động phát hiện và chèn trang Mục lục OpenXML chuẩn khi tài liệu chưa có"""
        from core.pipeline import DocumentStandardizer
        doc = Document()
        doc.add_paragraph("TRANG BÌA BÁO CÁO")
        # Section break
        p_cover = doc.paragraphs[-1]._element.get_or_add_pPr()
        p_cover.append(OxmlElement('w:sectPr'))

        p_intro = doc.add_paragraph("LỜI NÓI ĐẦU")
        doc.add_paragraph("Nội dung mở đầu của báo cáo thực nghiệm...")
        p_task = doc.add_paragraph("I. PHÂN CÔNG NHIỆM VỤ")
        doc.add_paragraph("Thành viên 1: Thiết kế mạng, Thành viên 2: Cấu hình...")
        p_ch1 = doc.add_paragraph("CHƯƠNG 1: TỔNG QUAN HỆ THỐNG")
        doc.add_paragraph("Nội dung chương 1...")

        standardizer = DocumentStandardizer()
        res = standardizer.process(doc)

        self.assertTrue(res["stats"]["toc_inserted"])
        self.assertTrue(TOCManager.has_toc_field(doc))

        # Tìm paragraph MỤC LỤC
        toc_heading_info = TOCManager.find_toc_heading(doc)
        self.assertIsNotNone(toc_heading_info)
        idx_toc, p_toc = toc_heading_info

        # Kiểm tra tiêu đề MỤC LỤC định dạng Non-TOC Heading, Center, không outlineLvl, page_break_before
        self.assertEqual(p_toc.text.strip(), "MỤC LỤC")
        self.assertEqual(p_toc.style.name, Config.STYLE_NON_TOC_HEADING)
        self.assertTrue(p_toc.paragraph_format.page_break_before)
        pPr = p_toc._element.find(qn('w:pPr'))
        self.assertIsNone(pPr.find(qn('w:outlineLvl')))

        # Đoạn ngay sau MỤC LỤC phải là đoạn chứa TOC field
        p_field = doc.paragraphs[idx_toc + 1]
        self.assertTrue(TOCManager.is_toc_paragraph(p_field))
        flds = p_field._element.xpath('.//w:fldSimple')
        self.assertGreater(len(flds), 0)
        self.assertIn('TOC', flds[0].get(qn('w:instr'), ''))

        # Chương 1 phải nằm sau Mục lục và có ngắt trang page_break_before = True
        idx_ch1 = None
        for i, p in enumerate(doc.paragraphs):
            if "CHƯƠNG 1" in p.text or "CHƯƠNG I" in p.text:
                idx_ch1 = i
                break
        self.assertIsNotNone(idx_ch1)
        self.assertGreater(idx_ch1, idx_toc + 1)
        self.assertTrue(doc.paragraphs[idx_ch1].paragraph_format.page_break_before)

    def test_cleanup_fake_manual_toc_entries(self):
        """Kiểm thử phát hiện và dọn dẹp các dòng chấm chấm gõ tay thủ công của sinh viên"""
        from core.pipeline import DocumentStandardizer
        doc = Document()
        p_toc = doc.add_paragraph("MỤC LỤC")
        # Các dòng mục lục chấm chấm thủ công
        doc.add_paragraph("LỜI NÓI ĐẦU .................................................... 1")
        doc.add_paragraph("CHƯƠNG 1: TỔNG QUAN ............................................ 2")
        doc.add_paragraph("1.1. Khái niệm và mục tiêu . . . . . . . . . . . . . . . . . . 3")
        doc.add_paragraph("CHƯƠNG 2: THỰC NGHIỆM \t\t\t 15")
        # Đoạn nội dung thật
        doc.add_paragraph("CHƯƠNG 1: TỔNG QUAN")
        doc.add_paragraph("Đây là nội dung thân bài thật của chương 1...")

        standardizer = DocumentStandardizer()
        res = standardizer.process(doc)

        self.assertTrue(res["stats"]["toc_inserted"])
        self.assertGreaterEqual(res["stats"]["toc_cleaned_count"], 4)

        # Đảm bảo không còn dòng chấm chấm nào trong văn bản
        for p in doc.paragraphs:
            self.assertFalse(TOCManager.is_fake_toc_entry(p))

        # Đảm bảo đã có trường TOC tự động
        self.assertTrue(TOCManager.has_toc_field(doc))

    def test_already_has_toc_field_no_duplicate(self):
        """Kiểm thử tài liệu đã có sẵn trường TOC chuẩn thì không bị chèn trùng lặp"""
        from core.pipeline import DocumentStandardizer
        doc = Document()
        p_toc = doc.add_paragraph("MỤC LỤC")
        p_field = doc.add_paragraph()
        TOCManager.create_toc_field(p_field, use_complex_field=False)
        p_ch1 = doc.add_paragraph("CHƯƠNG 1: TỔNG QUAN")
        doc.add_paragraph("Nội dung chương 1...")

        initial_paragraphs = len(doc.paragraphs)
        standardizer = DocumentStandardizer()
        res = standardizer.process(doc)

        self.assertFalse(res["stats"]["toc_inserted"])
        # Số đoạn văn không bị nhân đôi
        toc_headings = [p for p in doc.paragraphs if p.text.strip().upper() == "MỤC LỤC"]
        self.assertEqual(len(toc_headings), 1)

    def test_auto_insert_toc_disabled_flag(self):
        """Kiểm thử khi tắt cờ auto_insert_toc thì hệ thống không tự chèn Mục lục"""
        from core.pipeline import DocumentStandardizer
        doc = Document()
        doc.add_paragraph("CHƯƠNG 1: TỔNG QUAN")
        doc.add_paragraph("Nội dung chương 1...")
        doc.add_paragraph("1.1. Tiểu mục")
        doc.add_paragraph("Thân bài...")

        standardizer = DocumentStandardizer()
        res = standardizer.process(doc, auto_insert_toc=False)

        self.assertFalse(res["stats"]["toc_inserted"])
        self.assertFalse(TOCManager.has_toc_field(doc))

    def test_complex_field_toc_generation(self):
        """Kiểm thử tạo complex field TOC (w:fldChar begin/separate/end)"""
        doc = Document()
        p = doc.add_paragraph()
        TOCManager.create_toc_field(p, use_complex_field=True)

        elem = p._element
        fld_chars = elem.xpath('.//w:fldChar')
        self.assertEqual(len(fld_chars), 3)  # begin, separate, end
        instr_texts = elem.xpath('.//w:instrText')
        self.assertEqual(len(instr_texts), 1)
        self.assertIn("TOC", instr_texts[0].text)
        self.assertTrue(TOCManager.is_toc_paragraph(p))

    def test_cleanup_fake_entries_with_multiple_blank_lines(self):
        """Kiểm thử dọn dẹp fake TOC có nhiều dòng trắng liên tiếp kẹp giữa"""
        doc = Document()
        p_toc = doc.add_paragraph("MỤC LỤC")
        doc.add_paragraph("CHƯƠNG 1: GIỚI THIỆU .................................... 1")
        doc.add_paragraph("")
        doc.add_paragraph("")
        doc.add_paragraph("CHƯƠNG 2: NỘI DUNG ....................................... 5")
        doc.add_paragraph("CHƯƠNG 1: THẬT")
        doc.add_paragraph("Nội dung thân bài...")

        cleaned = TOCManager.cleanup_fake_toc_entries(doc, 0)
        self.assertEqual(cleaned, 2)
        remaining_texts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        self.assertIn("MỤC LỤC", remaining_texts)
        self.assertIn("CHƯƠNG 1: THẬT", remaining_texts)
        self.assertNotIn("CHƯƠNG 1: GIỚI THIỆU .................................... 1", remaining_texts)
        self.assertNotIn("CHƯƠNG 2: NỘI DUNG ....................................... 5", remaining_texts)

    def test_insertion_point_does_not_match_bibliography(self):
        """Kiểm thử không bị nhận nhầm DANH MỤC TÀI LIỆU THAM KHẢO ở cuối làm điểm chèn Mục lục"""
        doc = Document()
        doc.add_paragraph("TRANG BÌA")
        doc.add_paragraph("CHƯƠNG 1: KHỞI ĐẦU")
        doc.add_paragraph("Nội dung thân bài...")
        doc.add_paragraph("DANH MỤC TÀI LIỆU THAM KHẢO")
        doc.add_paragraph("[1] Sách giáo trình...")

        ins = TOCManager.find_insertion_point(doc, cover_end_idx=1)
        self.assertIsNotNone(ins)
        ins_idx, p_target = ins
        # Phải chọn trước CHƯƠNG 1 (index 1), TUYỆT ĐỐI không chọn DANH MỤC TÀI LIỆU THAM KHẢO (index 3)
        self.assertEqual(ins_idx, 1)
        self.assertIn("CHƯƠNG 1", p_target.text)

    def test_insertion_point_with_roman_numeral_heading(self):
        """Kiểm thử chèn Mục lục trước tiêu đề số La Mã (I. TỔNG QUAN) khi tài liệu không dùng chữ CHƯƠNG"""
        doc = Document()
        doc.add_paragraph("TRANG BÌA")
        doc.add_paragraph("I. TỔNG QUAN VỀ DỰ ÁN")
        doc.add_paragraph("Nội dung phần I...")
        doc.add_paragraph("II. NỘI DUNG THỰC HIỆN")

        ins = TOCManager.find_insertion_point(doc, cover_end_idx=1)
        self.assertIsNotNone(ins)
        ins_idx, p_target = ins
        self.assertEqual(ins_idx, 1)
        self.assertIn("I. TỔNG QUAN", p_target.text)

    def test_insertion_point_with_front_matter_list(self):
        """Kiểm thử chèn Mục lục trước Danh mục hình ảnh / Danh mục bảng (Front matter)"""
        doc = Document()
        doc.add_paragraph("TRANG BÌA")
        doc.add_paragraph("LỜI CẢM ƠN")
        doc.add_paragraph("DANH MỤC HÌNH ẢNH")
        doc.add_paragraph("CHƯƠNG 1: TỔNG QUAN")

        ins = TOCManager.find_insertion_point(doc, cover_end_idx=1)
        self.assertIsNotNone(ins)
        ins_idx, p_target = ins
        self.assertEqual(ins_idx, 2)
        self.assertIn("DANH MỤC HÌNH ẢNH", p_target.text)

    def test_cover_page_spacing_paragraphs_preserved(self):
        """Kiểm thử không xóa nhầm các đoạn trống của Trang bìa khi dọn dẹp khoảng trống trước insertion point"""
        doc = Document()
        doc.add_paragraph("TRANG BÌA DÒNG 1")
        doc.add_paragraph("")  # Dòng trống trong trang bìa
        doc.add_paragraph("TRANG BÌA DÒNG 2")
        # Giả sử cover_end_idx = 3
        doc.add_paragraph("CHƯƠNG 1: BẮT ĐẦU")
        doc.add_paragraph("Nội dung...")

        initial_len = len(doc.paragraphs)
        res = TOCManager.ensure_table_of_contents(doc, cover_end_idx=3)
        self.assertTrue(res["inserted"])
        # Dòng trống ở index 1 của trang bìa vẫn phải còn nguyên
        self.assertEqual(doc.paragraphs[1].text, "")
        self.assertEqual(doc.paragraphs[0].text, "TRANG BÌA DÒNG 1")
        self.assertEqual(doc.paragraphs[2].text, "TRANG BÌA DÒNG 2")

    def test_find_word_executable(self):
        """Kiểm thử hàm tìm kiếm Microsoft Word trên hệ điều hành Windows"""
        word_path = TOCManager.find_word_executable()
        # Trên máy Windows có cài Word, kết quả phải là đường dẫn thực tế hoặc None nếu không có
        if word_path:
            import os
            self.assertTrue(os.path.isfile(word_path))
            self.assertTrue(word_path.lower().endswith("winword.exe"))

    def test_heading_styles_configured_according_to_form_template(self):
        """Kiểm thử cấu hình Heading Styles theo Form mẫu: Heading 1 (16pt, bold), Heading 2 (14pt, bold), Heading 3 (14pt, bold+italic)"""
        doc = Document()
        sm = StyleManager(doc)
        sm.setup_document_styles()

        level_1_styles = [
            'Heading 1',
            Config.STYLE_HEADING_1_UNNUMBERED,
            Config.STYLE_NON_TOC_HEADING
        ]

        for s_name in level_1_styles:
            style = doc.styles[s_name]
            self.assertEqual(style.font.size.pt, Config.HEADING_1_FONT_SIZE, f"Style {s_name} phải có cỡ chữ {Config.HEADING_1_FONT_SIZE}pt")
            self.assertTrue(style.font.bold, f"Style {s_name} phải in đậm")
            self.assertFalse(style.font.italic, f"Style {s_name} phải là chữ đứng (italic=False)")
            rPr = style._element.get_or_add_rPr()
            sz_cs = rPr.find(qn('w:szCs'))
            self.assertIsNotNone(sz_cs, f"Style {s_name} phải có w:szCs")
            self.assertEqual(sz_cs.get(qn('w:val')), str(int(round(Config.HEADING_1_FONT_SIZE * 2))))

        h2_style = doc.styles['Heading 2']
        self.assertEqual(h2_style.font.size.pt, Config.HEADING_2_FONT_SIZE)
        self.assertTrue(h2_style.font.bold)
        self.assertFalse(h2_style.font.italic)

        h3_style = doc.styles['Heading 3']
        self.assertEqual(h3_style.font.size.pt, Config.HEADING_3_FONT_SIZE)
        self.assertTrue(h3_style.font.bold)
        self.assertTrue(h3_style.font.italic, "Heading 3 phải in nghiêng theo Form mẫu chuẩn")

    def test_pipeline_formats_headings_according_to_form_template(self):
        """Kiểm thử khi xử lý tài liệu, Heading 1 đạt 16pt, Heading 2 đạt 14pt bold, Heading 3 đạt 14pt bold+italic theo Form mẫu"""
        from core.pipeline import DocumentStandardizer
        doc = Document()
        p1 = doc.add_paragraph()
        r1 = p1.add_run("CHƯƠNG 1: TỔNG QUAN")
        r1.italic = True
        r1.font.size = Pt(12)

        p2 = doc.add_paragraph()
        r2 = p2.add_run("1.1. CƠ SỞ LÝ THUYẾT")
        r2.italic = True
        r2.font.size = Pt(13)

        p3 = doc.add_paragraph()
        r3 = p3.add_run("1.1.1. Khái niệm cơ bản")
        r3.italic = False
        r3.font.size = Pt(11)

        p4 = doc.add_paragraph()
        r4 = p4.add_run("1.1.1.1. Chi tiết định nghĩa")
        r4.font.size = Pt(10)

        # Body text có chữ in nghiêng
        p_body = doc.add_paragraph()
        r_b1 = p_body.add_run("Đây là đoạn văn bản thông thường có từ ")
        r_b2 = p_body.add_run("in nghiêng quan trọng")
        r_b2.italic = True
        r_b3 = p_body.add_run(" cần lưu ý.")

        standardizer = DocumentStandardizer()
        standardizer.process(doc, auto_insert_toc=False)

        # Kiểm tra p1 (Heading 1: 16pt bold, đứng)
        for r in p1.runs:
            self.assertEqual(r.font.size.pt, Config.HEADING_1_FONT_SIZE)
            self.assertTrue(r.bold)
            self.assertFalse(r.italic)

        # Kiểm tra p2 (Heading 2: 14pt bold, đứng)
        for r in p2.runs:
            self.assertEqual(r.font.size.pt, Config.HEADING_2_FONT_SIZE)
            self.assertTrue(r.bold)
            self.assertFalse(r.italic)

        # Kiểm tra p3 (Heading 3: 14pt bold, nghiêng)
        for r in p3.runs:
            self.assertEqual(r.font.size.pt, Config.HEADING_3_FONT_SIZE)
            self.assertTrue(r.bold)
            self.assertTrue(r.italic)

        # Kiểm tra body text vẫn giữ nguyên in nghiêng
        self.assertTrue(r_b2.italic, "Body text in nghiêng nội tại phải được bảo tồn")

    def test_linked_character_styles_synchronized(self):
        """Kiểm thử khi tài liệu có sẵn linked character styles (Heading 2 Char, Heading 3 Char) với size cũ 18pt/13.5pt, hệ thống tự động đồng bộ về 14pt upright"""
        from docx.enum.style import WD_STYLE_TYPE
        doc = Document()
        # Lấy hoặc tạo character style
        try:
            c_style = doc.styles['Heading 2 Char']
        except KeyError:
            c_style = doc.styles.add_style('Heading 2 Char', WD_STYLE_TYPE.CHARACTER)
        c_style.font.size = Pt(18)
        c_style.font.italic = True
        c_rPr = c_style._element.get_or_add_rPr()
        sz_cs = c_rPr.find(qn('w:szCs'))
        if sz_cs is None:
            sz_cs = OxmlElement('w:szCs')
            c_rPr.append(sz_cs)
        sz_cs.set(qn('w:val'), '36')
        if c_rPr.find(qn('w:iCs')) is None:
            c_rPr.append(OxmlElement('w:iCs'))

        sm = StyleManager(doc)
        sm.setup_document_styles()

        updated_c = doc.styles['Heading 2 Char']
        self.assertEqual(updated_c.font.size.pt, 14.0)
        self.assertTrue(updated_c.font.bold)
        self.assertFalse(updated_c.font.italic)
        self.assertIsNone(updated_c._element.get_or_add_rPr().find(qn('w:iCs')))
        self.assertEqual(updated_c._element.get_or_add_rPr().find(qn('w:szCs')).get(qn('w:val')), '28')

    def test_normal_style_configured_14pt_and_xml_sync(self):
        """Kiểm thử Normal style được cấu hình chuẩn 14pt và đồng bộ w:sz=28, w:szCs=28"""
        doc = Document()
        sm = StyleManager(doc)
        sm.setup_document_styles()

        normal_style = doc.styles['Normal']
        self.assertEqual(Config.BODY_FONT_SIZE, 14.0)
        self.assertEqual(normal_style.font.size.pt, 14.0)
        self.assertEqual(normal_style.font.name, Config.FONT_NAME)

        rPr = normal_style._element.get_or_add_rPr()
        sz = rPr.find(qn('w:sz'))
        self.assertIsNotNone(sz, "Normal style phải có thẻ w:sz")
        self.assertEqual(sz.get(qn('w:val')), '28', "w:sz của Normal style phải là 28 (14pt)")

        szCs = rPr.find(qn('w:szCs'))
        self.assertIsNotNone(szCs, "Normal style phải có thẻ w:szCs")
        self.assertEqual(szCs.get(qn('w:val')), '28', "w:szCs của Normal style phải là 28 (14pt)")

    def test_body_and_list_inherit_14pt_and_cover_preserved(self):
        """Kiểm thử toàn diện: Body text và List items kế thừa chuẩn 14pt, bìa giữ nguyên định dạng gốc"""
        from core.pipeline import DocumentStandardizer
        doc = Document()

        # 1. Trang bìa (Cover page)
        p_cov1 = doc.add_paragraph("TRƯỜNG ĐẠI HỌC BÁCH KHOA HÀ NỘI")
        r_c1 = p_cov1.runs[0]
        r_c1.font.size = Pt(13)
        r_c1.bold = True

        p_cov2 = doc.add_paragraph("BÁO CÁO TỐT NGHIỆP KỸ SƯ")
        r_c2 = p_cov2.runs[0]
        r_c2.font.size = Pt(18)
        r_c2.bold = True

        p_cov3 = doc.add_paragraph("Hà Nội, 2026")
        r_c3 = p_cov3.runs[0]
        r_c3.font.size = Pt(12)

        # Section break kết thúc bìa
        p_cov3._element.get_or_add_pPr().append(OxmlElement('w:sectPr'))

        # 2. Vùng nội dung: Heading, Body text, List items
        p_intro = doc.add_paragraph("LỜI MỞ ĐẦU")
        p_body1 = doc.add_paragraph("Đây là đoạn văn xuôi mở đầu. Có cụm từ ")
        r_b_bold = p_body1.add_run("in đậm quan trọng")
        r_b_bold.bold = True
        p_body1.add_run(" và cụm từ ")
        r_b_italic = p_body1.add_run("in nghiêng đặc biệt")
        r_b_italic.italic = True
        p_body1.add_run(" để kiểm tra.")

        p_list1 = doc.add_paragraph("• Mục danh sách cấp 1 thứ nhất")
        p_list2 = doc.add_paragraph("+ Mục danh sách cấp 2 chi tiết")

        p_ch1 = doc.add_paragraph("CHƯƠNG 1: TỔNG QUAN")
        p_body2 = doc.add_paragraph("Nội dung thân bài của chương 1 cũng thuộc kiểu Body text 14pt.")

        standardizer = DocumentStandardizer()
        res = standardizer.process(doc, auto_insert_toc=False)

        # Kiểm tra bìa giữ nguyên size gốc
        self.assertEqual(r_c1.font.size.pt, 13)
        self.assertEqual(r_c2.font.size.pt, 18)
        self.assertEqual(r_c3.font.size.pt, 12)

        # Kiểm tra Normal style của doc đã là 14pt
        self.assertEqual(doc.styles['Normal'].font.size.pt, 14.0)

        # Kiểm tra p_body1 và p_body2 có style là Normal
        self.assertEqual(p_body1.style.name, 'Normal')
        self.assertEqual(p_body2.style.name, 'Normal')

        # Kiểm tra run-level của body text: không bị ghi đè sz rác, kế thừa trọn vẹn Normal (14pt)
        for r in p_body1.runs:
            self.assertIsNone(r.font.size)
        self.assertTrue(r_b_bold.bold)
        self.assertTrue(r_b_italic.italic)

        # Kiểm tra List items
        self.assertEqual(p_list1.style.name, 'Normal')
        self.assertEqual(p_list2.style.name, 'Normal')
        self.assertTrue(p_list1.text.startswith("- "))
        self.assertTrue(p_list2.text.startswith("+ "))
        for r in p_list1.runs:
            self.assertIsNone(r.font.size, "List item cấp 1 phải kế thừa chuẩn cỡ chữ từ Normal")
        for r in p_list2.runs:
            self.assertIsNone(r.font.size, "List item cấp 2 phải kế thừa chuẩn cỡ chữ từ Normal")

if __name__ == '__main__':
    unittest.main()
