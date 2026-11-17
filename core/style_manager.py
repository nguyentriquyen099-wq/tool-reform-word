"""
Quản lý Style tài liệu Word (.docx)
Giải quyết Phân hệ 5 & Lưu ý kỹ thuật số 2:
- Tách riêng Style cho "Heading không đánh số"
  + Heading 1 Unnumbered: có outlineLevel = 0 để tự động vào TOC
  + Non-TOC Heading: định dạng như Heading 1 nhưng KHÔNG có outlineLevel để không vào TOC
- Chuẩn hóa các Style gốc (Normal, Heading 1, 2, 3, Caption) trong styles.xml
- Đảm bảo thiết lập đầy đủ w:rFonts (ascii, hAnsi, cs, eastAsia) cho tiếng Việt Unicode
"""

import logging
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
from config import Config

logger = logging.getLogger(__name__)

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

class StyleManager:
    def __init__(self, doc):
        self.doc = doc
        self.styles = doc.styles

    def setup_document_styles(self):
        """
        Cấu hình toàn diện các Style gốc của tài liệu
        """
        self._setup_normal_style()
        self._setup_heading_1()
        self._setup_heading_2()
        self._setup_heading_3()
        self._setup_heading_4()
        self._setup_heading_1_unnumbered()
        self._setup_non_toc_heading()
        self._setup_caption_style()

    def _set_rfonts(self, rPr_elem, font_name: str):
        """
        Gán đầy đủ w:rFonts bao gồm ascii, hAnsi, cs, eastAsia để hỗ trợ tiếng Việt Unicode
        """
        if rPr_elem is None:
            return
        rfonts = rPr_elem.find(qn('w:rFonts'))
        if rfonts is None:
            rfonts = OxmlElement('w:rFonts')
            rPr_elem.insert(0, rfonts)
        rfonts.set(qn('w:ascii'), font_name)
        rfonts.set(qn('w:hAnsi'), font_name)
        rfonts.set(qn('w:cs'), font_name)
        rfonts.set(qn('w:eastAsia'), font_name)

    def _apply_heading_font_properties(self, style, font_size_pt: float, bold: bool = True, italic: bool = False):
        """
        Cấu hình toàn diện thuộc tính Font cho Style đề mục (Heading) hoặc Caption:
        - Font: Times New Roman, màu đen (#000000)
        - Cỡ chữ chuẩn: 14pt (hoặc theo cấu hình)
        - In đậm (bold=True) và Chữ đứng (italic=False)
        - Gán đầy đủ w:rFonts (ascii, hAnsi, cs, eastAsia)
        - Đồng bộ thuộc tính Complex Scripts (w:szCs, w:bCs, xóa w:iCs)
        - Tự động đồng bộ các linked character styles nếu có (ví dụ: 'Heading 2 Char')
        """
        font = style.font
        font.name = Config.FONT_NAME
        font.size = Pt(font_size_pt)
        font.bold = bold
        font.italic = italic
        font.color.rgb = RGBColor(0, 0, 0)

        rPr = style._element.get_or_add_rPr()
        self._set_rfonts(rPr, Config.FONT_NAME)

        half_pts = str(int(round(font_size_pt * 2)))

        # Đồng bộ cỡ chữ trực tiếp w:sz và w:szCs (half-points: 14pt = 28)
        sz = rPr.find(qn('w:sz'))
        if sz is None:
            sz = OxmlElement('w:sz')
            rPr.append(sz)
        sz.set(qn('w:val'), half_pts)

        # Đồng bộ cỡ chữ Complex Scripts (w:szCs)
        sz_cs = rPr.find(qn('w:szCs'))
        if sz_cs is None:
            sz_cs = OxmlElement('w:szCs')
            rPr.append(sz_cs)
        sz_cs.set(qn('w:val'), half_pts)

        # Đồng bộ in đậm cho Complex Scripts (w:bCs)
        if bold:
            if rPr.find(qn('w:bCs')) is None:
                rPr.append(OxmlElement('w:bCs'))
        else:
            b_cs = rPr.find(qn('w:bCs'))
            if b_cs is not None:
                rPr.remove(b_cs)

        # Đồng bộ in nghiêng (w:i và w:iCs)
        if not italic:
            i_cs = rPr.find(qn('w:iCs'))
            if i_cs is not None:
                rPr.remove(i_cs)
            i_elem = rPr.find(qn('w:i'))
            if i_elem is not None:
                i_elem.set(qn('w:val'), '0')
        else:
            if rPr.find(qn('w:iCs')) is None:
                rPr.append(OxmlElement('w:iCs'))

        # Đồng bộ với linked character style nếu có (ví dụ: Heading 1 Char, Heading 2 Char,...)
        char_style_names = [f"{style.name} Char", f"{style.name}Char"]
        for s in self.styles:
            if s.type == WD_STYLE_TYPE.CHARACTER and s.name in char_style_names:
                try:
                    s_font = s.font
                    s_font.name = Config.FONT_NAME
                    s_font.size = Pt(font_size_pt)
                    s_font.bold = bold
                    s_font.italic = italic
                    s_font.color.rgb = RGBColor(0, 0, 0)
                    c_rPr = s._element.get_or_add_rPr()
                    self._set_rfonts(c_rPr, Config.FONT_NAME)
                    c_sz = c_rPr.find(qn('w:sz'))
                    if c_sz is None:
                        c_sz = OxmlElement('w:sz')
                        c_rPr.append(c_sz)
                    c_sz.set(qn('w:val'), half_pts)
                    c_sz_cs = c_rPr.find(qn('w:szCs'))
                    if c_sz_cs is None:
                        c_sz_cs = OxmlElement('w:szCs')
                        c_rPr.append(c_sz_cs)
                    c_sz_cs.set(qn('w:val'), half_pts)
                    if bold:
                        if c_rPr.find(qn('w:bCs')) is None:
                            c_rPr.append(OxmlElement('w:bCs'))
                    if not italic:
                        c_i_cs = c_rPr.find(qn('w:iCs'))
                        if c_i_cs is not None:
                            c_rPr.remove(c_i_cs)
                        c_i = c_rPr.find(qn('w:i'))
                        if c_i is not None:
                            c_i.set(qn('w:val'), '0')
                except Exception as e:
                    logger.debug(f"Lỗi đồng bộ linked character style {s.name}: {e}")

    def _setup_normal_style(self):
        """
        Style Normal (Body text): Times New Roman, 14pt, Justify, Spacing 1.3
        Đồng bộ đầy đủ thẻ XML w:sz="28" và w:szCs="28" cho tiếng Việt Unicode
        """
        try:
            style = self.styles['Normal']
            font = style.font
            font.name = Config.FONT_NAME
            font.size = Pt(Config.BODY_FONT_SIZE)
            font.color.rgb = RGBColor(0, 0, 0)

            # Đảm bảo rFonts trong XML
            rPr = style._element.get_or_add_rPr()
            self._set_rfonts(rPr, Config.FONT_NAME)

            # Đồng bộ cỡ chữ w:sz và w:szCs (half-points: 14pt = 28)
            half_pts = str(int(round(Config.BODY_FONT_SIZE * 2)))
            sz = rPr.find(qn('w:sz'))
            if sz is None:
                sz = OxmlElement('w:sz')
                rPr.append(sz)
            sz.set(qn('w:val'), half_pts)

            sz_cs = rPr.find(qn('w:szCs'))
            if sz_cs is None:
                sz_cs = OxmlElement('w:szCs')
                rPr.append(sz_cs)
            sz_cs.set(qn('w:val'), half_pts)

            # Đồng bộ linked character style nếu có (ví dụ: Normal Char)
            char_style_names = [f"{style.name} Char", f"{style.name}Char"]
            for s in self.styles:
                if s.type == WD_STYLE_TYPE.CHARACTER and s.name in char_style_names:
                    try:
                        s_font = s.font
                        s_font.name = Config.FONT_NAME
                        s_font.size = Pt(Config.BODY_FONT_SIZE)
                        s_font.color.rgb = RGBColor(0, 0, 0)
                        c_rPr = s._element.get_or_add_rPr()
                        self._set_rfonts(c_rPr, Config.FONT_NAME)
                        c_sz = c_rPr.find(qn('w:sz'))
                        if c_sz is None:
                            c_sz = OxmlElement('w:sz')
                            c_rPr.append(c_sz)
                        c_sz.set(qn('w:val'), half_pts)
                        c_sz_cs = c_rPr.find(qn('w:szCs'))
                        if c_sz_cs is None:
                            c_sz_cs = OxmlElement('w:szCs')
                            c_rPr.append(c_sz_cs)
                        c_sz_cs.set(qn('w:val'), half_pts)
                    except Exception as e:
                        logger.debug(f"Lỗi đồng bộ linked character style {s.name}: {e}")

            # Đồng bộ thêm style Body Text / Body Text Indent nếu tồn tại trong tài liệu
            for bt_name in ['Body Text', 'Body Text Indent']:
                try:
                    bt_style = self.styles[bt_name]
                    bt_style.font.name = Config.FONT_NAME
                    bt_style.font.size = Pt(Config.BODY_FONT_SIZE)
                    bt_style.font.color.rgb = RGBColor(0, 0, 0)
                    bt_rPr = bt_style._element.get_or_add_rPr()
                    self._set_rfonts(bt_rPr, Config.FONT_NAME)
                    bt_sz = bt_rPr.find(qn('w:sz'))
                    if bt_sz is None:
                        bt_sz = OxmlElement('w:sz')
                        bt_rPr.append(bt_sz)
                    bt_sz.set(qn('w:val'), half_pts)
                    bt_sz_cs = bt_rPr.find(qn('w:szCs'))
                    if bt_sz_cs is None:
                        bt_sz_cs = OxmlElement('w:szCs')
                        bt_rPr.append(bt_sz_cs)
                    bt_sz_cs.set(qn('w:val'), half_pts)
                except KeyError:
                    pass
                except Exception as e:
                    logger.debug(f"Lỗi khi cấu hình {bt_name} style: {e}")

            # Paragraph format mặc định
            pf = style.paragraph_format
            pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            pf.line_spacing = Config.LINE_SPACING
            pf.space_before = Pt(Config.SPACE_BEFORE_PT)
            pf.space_after = Pt(Config.SPACE_AFTER_PT)
        except Exception as e:
            logger.debug(f"Lỗi khi cấu hình Normal style: {e}")

    def _get_or_add_paragraph_style(self, name: str, style_id: str = None, builtin: bool = False):
        """
        Lấy style theo tên hoặc style_id. Nếu chưa có, mới tạo mới một lần duy nhất,
        tránh sinh ra các style trùng lặp như 'Heading 11'.
        """
        try:
            return self.styles[name]
        except KeyError:
            pass
        for s in self.styles:
            if s.name == name or (style_id and s.style_id == style_id):
                return s
        return self.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH, builtin=builtin)

    def _setup_heading_1(self):
        """
        Heading 1: 16pt Bold, Căn giữa (Center), 12pt before, 6pt after, left_indent = 0, outlineLevel 0
        Chuẩn hóa 100% theo Form mẫu (Heading 1 căn giữa, 16pt)
        """
        try:
            style = self._get_or_add_paragraph_style('Heading 1', 'Heading1', builtin=True)
            try:
                style.base_style = self.styles['Normal']
            except Exception:
                pass

            self._apply_heading_font_properties(style, Config.HEADING_1_FONT_SIZE, bold=True, italic=False)

            pf = style.paragraph_format
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.left_indent = Cm(0)
            pf.first_line_indent = Cm(0)
            pf.space_before = Pt(Config.HEADING_1_SPACE_BEFORE_PT)
            pf.space_after = Pt(Config.HEADING_1_SPACE_AFTER_PT)
            pf.line_spacing = Config.LINE_SPACING
            pf.keep_with_next = True

            # Đảm bảo outlineLvl = 0 và jc = center trong XML
            pPr = style._element.get_or_add_pPr()
            outline = pPr.find(qn('w:outlineLvl'))
            if outline is None:
                outline = OxmlElement('w:outlineLvl')
                pPr.append(outline)
            outline.set(qn('w:val'), '0')

            jc = pPr.find(qn('w:jc'))
            if jc is None:
                jc = OxmlElement('w:jc')
                pPr.append(jc)
            jc.set(qn('w:val'), 'center')
        except Exception as e:
            logger.debug(f"Lỗi khi cấu hình Heading 1: {e}")

    def _setup_heading_2(self):
        """
        Heading 2: 14pt Bold, Căn trái (Left), 6pt before, 3pt after, left_indent = 0cm, outlineLevel 1
        """
        try:
            style = self._get_or_add_paragraph_style('Heading 2', 'Heading2', builtin=True)
            try:
                style.base_style = self.styles['Normal']
            except Exception:
                pass

            self._apply_heading_font_properties(style, Config.HEADING_2_FONT_SIZE, bold=True, italic=False)

            pf = style.paragraph_format
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf.left_indent = Cm(Config.HEADING_INDENT_PER_LEVEL_CM * 1)
            pf.first_line_indent = Cm(0)
            pf.space_before = Pt(Config.HEADING_SPACE_BEFORE_PT)
            pf.space_after = Pt(Config.HEADING_SPACE_AFTER_PT)
            pf.line_spacing = Config.LINE_SPACING
            pf.keep_with_next = True

            pPr = style._element.get_or_add_pPr()
            outline = pPr.find(qn('w:outlineLvl'))
            if outline is None:
                outline = OxmlElement('w:outlineLvl')
                pPr.append(outline)
            outline.set(qn('w:val'), '1')

            jc = pPr.find(qn('w:jc'))
            if jc is None:
                jc = OxmlElement('w:jc')
                pPr.append(jc)
            jc.set(qn('w:val'), 'left')
        except Exception as e:
            logger.debug(f"Lỗi khi cấu hình Heading 2: {e}")

    def _setup_heading_3(self):
        """
        Heading 3: 14pt Bold + Italic, Căn trái (Left), 6pt before, 3pt after, left_indent = 0cm, outlineLevel 2
        Chuẩn hóa 100% theo Form mẫu (Heading 3 in đậm và in nghiêng)
        """
        try:
            style = self._get_or_add_paragraph_style('Heading 3', 'Heading3', builtin=True)
            try:
                style.base_style = self.styles['Normal']
            except Exception:
                pass

            self._apply_heading_font_properties(style, Config.HEADING_3_FONT_SIZE, bold=True, italic=True)

            pf = style.paragraph_format
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf.left_indent = Cm(Config.HEADING_INDENT_PER_LEVEL_CM * 2)
            pf.first_line_indent = Cm(0)
            pf.space_before = Pt(Config.HEADING_SPACE_BEFORE_PT)
            pf.space_after = Pt(Config.HEADING_SPACE_AFTER_PT)
            pf.line_spacing = Config.LINE_SPACING
            pf.keep_with_next = True

            pPr = style._element.get_or_add_pPr()
            outline = pPr.find(qn('w:outlineLvl'))
            if outline is None:
                outline = OxmlElement('w:outlineLvl')
                pPr.append(outline)
            outline.set(qn('w:val'), '2')

            jc = pPr.find(qn('w:jc'))
            if jc is None:
                jc = OxmlElement('w:jc')
                pPr.append(jc)
            jc.set(qn('w:val'), 'left')
        except Exception as e:
            logger.debug(f"Lỗi khi cấu hình Heading 3: {e}")

    def _setup_heading_4(self):
        """
        Heading 4: 14pt Bold + Italic, Căn trái (Left), 6pt before, 3pt after, left_indent = 0cm, outlineLevel 3
        """
        try:
            style = self._get_or_add_paragraph_style('Heading 4', 'Heading4', builtin=True)
            try:
                style.base_style = self.styles['Normal']
            except Exception:
                pass

            self._apply_heading_font_properties(style, Config.HEADING_4_FONT_SIZE, bold=True, italic=True)

            pf = style.paragraph_format
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf.left_indent = Cm(Config.HEADING_INDENT_PER_LEVEL_CM * 3)
            pf.first_line_indent = Cm(0)
            pf.space_before = Pt(Config.HEADING_SPACE_BEFORE_PT)
            pf.space_after = Pt(Config.HEADING_SPACE_AFTER_PT)
            pf.line_spacing = Config.LINE_SPACING
            pf.keep_with_next = True

            pPr = style._element.get_or_add_pPr()
            outline = pPr.find(qn('w:outlineLvl'))
            if outline is None:
                outline = OxmlElement('w:outlineLvl')
                pPr.append(outline)
            outline.set(qn('w:val'), '3')

            jc = pPr.find(qn('w:jc'))
            if jc is None:
                jc = OxmlElement('w:jc')
                pPr.append(jc)
            jc.set(qn('w:val'), 'left')
        except Exception as e:
            logger.debug(f"Lỗi khi cấu hình Heading 4: {e}")

    def _setup_heading_1_unnumbered(self):
        """
        Lưu ý kỹ thuật 2: Style Heading 1 Unnumbered
        Dành cho: LỜI MỞ ĐẦU, KẾT LUẬN, TÀI LIỆU THAM KHẢO
        Định dạng: 14pt Bold, căn giữa/trái, outlineLvl = 0 ĐỂ HIỂN THỊ TRONG MỤC LỤC
        """
        style_name = Config.STYLE_HEADING_1_UNNUMBERED
        try:
            try:
                style = self.styles[style_name]
            except KeyError:
                style = self.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
                try:
                    style.base_style = self.styles['Normal']
                except Exception:
                    pass

            self._apply_heading_font_properties(style, Config.HEADING_1_FONT_SIZE, bold=True, italic=False)

            pf = style.paragraph_format
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.left_indent = Cm(0)
            pf.first_line_indent = Cm(0)
            pf.space_before = Pt(Config.HEADING_1_SPACE_BEFORE_PT)
            pf.space_after = Pt(Config.HEADING_1_SPACE_AFTER_PT)
            pf.line_spacing = Config.LINE_SPACING
            pf.keep_with_next = True

            # Gắn outlineLvl = 0 để Word TOC tự động thu thập
            pPr = style._element.get_or_add_pPr()
            outline = pPr.find(qn('w:outlineLvl'))
            if outline is None:
                outline = OxmlElement('w:outlineLvl')
                pPr.append(outline)
            outline.set(qn('w:val'), '0')

        except Exception as e:
            logger.debug(f"Lỗi khi cấu hình {style_name}: {e}")

    def _setup_non_toc_heading(self):
        """
        Lưu ý kỹ thuật 2: Style Non-TOC Heading
        Dành cho: LỜI CẢM ƠN, MỤC LỤC, DANH MỤC HÌNH ẢNH, DANH MỤC BẢNG BIỂU
        Định dạng: Giống Heading 1 nhưng KHÔNG có outlineLvl (không xuất hiện trong TOC)
        """
        style_name = Config.STYLE_NON_TOC_HEADING
        try:
            try:
                style = self.styles[style_name]
            except KeyError:
                style = self.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
                try:
                    style.base_style = self.styles['Normal']
                except Exception:
                    pass

            self._apply_heading_font_properties(style, Config.HEADING_1_FONT_SIZE, bold=True, italic=False)

            pf = style.paragraph_format
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.left_indent = Cm(0)
            pf.first_line_indent = Cm(0)
            pf.space_before = Pt(Config.HEADING_1_SPACE_BEFORE_PT)
            pf.space_after = Pt(Config.HEADING_1_SPACE_AFTER_PT)
            pf.line_spacing = Config.LINE_SPACING
            pf.keep_with_next = True

            # Đảm bảo KHÔNG CÓ outlineLvl
            pPr = style._element.get_or_add_pPr()
            outline = pPr.find(qn('w:outlineLvl'))
            if outline is not None:
                pPr.remove(outline)

        except Exception as e:
            logger.debug(f"Lỗi khi cấu hình {style_name}: {e}")

    def _setup_caption_style(self):
        """
        Style Caption: Times New Roman, 13pt/14pt, Italic, Center, 6pt before, 6pt after
        Đồng bộ cấu hình cho cả Style Caption mặc định và custom styles 'hinhanh', 'bang' theo Form mẫu
        """
        try:
            try:
                style = self.styles['Caption']
            except KeyError:
                style = self.styles.add_style('Caption', WD_STYLE_TYPE.PARAGRAPH, builtin=True)
                try:
                    style.base_style = self.styles['Normal']
                except Exception:
                    pass

            self._apply_heading_font_properties(style, Config.CAPTION_FONT_SIZE, bold=False, italic=True)

            pf = style.paragraph_format
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.first_line_indent = Cm(0)
            pf.left_indent = Cm(0)
            pf.space_before = Pt(6)
            pf.space_after = Pt(6)
            pf.line_spacing = 1.15
            pf.keep_with_next = True

            pPr = style._element.get_or_add_pPr()
            jc = pPr.find(qn('w:jc'))
            if jc is None:
                jc = OxmlElement('w:jc')
                pPr.append(jc)
            jc.set(qn('w:val'), 'center')
        except Exception as e:
            logger.debug(f"Lỗi khi cấu hình Caption style: {e}")

        # Đồng bộ các Style Caption riêng 'hinhanh' và 'bang' nếu có hoặc tạo mới
        for cap_name in ['hinhanh', 'bang']:
            try:
                try:
                    c_style = self.styles[cap_name]
                except KeyError:
                    c_style = self.styles.add_style(cap_name, WD_STYLE_TYPE.PARAGRAPH)
                    try:
                        c_style.base_style = self.styles['Normal']
                    except Exception:
                        pass
                self._apply_heading_font_properties(c_style, Config.CAPTION_FONT_SIZE, bold=False, italic=True)
                cpf = c_style.paragraph_format
                cpf.alignment = WD_ALIGN_PARAGRAPH.CENTER
                cpf.first_line_indent = Cm(0)
                cpf.left_indent = Cm(0)
                cpf.space_before = Pt(6)
                cpf.space_after = Pt(6)
                cpf.line_spacing = 1.15
                cpf.keep_with_next = True

                c_pPr = c_style._element.get_or_add_pPr()
                c_jc = c_pPr.find(qn('w:jc'))
                if c_jc is None:
                    c_jc = OxmlElement('w:jc')
                    c_pPr.append(c_jc)
                c_jc.set(qn('w:val'), 'center')
            except Exception as e:
                logger.debug(f"Lỗi khi cấu hình style {cap_name}: {e}")
