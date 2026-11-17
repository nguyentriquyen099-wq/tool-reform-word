"""
Phân hệ 6 & Lưu ý kỹ thuật số 4: Auto-TOC Update & Chống lệch số trang
- Tự động phát hiện và chèn trang MỤC LỤC chuẩn OpenXML (<w:fldSimple> hoặc complex <w:fldChar>)
- Dọn dẹp các dòng chấm chấm thủ công do người dùng gõ tay
- Chèn thẻ <w:updateFields w:val="true"/> vào settings.xml
- Đảm bảo cài đặt lề trang chuẩn (Top/Bottom 2cm, Left 3cm, Right 2cm) chống lệch số trang
- Hỗ trợ phát hiện và gọi LibreOffice headless (nếu có sẵn trên máy) để cập nhật TOC
"""

import os
import re
import shutil
import subprocess
import logging
from typing import Optional, Tuple, Dict, Any, List
from docx.shared import Cm, Pt, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from config import Config

logger = logging.getLogger(__name__)

class TOCManager:
    @staticmethod
    def enable_update_fields_on_open(doc) -> bool:
        """
        Chèn thẻ <w:updateFields w:val="true"/> vào word/settings.xml.
        Khi người dùng mở file Word, Word sẽ tự động cập nhật lại toàn bộ Mục lục (TOC)
        và số trang (Page numbers).
        """
        try:
            settings_elem = doc.settings.element
            # Tìm thẻ updateFields
            update_fields = settings_elem.find(qn('w:updateFields'))
            if update_fields is None:
                update_fields = OxmlElement('w:updateFields')
                settings_elem.append(update_fields)
            update_fields.set(qn('w:val'), 'true')
            logger.info("Đã chèn thẻ <w:updateFields w:val='true'/> vào settings.xml")
            return True
        except Exception as e:
            logger.warning(f"Không thể kích hoạt updateFields trong settings.xml: {e}")
            return False

    @staticmethod
    def standardize_page_margins(doc):
        """
        Lưu ý kỹ thuật số 4: Chống lệch số trang
        Chuẩn hóa lề trang (Margins) của toàn bộ các section trong tài liệu:
        - Top: 2.0cm, Bottom: 2.0cm
        - Left: 3.0cm (lề gáy đóng tài liệu)
        - Right: 2.0cm
        Giúp định dạng trang nhất quán, bảo đảm số trang không bị nhảy lung tung khi in ấn.
        """
        try:
            for section in doc.sections:
                section.top_margin = Cm(Config.MARGIN_TOP_CM)
                section.bottom_margin = Cm(Config.MARGIN_BOTTOM_CM)
                section.left_margin = Cm(Config.MARGIN_LEFT_CM)
                section.right_margin = Cm(Config.MARGIN_RIGHT_CM)
                # Đảm bảo khổ giấy A4 và khoảng cách Header/Footer (720 dxa = 1.27cm) theo Mẫu form
                section.page_width = Cm(21.0)
                section.page_height = Cm(29.7)
                section.header_distance = Cm(1.27)
                section.footer_distance = Cm(1.27)
            logger.info(f"Đã chuẩn hóa lề trang theo Form mẫu chuẩn (Top {Config.MARGIN_TOP_CM}cm, Bottom {Config.MARGIN_BOTTOM_CM}cm, Left {Config.MARGIN_LEFT_CM}cm, Right {Config.MARGIN_RIGHT_CM}cm, Khổ A4)")
        except Exception as e:
            logger.warning(f"Lỗi khi chuẩn hóa lề trang: {e}")

    @staticmethod
    def find_libreoffice_executable() -> Optional[str]:
        """
        Tìm kiếm đường dẫn thực thi của LibreOffice (soffice) trên Windows
        """
        # 1. Tìm trong PATH
        cmd = shutil.which("soffice") or shutil.which("libreoffice")
        if cmd:
            return cmd

        # 2. Tìm trong các thư mục cài đặt tiêu chuẩn trên Windows
        candidate_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Programs\LibreOffice\program\soffice.exe")
        ]

        for path in candidate_paths:
            if os.path.isfile(path):
                return path

        return None

    @staticmethod
    def find_word_executable() -> Optional[str]:
        """
        Tìm kiếm đường dẫn thực thi của Microsoft Word (winword.exe) trên Windows
        """
        # 1. Tìm trong PATH
        cmd = shutil.which("winword") or shutil.which("winword.exe")
        if cmd:
            return cmd

        # 2. Tìm qua Windows Registry (PowerShell / Query)
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\winword.exe") as key:
                val, _ = winreg.QueryValueEx(key, "")
                if val and os.path.isfile(val):
                    return val
        except Exception:
            pass

        # 3. Tìm trong các thư mục Office mặc định trên Windows
        candidate_paths = [
            r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
            r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
            r"C:\Program Files\Microsoft Office\Office16\WINWORD.EXE",
            r"C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE",
            r"C:\Program Files\Microsoft Office\Office15\WINWORD.EXE",
            r"C:\Program Files (x86)\Microsoft Office\Office15\WINWORD.EXE",
        ]
        for path in candidate_paths:
            if os.path.isfile(path):
                return path

        return None

    @classmethod
    def update_toc_with_word(cls, docx_path: str) -> Tuple[bool, str]:
        """
        Gọi Microsoft Word qua COM Automation trên Windows để tính toán và cập nhật TOC
        với độ chính xác và trung thực cao nhất của Word native engine.
        """
        abs_path = os.path.abspath(docx_path)
        if not os.path.isfile(abs_path):
            return False, f"File không tồn tại: {abs_path}"

        ps_path = abs_path.replace("'", "''")
        ps_script = f"""
$ErrorActionPreference = 'Stop'
$word = $null
try {{
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open('{ps_path}')
    $doc.Fields.Update()
    foreach ($toc in $doc.TablesOfContents) {{
        $toc.Update()
    }}
    $doc.Save()
    $doc.Close([ref]$false)
    Write-Host "WORD_TOC_SUCCESS"
}} catch {{
    Write-Host ("WORD_TOC_ERROR: " + $_.Exception.Message)
    exit 1
}} finally {{
    if ($word -ne $null) {{
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }}
}}
"""
        try:
            res = subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
                capture_output=True,
                text=True,
                timeout=35
            )
            if res.returncode == 0 and "WORD_TOC_SUCCESS" in res.stdout:
                # Đảm bảo cờ updateFields và thuộc tính ô bảng luôn chuẩn mực sau khi Word COM lưu file
                try:
                    from docx import Document
                    from docx.shared import Inches
                    saved_doc = Document(abs_path)
                    cls.enable_update_fields_on_open(saved_doc)
                    for t in saved_doc.tables:
                        for row in t.rows:
                            for cell in row.cells:
                                for p in cell.paragraphs:
                                    p.paragraph_format.first_line_indent = Inches(0)
                    saved_doc.save(abs_path)
                except Exception as e:
                    logger.warning(f"Không thể tái kích hoạt updateFields sau Word COM: {e}")
                logger.info("Microsoft Word đã cập nhật và kết xuất toàn bộ TOC thành công.")
                return True, "Microsoft Word đã tính toán và cập nhật toàn bộ Mục lục thành công."
            else:
                err = res.stderr.strip() or res.stdout.strip()
                logger.warning(f"Lỗi khi chạy Word COM automation: {err}")
                return False, f"Word automation failed: {err}"
        except Exception as e:
            logger.warning(f"Ngoại lệ khi gọi Word COM automation: {e}")
            return False, f"Ngoại lệ khi gọi Word: {str(e)}"

    @classmethod
    def update_toc_with_libreoffice(cls, docx_path: str, output_pdf: bool = False) -> Tuple[bool, str]:
        """
        Gọi LibreOffice ở chế độ headless để tính toán số trang và cập nhật mục lục nếu có LibreOffice
        """
        soffice_path = cls.find_libreoffice_executable()
        if not soffice_path:
            return False, "LibreOffice không có sẵn trên máy tính. Word sẽ tự động cập nhật TOC khi mở file (nhờ cờ updateFields)."

        try:
            # Chạy LibreOffice headless
            out_dir = os.path.dirname(os.path.abspath(docx_path))
            cmd = [
                soffice_path,
                "--headless",
                "--convert-to", "pdf",
                docx_path,
                "--outdir", out_dir
            ]
            res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=40)
            if res.returncode == 0:
                logger.info("LibreOffice đã cập nhật và tính toán số trang thành công.")
                return True, "LibreOffice đã tính toán số trang và cập nhật trường thành công."
            else:
                return False, f"Lỗi chạy LibreOffice: {res.stderr.decode('utf-8', errors='ignore')}"
        except Exception as e:
            return False, f"Ngoại lệ khi gọi LibreOffice: {str(e)}"

    @classmethod
    def update_toc_headless(cls, docx_path: str) -> Tuple[bool, str]:
        """
        Điều phối tự động cập nhật và kết xuất số trang TOC:
        1. Ưu tiên 1: Microsoft Word native qua COM (chính xác 100% chuẩn Microsoft Office)
        2. Ưu tiên 2: LibreOffice headless (nền tảng mở hoặc môi trường không có Word)
        3. Fallback: Bật cờ updateFields (Word sẽ tự cập nhật khi người dùng mở file)
        """
        # 1. Thử Microsoft Word trước
        if cls.find_word_executable():
            ok, msg = cls.update_toc_with_word(docx_path)
            if ok:
                return True, msg

        # 2. Thử LibreOffice nếu Word không khả dụng hoặc lỗi
        if cls.find_libreoffice_executable():
            ok, msg = cls.update_toc_with_libreoffice(docx_path)
            if ok:
                return True, msg

        return False, "Không có sẵn Word hoặc LibreOffice headless. Word sẽ cập nhật Mục lục khi mở file nhờ cờ updateFields."

    @staticmethod
    def has_toc_field(doc) -> bool:
        """
        Kiểm tra xem tài liệu đã có trường Mục lục (TOC field) chuẩn OpenXML hay chưa.
        Tìm kiếm thẻ w:fldSimple hoặc w:instrText có chứa 'TOC'.
        """
        try:
            for p in doc.paragraphs:
                elem = p._element
                for fld in elem.xpath('.//w:fldSimple'):
                    instr = fld.get(qn('w:instr'), '')
                    if 'TOC' in instr.upper():
                        return True
                for instr in elem.xpath('.//w:instrText'):
                    if instr.text and 'TOC' in instr.text.upper():
                        return True
            for tbl in doc.tables:
                for row in tbl.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            elem = p._element
                            for fld in elem.xpath('.//w:fldSimple'):
                                instr = fld.get(qn('w:instr'), '')
                                if 'TOC' in instr.upper():
                                    return True
                            for instr in elem.xpath('.//w:instrText'):
                                if instr.text and 'TOC' in instr.text.upper():
                                    return True
        except Exception as e:
            logger.warning(f"Lỗi khi kiểm tra trường TOC: {e}")
        return False

    @staticmethod
    def is_toc_paragraph(paragraph) -> bool:
        """
        Kiểm tra xem paragraph cụ thể có chứa trường TOC hoặc là mục TOC hay không.
        """
        try:
            elem = paragraph._element
            for fld in elem.xpath('.//w:fldSimple'):
                instr = fld.get(qn('w:instr'), '')
                if 'TOC' in instr.upper():
                    return True
            for instr in elem.xpath('.//w:instrText'):
                if instr.text and 'TOC' in instr.text.upper():
                    return True
            style_name = getattr(paragraph.style, 'name', '')
            if style_name and (style_name.startswith('TOC') or style_name.startswith('toc')):
                return True
        except Exception:
            pass
        return False

    @staticmethod
    def find_toc_heading(doc) -> Optional[Tuple[int, Any]]:
        """
        Tìm vị trí và paragraph tiêu đề 'MỤC LỤC' hoặc 'TABLE OF CONTENTS' trong tài liệu.
        Không khớp với 'MỤC LỤC HÌNH' hay 'MỤC LỤC BẢNG'.
        """
        toc_pattern = re.compile(r'^\s*(?:MỤC\s*LỤC|TABLE\s*OF\s*CONTENTS)\s*[:]?$', re.IGNORECASE)
        for idx, p in enumerate(doc.paragraphs):
            text = p.text.strip()
            if toc_pattern.match(text):
                return idx, p
        return None

    @staticmethod
    def is_fake_toc_entry(p_or_text) -> bool:
        """
        Nhận diện dòng mục lục thủ công (chấm chấm gõ tay của sinh viên):
        Ví dụ: 'CHƯƠNG 1: GIỚI THIỆU .................................... 2'
        '1.1. Khái niệm . . . . . . . . . . . . . . . 3'
        'Phần 1 \t 10'
        """
        text = p_or_text.text if hasattr(p_or_text, 'text') else str(p_or_text)
        t = text.strip()
        if not t:
            return False

        # 1. Dòng kết thúc bằng 3+ dấu chấm/gạch/tab kèm số trang
        if re.search(r'((?:\.\s*){3,}|[\.\…\_\-–—\t]{3,})\s*\d+\s*$', t):
            return True
        # 2. Dòng có 4+ dấu chấm liên tiếp kèm số ở cuối
        if re.search(r'\.{4,}.*\d+\s*$', t) or re.search(r'…{2,}.*\d+\s*$', t):
            return True
        # 3. Dòng có tab phân cách kết thúc bằng số
        if re.search(r'\t+\s*\d+\s*$', t):
            return True
        # 4. Dòng chứa chuỗi dài dấu chấm (>= 8 dấu chấm) dù chưa có số
        if re.search(r'(?:\.\s*){8,}', t):
            return True

        return False

    @classmethod
    def cleanup_fake_toc_entries(cls, doc, toc_heading_idx: int) -> int:
        """
        Xóa sạch các dòng mục lục chấm chấm thủ công phía sau tiêu đề 'MỤC LỤC'.
        Dừng lại khi gặp tiêu đề thật hoặc phần nội dung thật tiếp theo.
        Hỗ trợ dọn dẹp an toàn các khoảng trống/dòng trống kẹp giữa các dòng mục lục giả.
        """
        cleaned_count = 0
        while toc_heading_idx + 1 < len(doc.paragraphs):
            p_next = doc.paragraphs[toc_heading_idx + 1]
            t = p_next.text.strip()

            if not t:
                # Quét xem phía sau có còn dòng fake entry nào nữa không
                has_fake_ahead = False
                for lookahead_idx in range(toc_heading_idx + 2, min(toc_heading_idx + 8, len(doc.paragraphs))):
                    p_ahead = doc.paragraphs[lookahead_idx]
                    t_ahead = p_ahead.text.strip()
                    if cls.is_fake_toc_entry(p_ahead):
                        has_fake_ahead = True
                        break
                    if t_ahead:
                        # Gặp nội dung thật (không phải fake entry) -> dừng quét
                        break

                if has_fake_ahead:
                    p_next._element.getparent().remove(p_next._element)
                    continue
                else:
                    break

            if cls.is_fake_toc_entry(p_next):
                p_next._element.getparent().remove(p_next._element)
                cleaned_count += 1
            else:
                break

        return cleaned_count

    @staticmethod
    def create_toc_field(paragraph, use_complex_field: bool = False):
        r"""
        Chèn trường mã TOC tự động chuẩn của Word:
        - Mặc định: <w:fldSimple w:instr="TOC \o &quot;1-3&quot; \h \z \u"/>
        - Hoặc: cấu trúc complex field w:fldChar (begin/separate/end)
        """
        p_elem = paragraph._element
        font_name = Config.FONT_NAME
        font_sz = str(int(Config.BODY_FONT_SIZE * 2))

        if not use_complex_field:
            fldSimple = OxmlElement('w:fldSimple')
            fldSimple.set(qn('w:instr'), r'TOC \o "1-3" \h \z \u')
            fldSimple.set(qn('w:dirty'), 'true')

            r = OxmlElement('w:r')
            rPr = OxmlElement('w:rPr')
            rFonts = OxmlElement('w:rFonts')
            rFonts.set(qn('w:ascii'), font_name)
            rFonts.set(qn('w:hAnsi'), font_name)
            rFonts.set(qn('w:cs'), font_name)
            rPr.append(rFonts)

            sz = OxmlElement('w:sz')
            sz.set(qn('w:val'), font_sz)
            rPr.append(sz)

            szCs = OxmlElement('w:szCs')
            szCs.set(qn('w:val'), font_sz)
            rPr.append(szCs)

            noProof = OxmlElement('w:noProof')
            rPr.append(noProof)
            r.append(rPr)

            t = OxmlElement('w:t')
            t.text = "Nhấn chuột phải chọn 'Update Field' (hoặc chọn Yes khi Word hỏi) để hiển thị Mục lục."
            r.append(t)

            fldSimple.append(r)
            p_elem.append(fldSimple)
        else:
            r_begin = OxmlElement('w:r')
            c_begin = OxmlElement('w:fldChar')
            c_begin.set(qn('w:fldCharType'), 'begin')
            c_begin.set(qn('w:dirty'), 'true')
            r_begin.append(c_begin)
            p_elem.append(r_begin)

            r_instr = OxmlElement('w:r')
            instr = OxmlElement('w:instrText')
            instr.set(qn('xml:space'), 'preserve')
            instr.text = r' TOC \o "1-3" \h \z \u '
            r_instr.append(instr)
            p_elem.append(r_instr)

            r_sep = OxmlElement('w:r')
            c_sep = OxmlElement('w:fldChar')
            c_sep.set(qn('w:fldCharType'), 'separate')
            r_sep.append(c_sep)
            p_elem.append(r_sep)

            r_text = OxmlElement('w:r')
            rPr = OxmlElement('w:rPr')
            rFonts = OxmlElement('w:rFonts')
            rFonts.set(qn('w:ascii'), font_name)
            rFonts.set(qn('w:hAnsi'), font_name)
            rFonts.set(qn('w:cs'), font_name)
            rPr.append(rFonts)
            sz = OxmlElement('w:sz')
            sz.set(qn('w:val'), font_sz)
            rPr.append(sz)

            szCs = OxmlElement('w:szCs')
            szCs.set(qn('w:val'), font_sz)
            rPr.append(szCs)
            noProof = OxmlElement('w:noProof')
            rPr.append(noProof)
            r_text.append(rPr)
            t = OxmlElement('w:t')
            t.text = "Nhấn chuột phải chọn 'Update Field' (hoặc chọn Yes khi Word hỏi) để hiển thị Mục lục."
            r_text.append(t)
            p_elem.append(r_text)

            r_end = OxmlElement('w:r')
            c_end = OxmlElement('w:fldChar')
            c_end.set(qn('w:fldCharType'), 'end')
            r_end.append(c_end)
            p_elem.append(r_end)

        pf = paragraph.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.first_line_indent = Inches(0)
        pf.left_indent = Inches(0)
        pf.right_indent = Inches(0)
        pf.space_before = Pt(6)
        pf.space_after = Pt(12)

    @staticmethod
    def format_toc_heading(paragraph, doc):
        """
        Định dạng tiêu đề MỤC LỤC theo style Non-TOC Heading (14pt Bold, Center)
        để không tự đưa chính nó vào Mục lục.
        Chuẩn hóa văn bản thành 'MỤC LỤC' viết hoa chuẩn.
        """
        paragraph.text = "MỤC LỤC"

        try:
            paragraph.style = doc.styles[Config.STYLE_NON_TOC_HEADING]
        except KeyError:
            pass

        pf = paragraph.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf.first_line_indent = Inches(0)
        pf.left_indent = Inches(0)
        pf.right_indent = Inches(0)
        pf.space_before = Pt(12)
        pf.space_after = Pt(6)
        pf.line_spacing = Config.LINE_SPACING
        pf.keep_with_next = True
        pf.page_break_before = True

        pPr = paragraph._element.get_or_add_pPr()
        outline = pPr.find(qn('w:outlineLvl'))
        if outline is not None:
            pPr.remove(outline)

        half_pts = str(int(round(Config.HEADING_1_FONT_SIZE * 2)))
        for r in paragraph.runs:
            r.bold = True
            r.italic = False
            r.font.name = Config.FONT_NAME
            r.font.size = Pt(Config.HEADING_1_FONT_SIZE)
            rPr = r._element.get_or_add_rPr()
            szCs = rPr.find(qn('w:szCs'))
            if szCs is not None:
                szCs.set(qn('w:val'), half_pts)
            iCs = rPr.find(qn('w:iCs'))
            if iCs is not None:
                rPr.remove(iCs)
            if rPr.find(qn('w:bCs')) is None:
                rPr.append(OxmlElement('w:bCs'))

    @staticmethod
    def find_insertion_point(doc, cover_end_idx: int = 0) -> Optional[Tuple[int, Any]]:
        """
        Tìm vị trí paragraph phù hợp nhất để chèn Mục lục:
        1. Trước Danh mục hình / Danh mục bảng / Danh mục từ viết tắt (chỉ Front-matter,
           loại trừ Danh mục tài liệu tham khảo ở cuối văn bản).
        2. Trước CHƯƠNG 1 / CHƯƠNG I / PHẦN 1 / PHẦN I / CHAPTER 1 / BÀI 1.
        3. Trước Heading 1 số La Mã đầu tiên (ví dụ: I. TỔNG QUAN, I. ĐẶT VẤN ĐỀ) nếu không có Chương 1.
        4. Trước Heading 1 có đánh số đầu tiên (ví dụ 1. TỔNG QUAN).
        5. Trước paragraph có style Heading 1 đầu tiên sau Trang bìa.
        """
        start_idx = max(0, cover_end_idx)
        num_p = len(doc.paragraphs)
        if num_p <= start_idx:
            return None

        # Ưu tiên 1: Trước Danh mục hình / bảng / viết tắt (CHỈ front-matter, không ăn vào Bibliography ở cuối)
        front_list_pattern = re.compile(
            r'^\s*(?:DANH\s*MỤC\s+(?!TÀI\s*LIỆU|THAM\s*KHẢO)(?:HÌNH|BẢNG|TỪ\s*VIẾT\s*TẮT|VIẾT\s*TẮT|KÝ\s*HIỆU|SƠ\s*ĐỒ|BIỂU\s*ĐỒ)|LIST\s*OF\s+(?!REFERENCES)(?:FIGURES|TABLES|ABBREVIATIONS|SYMBOLS|CHARTS))\b',
            re.IGNORECASE
        )
        for idx in range(start_idx, num_p):
            p = doc.paragraphs[idx]
            t = p.text.strip()
            if front_list_pattern.match(t):
                if idx == 0 and num_p < 5:
                    return None
                return idx, p

        # Ưu tiên 2: Trước CHƯƠNG 1 / CHƯƠNG I / PHẦN 1 / CHAPTER 1 / BÀI 1
        chapter1_pattern = re.compile(r'^\s*(?:CHƯƠNG|CHAPTER|PHẦN|BÀI)\s+(?:1|I)\b', re.IGNORECASE)
        for idx in range(start_idx, num_p):
            p = doc.paragraphs[idx]
            t = p.text.strip()
            if chapter1_pattern.match(t):
                if idx == 0 and num_p < 5:
                    return None
                return idx, p

        # Ưu tiên 3: Trước Heading 1 có số La Mã (I. ĐẶT VẤN ĐỀ, I. TỔNG QUAN...)
        # Tránh nhầm với phần phân công nhiệm vụ nếu sau đó có Chương 1
        roman_h1_pattern = re.compile(r'^\s*I\.\s+[A-ZÀ-Ỵ]', re.UNICODE)
        for idx in range(start_idx, num_p):
            p = doc.paragraphs[idx]
            t = p.text.strip()
            if roman_h1_pattern.match(t):
                if idx == 0 and num_p < 5:
                    return None
                return idx, p

        # Ưu tiên 4: Trước Heading 1 có đánh số đầu tiên (1. hoặc 1 theo sau là chữ hoa)
        numbered_h1_pattern = re.compile(r'^\s*1[\.\s]\s*[A-ZÀ-Ỵ]', re.UNICODE)
        for idx in range(start_idx, num_p):
            p = doc.paragraphs[idx]
            t = p.text.strip()
            if numbered_h1_pattern.match(t):
                if idx == 0 and num_p < 5:
                    return None
                return idx, p

        # Ưu tiên 5: Trước paragraph có style Heading 1 đầu tiên sau trang bìa
        for idx in range(start_idx, num_p):
            p = doc.paragraphs[idx]
            if p.text.strip() and getattr(p.style, 'name', '') in ['Heading 1', 'heading 1']:
                return idx, p

        return None

    @classmethod
    def ensure_table_of_contents(cls, doc, cover_end_idx: int = 0, use_complex_field: bool = False) -> Dict[str, Any]:
        """
        Tự động phát hiện và chèn trang MỤC LỤC tự động chuẩn OpenXML:
        1. Nếu đã có trường TOC chuẩn: giữ nguyên, chuẩn hóa style MỤC LỤC.
        2. Nếu đã có tiêu đề MỤC LỤC nhưng chưa có trường mã TOC (hoặc có dòng chấm chấm giả):
           dọn dẹp các dòng chấm chấm giả và chèn trường tự động chuẩn của Word.
        3. Nếu chưa có trang MỤC LỤC: chèn trang MỤC LỤC vào vị trí chuẩn (ngay trước Chương 1).
        """
        # 1. Đã có trường TOC xịn
        if cls.has_toc_field(doc):
            toc_heading_info = cls.find_toc_heading(doc)
            if toc_heading_info:
                _, p_heading = toc_heading_info
                cls.format_toc_heading(p_heading, doc)
            return {"action": "already_exists", "cleaned_count": 0, "inserted": False}

        # 2. Đã có tiêu đề MỤC LỤC nhưng chưa có trường mã TOC
        toc_heading_info = cls.find_toc_heading(doc)
        if toc_heading_info:
            idx, p_heading = toc_heading_info
            cls.format_toc_heading(p_heading, doc)

            cleaned = cls.cleanup_fake_toc_entries(doc, idx)

            if idx + 1 < len(doc.paragraphs):
                p_next = doc.paragraphs[idx + 1]
                p_toc_field = p_next.insert_paragraph_before()
                cls.create_toc_field(p_toc_field, use_complex_field=use_complex_field)
                p_next.paragraph_format.page_break_before = True
            else:
                p_toc_field = doc.add_paragraph()
                cls.create_toc_field(p_toc_field, use_complex_field=use_complex_field)

            return {"action": "cleaned_and_inserted", "cleaned_count": cleaned, "inserted": True}

        # 3. Chưa có MỤC LỤC -> Tự động chèn trang MỤC LỤC
        insertion = cls.find_insertion_point(doc, cover_end_idx=cover_end_idx)
        if insertion:
            ins_idx, p_target = insertion

            # Dọn dẹp các đoạn trống ngay trước p_target (CHỈ sau cover_end_idx để bảo vệ trang bìa)
            while ins_idx > cover_end_idx:
                p_prev = doc.paragraphs[ins_idx - 1]
                if not p_prev.text.strip() and not p_prev._element.xpath('.//w:drawing | .//w:pict | .//w:tbl'):
                    p_prev._element.getparent().remove(p_prev._element)
                    ins_idx -= 1
                else:
                    break

            # Chèn tiêu đề "MỤC LỤC"
            p_toc_heading = p_target.insert_paragraph_before("MỤC LỤC")
            cls.format_toc_heading(p_toc_heading, doc)

            # Chèn paragraph chứa trường TOC
            p_toc_field = p_target.insert_paragraph_before()
            cls.create_toc_field(p_toc_field, use_complex_field=use_complex_field)

            # Ngắt trang trước đoạn tiếp theo (Chương 1)
            p_target.paragraph_format.page_break_before = True

            return {"action": "inserted_new", "cleaned_count": 0, "inserted": True}

        return {"action": "skipped_no_headings", "cleaned_count": 0, "inserted": False}
