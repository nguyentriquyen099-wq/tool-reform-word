"""
Core Engine: DocumentStandardizer Pipeline
Tích hợp toàn diện 6 Phân hệ và 4 Lưu ý kỹ thuật cốt lõi:
1. Phân hệ 1: Document Zoning (Khóa Trang bìa, Tách Bảng biểu)
2. Phân hệ 2: Waterfall Classifier (Phân loại 5 tầng O(N))
3. Phân hệ 3: Hierarchical State Machine (HSM) & AI Fallback (DeepSeek)
4. Phân hệ 4: List Normalizer (Chuẩn hóa bullet '-' và '+')
5. Phân hệ 5: Run-level Manipulation & Style Reset (Bảo tồn bold/italic, không xóa run)
6. Phân hệ 6: Auto-TOC Update & Chống lệch trang (settings.xml updateFields, Margin A4)
"""

import os
import io
import time
import tempfile
import unicodedata
import logging
from typing import Dict, Any, Optional, List, Tuple
from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from config import Config
from core.numbering_parser import NumberingParser
from core.style_manager import StyleManager
from core.document_zoner import DocumentZoner
from core.classifier import WaterfallClassifier, ParagraphType
from core.hsm import HierarchicalStateMachine, AIFallbackClassifier
from core.list_normalizer import ListNormalizer
from core.run_manipulator import RunManipulator, int_to_roman
from core.toc_manager import TOCManager

logger = logging.getLogger(__name__)

class DocumentStandardizer:
    def __init__(
        self,
        api_key: Optional[str] = None,
        base_url: Optional[str] = None,
        model: Optional[str] = None,
        normalize_chapter_to_roman: Optional[bool] = None,
        normalize_heading_caps: Optional[bool] = None
    ):
        self.api_key = api_key or Config.DEEPSEEK_API_KEY
        self.base_url = base_url or Config.DEEPSEEK_BASE_URL
        self.model = model or Config.DEEPSEEK_MODEL
        self.normalize_chapter_to_roman = (
            normalize_chapter_to_roman
            if normalize_chapter_to_roman is not None
            else Config.NORMALIZE_CHAPTER_TO_ROMAN
        )
        self.normalize_heading_caps = (
            normalize_heading_caps
            if normalize_heading_caps is not None
            else Config.NORMALIZE_HEADING_CAPS
        )
        self.ai_classifier = AIFallbackClassifier(api_key=self.api_key, base_url=self.base_url, model=self.model)

    @staticmethod
    def _count_tables_recursive(tables) -> int:
        """Đếm chính xác toàn bộ bảng biểu bao gồm cả nested tables"""
        count = 0
        for tbl in tables:
            count += 1
            for row in tbl.rows:
                for cell in row.cells:
                    if cell.tables:
                        count += DocumentStandardizer._count_tables_recursive(cell.tables)
        return count

    def process(
        self,
        input_source,
        output_target=None,
        log_callback=None,
        auto_insert_toc: bool = True,
        evaluate_toc: Optional[bool] = None,
        normalize_chapter_to_roman: Optional[bool] = None,
        normalize_heading_caps: Optional[bool] = None
    ) -> Dict[str, Any]:
        """
        Thực hiện toàn bộ quy trình chuẩn hóa tài liệu Word (.docx).
        input_source: Đường dẫn file (str/Path) hoặc bytes / BytesIO
        output_target: Đường dẫn file lưu kết quả hoặc None (trả về BytesIO)
        log_callback: Hàm nhận log trạng thái callback(str, progress_float)
        auto_insert_toc: Tự động phát hiện và chèn trang Mục lục OpenXML nếu chưa có (mặc định True)
        evaluate_toc: Tự động gọi Microsoft Word / LibreOffice để kết xuất số trang TOC (mặc định bật khi lưu file)
        normalize_chapter_to_roman: Tự động chuẩn hóa số Chương từ số Ả Rập sang số La Mã (CHƯƠNG I, II, III...)
        normalize_heading_caps: Tự động chuẩn hóa Heading cấp con bị gõ ALL CAPS về Sentence Case
        """
        if evaluate_toc is None:
            evaluate_toc = (output_target is not None)
        norm_chapter = (
            normalize_chapter_to_roman
            if normalize_chapter_to_roman is not None
            else self.normalize_chapter_to_roman
        )
        norm_caps = (
            normalize_heading_caps
            if normalize_heading_caps is not None
            else self.normalize_heading_caps
        )
        def emit_log(msg: str, progress: float = 0.0):
            logger.info(msg)
            if log_callback:
                log_callback(msg, progress)

        emit_log("Bắt đầu mở và nạp tài liệu Word (.docx)...", 0.05)
        if hasattr(input_source, 'paragraphs') and hasattr(input_source, 'tables'):
            doc = input_source
        elif isinstance(input_source, bytes):
            doc = Document(io.BytesIO(input_source))
        else:
            doc = Document(input_source)

        stats = {
            "total_paragraphs": len(doc.paragraphs),
            "cover_paragraphs": 0,
            "tables_count": self._count_tables_recursive(doc.tables),
            "media_math_count": 0,
            "caption_count": 0,
            "heading_1_count": 0,
            "heading_2_count": 0,
            "heading_3_count": 0,
            "heading_4_count": 0,
            "heading_unnumbered_count": 0,
            "chapters_roman_count": 0,
            "headings_casing_normalized_count": 0,
            "list_items_count": 0,
            "body_paragraphs_count": 0,
            "ai_classified_count": 0,
            "toc_flag_enabled": False,
            "toc_inserted": False,
            "toc_cleaned_count": 0,
            "toc_evaluated": False,
            "libreoffice_executed": False,
            "processing_time": 0.0
        }

        start_time = time.time()

        # BƯỚC 1: Phân tích numbering.xml & Cấu hình Styles
        emit_log("Khởi tạo bộ phân tích numbering.xml và thiết lập Styles chuẩn...", 0.15)
        num_parser = NumberingParser(doc)
        style_mgr = StyleManager(doc)
        style_mgr.setup_document_styles()

        # BƯỚC 2: Chuẩn hóa lề trang A4 chống lệch số trang (Lưu ý 4 & Phân hệ 6)
        emit_log(f"Chuẩn hóa lề trang A4 (Top {Config.MARGIN_TOP_CM}cm, Bottom {Config.MARGIN_BOTTOM_CM}cm, Left {Config.MARGIN_LEFT_CM}cm, Right {Config.MARGIN_RIGHT_CM}cm)...", 0.25)
        TOCManager.standardize_page_margins(doc)

        # BƯỚC 3: Phân hệ 1 - Document Zoning (Trang bìa & Bảng biểu)
        emit_log("Phân hệ 1: Xác định vùng Trang bìa và tách biệt Bảng biểu...", 0.35)
        zoner = DocumentZoner(doc)
        cover_end_idx, cover_paragraphs = zoner.identify_zones()
        zoner.format_cover_page()
        zoner.format_tables()
        stats["cover_paragraphs"] = len(cover_paragraphs)

        # BƯỚC 3.5: Phân hệ 6 - Đảm bảo Trang Mục lục (Auto-TOC Generation & Cleanup)
        if auto_insert_toc:
            emit_log("Phân hệ 6: Kiểm tra và chuẩn hóa Trang Mục lục tự động...", 0.40)
            toc_result = TOCManager.ensure_table_of_contents(doc, cover_end_idx=cover_end_idx)
            stats["toc_inserted"] = toc_result.get("inserted", False)
            stats["toc_cleaned_count"] = toc_result.get("cleaned_count", 0)
            if toc_result.get("action") == "inserted_new":
                emit_log("Phân hệ 6: Đã tự động tạo trang Mục lục chuẩn OpenXML trước phần nội dung chính.", 0.42)
            elif toc_result.get("action") == "cleaned_and_inserted":
                emit_log(f"Phân hệ 6: Đã dọn {stats['toc_cleaned_count']} dòng chấm chấm thủ công và chèn trường Mục lục tự động.", 0.42)
            elif toc_result.get("action") == "already_exists":
                emit_log("Phân hệ 6: Tài liệu đã có trường Mục lục chuẩn.", 0.42)

        # BƯỚC 4: Phân hệ 2 - Waterfall Classifier (1 vòng lặp O(N))
        emit_log("Phân hệ 2: Bắt đầu duyệt phân loại Thác nước (Waterfall Classifier)...", 0.45)
        classifier = WaterfallClassifier(
            doc,
            numbering_parser=num_parser,
            ai_classifier=self.ai_classifier,
            normalize_chapter_to_roman=norm_chapter
        )

        ai_candidates = []
        body_start_idx = max(0, cover_end_idx)
        total_body = len(doc.paragraphs) - body_start_idx

        for i in range(body_start_idx, len(doc.paragraphs)):
            p = doc.paragraphs[i]
            text = p.text.strip()

            # Bỏ qua đoạn hoàn toàn trống (ngoại trừ media/math hoặc đoạn chứa trường TOC)
            if not text and not classifier.is_tier1_media_or_math(p) and not TOCManager.is_toc_paragraph(p):
                RunManipulator.clean_trailing_whitespace(p)
                RunManipulator.reset_run_style_inheritance(p)
                continue

            # Đoạn chứa trường TOC (Mục lục) -> Giữ nguyên, không áp đặt thụt lề 1.27cm của Body text
            if TOCManager.is_toc_paragraph(p):
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.first_line_indent = Inches(0)
                p.paragraph_format.left_indent = Inches(0)
                p.paragraph_format.right_indent = Inches(0)
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(12)
                continue

            # Tầng 1: Media / Math
            if classifier.is_tier1_media_or_math(p):
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.first_line_indent = Inches(0)
                p.paragraph_format.left_indent = Inches(0)
                p.paragraph_format.right_indent = Inches(0)
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(6)
                RunManipulator.clean_trailing_whitespace(p)
                RunManipulator.reset_run_style_inheritance(p, keep_bold=True, keep_italic=True)
                stats["media_math_count"] += 1
                continue

            # Tầng 2: Caption
            if classifier.is_tier2_caption(p):
                cap_style_name = 'Caption'
                p_text_norm = unicodedata.normalize('NFC', p.text.strip().lower())
                if any(p_text_norm.startswith(k) for k in ['hình', 'hinh', 'sơ đồ', 'so do', 'ảnh', 'anh', 'đồ thị', 'do thi']):
                    if 'hinhanh' in doc.styles:
                        cap_style_name = 'hinhanh'
                elif any(p_text_norm.startswith(k) for k in ['bảng', 'bang', 'biểu đồ', 'bieu do']):
                    if 'bang' in doc.styles:
                        cap_style_name = 'bang'
                try:
                    p.style = doc.styles[cap_style_name]
                except (KeyError, ValueError):
                    try:
                        p.style = doc.styles['Caption']
                    except Exception:
                        pass
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.first_line_indent = Inches(0)
                p.paragraph_format.left_indent = Inches(0)
                p.paragraph_format.right_indent = Inches(0)
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(6)
                for run in p.runs:
                    run.italic = True
                    run.font.name = Config.FONT_NAME
                    run.font.size = Pt(Config.CAPTION_FONT_SIZE)
                stats["caption_count"] += 1
                continue

            # Tầng 3: Heading
            is_heading, heading_type, heading_lvl = classifier.evaluate_tier3_heading(p)
            if is_heading:
                if norm_chapter and (heading_type == ParagraphType.HEADING_1 or heading_lvl == 1):
                    if RunManipulator.convert_chapter_number_to_roman(p):
                        stats["chapters_roman_count"] += 1
                self._apply_heading_formatting(
                    doc, p, heading_type, heading_lvl,
                    normalize_chapter_to_roman=norm_chapter,
                    normalize_heading_caps=norm_caps,
                    stats=stats
                )
                if heading_type == ParagraphType.HEADING_1:
                    stats["heading_1_count"] += 1
                elif heading_type == ParagraphType.HEADING_2:
                    stats["heading_2_count"] += 1
                elif heading_type == ParagraphType.HEADING_3:
                    stats["heading_3_count"] += 1
                elif heading_type == ParagraphType.HEADING_4:
                    stats["heading_4_count"] += 1
                elif heading_type in [ParagraphType.HEADING_UNNUMBERED, ParagraphType.HEADING_NON_TOC]:
                    stats["heading_unnumbered_count"] += 1
                continue

            # Kiểm tra ứng viên nghi vấn cho AI Fallback trước khi gán List/Body
            if classifier.is_ambiguous_candidate(p):
                prev_text = doc.paragraphs[i-1].text[:80] if i > 0 else ""
                next_text = doc.paragraphs[i+1].text[:80] if i + 1 < len(doc.paragraphs) else ""
                ai_candidates.append({
                    "id": i,
                    "text": text,
                    "is_bold": True,
                    "prev_text": prev_text,
                    "next_text": next_text
                })
                # Đánh dấu tạm thời, sẽ xử lý sau batch
                continue

            # Tầng 4: List items
            is_list, list_lvl = classifier.is_tier4_list(p)
            if is_list:
                ListNormalizer.normalize_list_item(p, level=list_lvl)
                stats["list_items_count"] += 1
                continue

            # Tầng 5: Body Text
            self._apply_body_formatting(doc, p)
            stats["body_paragraphs_count"] += 1

        # BƯỚC 5: Phân hệ 3 - AI Fallback xử lý danh sách ứng viên nghi vấn
        if ai_candidates:
            emit_log(f"Phân hệ 3: AI Fallback phân tích {len(ai_candidates)} đoạn nghi vấn...", 0.70)
            ai_results = self.ai_classifier.classify_candidates(ai_candidates)
            stats["ai_classified_count"] = len(ai_results)

            for cand in ai_candidates:
                c_id = cand["id"]
                p = doc.paragraphs[c_id]
                c_type = ai_results.get(c_id, "BODY")

                if c_type == "HEADING_1":
                    if norm_chapter:
                        if RunManipulator.convert_chapter_number_to_roman(p):
                            stats["chapters_roman_count"] += 1
                    self._apply_heading_formatting(
                        doc, p, ParagraphType.HEADING_1, 1,
                        normalize_chapter_to_roman=norm_chapter,
                        normalize_heading_caps=norm_caps,
                        stats=stats
                    )
                    stats["heading_1_count"] += 1
                elif c_type == "HEADING_2":
                    self._apply_heading_formatting(
                        doc, p, ParagraphType.HEADING_2, 2,
                        normalize_heading_caps=norm_caps,
                        stats=stats
                    )
                    stats["heading_2_count"] += 1
                elif c_type == "HEADING_3":
                    self._apply_heading_formatting(
                        doc, p, ParagraphType.HEADING_3, 3,
                        normalize_heading_caps=norm_caps,
                        stats=stats
                    )
                    stats["heading_3_count"] += 1
                elif c_type == "HEADING_4":
                    self._apply_heading_formatting(
                        doc, p, ParagraphType.HEADING_4, 4,
                        normalize_heading_caps=norm_caps,
                        stats=stats
                    )
                    stats["heading_4_count"] += 1
                elif c_type == "HEADING_UNNUMBERED":
                    self._apply_heading_formatting(
                        doc, p, ParagraphType.HEADING_UNNUMBERED, 1,
                        normalize_heading_caps=norm_caps,
                        stats=stats
                    )
                    stats["heading_unnumbered_count"] += 1
                elif c_type == "HEADING_NON_TOC":
                    self._apply_heading_formatting(
                        doc, p, ParagraphType.HEADING_NON_TOC, 1,
                        normalize_heading_caps=norm_caps,
                        stats=stats
                    )
                    stats["heading_unnumbered_count"] += 1
                elif c_type == "CAPTION":
                    cap_style_name = 'Caption'
                    p_text_norm = unicodedata.normalize('NFC', p.text.strip().lower())
                    if any(p_text_norm.startswith(k) for k in ['hình', 'hinh', 'sơ đồ', 'so do', 'ảnh', 'anh', 'đồ thị', 'do thi']):
                        if 'hinhanh' in doc.styles:
                            cap_style_name = 'hinhanh'
                    elif any(p_text_norm.startswith(k) for k in ['bảng', 'bang', 'biểu đồ', 'bieu do']):
                        if 'bang' in doc.styles:
                            cap_style_name = 'bang'
                    try:
                        p.style = doc.styles[cap_style_name]
                    except (KeyError, ValueError):
                        try:
                            p.style = doc.styles['Caption']
                        except Exception:
                            pass
                    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.first_line_indent = Inches(0)
                    p.paragraph_format.left_indent = Inches(0)
                    p.paragraph_format.right_indent = Inches(0)
                    p.paragraph_format.space_before = Pt(6)
                    p.paragraph_format.space_after = Pt(6)
                    for r in p.runs:
                        r.italic = True
                        r.font.name = Config.FONT_NAME
                        r.font.size = Pt(Config.CAPTION_FONT_SIZE)
                    stats["caption_count"] += 1
                elif c_type == "LIST":
                    ListNormalizer.normalize_list_item(p, level=1)
                    stats["list_items_count"] += 1
                else:
                    self._apply_body_formatting(doc, p)
                    stats["body_paragraphs_count"] += 1

        # BƯỚC 6: Phân hệ 6 - Kích hoạt cờ updateFields & Tự động kết xuất Mục lục
        emit_log("Phân hệ 6: Thiết lập cờ tự động cập nhật Mục lục (<w:updateFields/>)...", 0.88)
        stats["toc_flag_enabled"] = TOCManager.enable_update_fields_on_open(doc)

        emit_log("Đang lưu trữ và cập nhật hiển thị Mục lục...", 0.92)

        if output_target:
            doc.save(output_target)
            if evaluate_toc:
                eval_ok, eval_msg = TOCManager.update_toc_headless(output_target)
                stats["toc_evaluated"] = eval_ok
                stats["libreoffice_executed"] = eval_ok
                if eval_ok:
                    emit_log(f"Phân hệ 6: {eval_msg}", 0.96)
            stats["processing_time"] = round(time.time() - start_time, 2)
            emit_log(f"Hoàn thành chuẩn hóa trong {stats['processing_time']}s! File đã sẵn sàng.", 1.0)
            return {"stats": stats, "output_path": output_target}
        else:
            # Chế độ In-memory (Streamlit hoặc testing)
            if evaluate_toc:
                temp_path = None
                try:
                    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp:
                        temp_path = tmp.name
                    doc.save(temp_path)
                    eval_ok, eval_msg = TOCManager.update_toc_headless(temp_path)
                    stats["toc_evaluated"] = eval_ok
                    stats["libreoffice_executed"] = eval_ok
                    if eval_ok:
                        emit_log(f"Phân hệ 6: {eval_msg}", 0.96)
                    with open(temp_path, 'rb') as f:
                        buffer = io.BytesIO(f.read())
                except Exception as e:
                    logger.warning(f"Lỗi khi kết xuất TOC headless in-memory: {e}")
                    buffer = io.BytesIO()
                    doc.save(buffer)
                    buffer.seek(0)
                finally:
                    if temp_path and os.path.exists(temp_path):
                        try:
                            os.remove(temp_path)
                        except Exception:
                            pass
            else:
                buffer = io.BytesIO()
                doc.save(buffer)
                buffer.seek(0)

            stats["processing_time"] = round(time.time() - start_time, 2)
            emit_log(f"Hoàn thành chuẩn hóa trong {stats['processing_time']}s! Sẵn sàng tải xuống.", 1.0)
            return {"stats": stats, "buffer": buffer}

    @staticmethod
    def _set_outline_level(paragraph, val: str):
        """Chủ động gán thuộc tính w:outlineLvl vào w:pPr của paragraph để TOC & Navigation Pane luôn hiển thị chuẩn"""
        pPr = paragraph._element.get_or_add_pPr()
        outline = pPr.find(qn('w:outlineLvl'))
        if outline is None:
            outline = OxmlElement('w:outlineLvl')
            pPr.append(outline)
        outline.set(qn('w:val'), val)

    def _apply_heading_formatting(
        self,
        doc,
        paragraph,
        heading_type: str,
        level: int = 1,
        normalize_chapter_to_roman: Optional[bool] = None,
        normalize_heading_caps: Optional[bool] = None,
        stats: Optional[Dict[str, Any]] = None
    ):
        """
        Áp dụng định dạng Heading theo đúng quy chuẩn phân cấp và Whitelist
        """
        norm_chap = (
            normalize_chapter_to_roman
            if normalize_chapter_to_roman is not None
            else getattr(self, 'normalize_chapter_to_roman', False)
        )
        norm_caps = (
            normalize_heading_caps
            if normalize_heading_caps is not None
            else getattr(self, 'normalize_heading_caps', True)
        )
        # Nếu bật cấu hình chuẩn hóa số Chương sang số La Mã
        if norm_chap and (heading_type == ParagraphType.HEADING_1 or level == 1):
            RunManipulator.convert_chapter_number_to_roman(paragraph)

        # Làm sạch rác/khoảng trắng thừa ở đầu và cuối
        RunManipulator.clean_leading_junk_and_symbols(paragraph, remove_bullets=False)
        RunManipulator.clean_trailing_whitespace(paragraph)

        # Chuẩn hóa chữ hoa cho Heading cấp con (Heading 2, 3...) nếu bị gõ ALL CAPS
        if norm_caps and (level >= 2 or heading_type in [ParagraphType.HEADING_2, ParagraphType.HEADING_3, ParagraphType.HEADING_4]):
            if RunManipulator.normalize_heading_casing(paragraph):
                if stats is not None and "headings_casing_normalized_count" in stats:
                    stats["headings_casing_normalized_count"] += 1

        # Reset override rác cũ (màu sắc, font lạ) trước khi gán định dạng Heading chuẩn
        RunManipulator.reset_run_style_inheritance(paragraph, keep_bold=True, keep_italic=True)

        pf = paragraph.paragraph_format
        pf.first_line_indent = Inches(0)
        pf.right_indent = Inches(0)
        pf.keep_with_next = True
        pf.line_spacing = Config.LINE_SPACING

        def _format_heading_runs(target_size: float, is_italic: bool = False):
            half_pts = str(int(round(target_size * 2)))
            for run in paragraph.runs:
                run.bold = True
                run.italic = is_italic
                run.font.name = Config.FONT_NAME
                run.font.size = Pt(target_size)
                rPr = run._element.get_or_add_rPr()
                szCs = rPr.find(qn('w:szCs'))
                if szCs is not None:
                    szCs.set(qn('w:val'), half_pts)
                if is_italic:
                    if rPr.find(qn('w:i')) is None:
                        rPr.append(OxmlElement('w:i'))
                    if rPr.find(qn('w:iCs')) is None:
                        rPr.append(OxmlElement('w:iCs'))
                else:
                    i_elem = rPr.find(qn('w:i'))
                    if i_elem is not None:
                        rPr.remove(i_elem)
                    iCs = rPr.find(qn('w:iCs'))
                    if iCs is not None:
                        rPr.remove(iCs)
                if rPr.find(qn('w:bCs')) is None:
                    rPr.append(OxmlElement('w:bCs'))

        if heading_type == ParagraphType.HEADING_UNNUMBERED:
            try:
                paragraph.style = doc.styles[Config.STYLE_HEADING_1_UNNUMBERED]
            except KeyError:
                pass
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.left_indent = Cm(0)
            pf.space_before = Pt(Config.HEADING_1_SPACE_BEFORE_PT)
            pf.space_after = Pt(Config.HEADING_1_SPACE_AFTER_PT)
            self._set_outline_level(paragraph, '0')
            _format_heading_runs(Config.HEADING_1_FONT_SIZE, is_italic=False)

        elif heading_type == ParagraphType.HEADING_NON_TOC:
            try:
                paragraph.style = doc.styles[Config.STYLE_NON_TOC_HEADING]
            except KeyError:
                pass
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.left_indent = Cm(0)
            pf.space_before = Pt(Config.HEADING_1_SPACE_BEFORE_PT)
            pf.space_after = Pt(Config.HEADING_1_SPACE_AFTER_PT)
            # Đảm bảo gỡ outlineLvl để không lọt vào TOC
            pPr = paragraph._element.get_or_add_pPr()
            outline = pPr.find(qn('w:outlineLvl'))
            if outline is not None:
                pPr.remove(outline)
            _format_heading_runs(Config.HEADING_1_FONT_SIZE, is_italic=False)

        elif heading_type == ParagraphType.HEADING_1 or level == 1:
            try:
                paragraph.style = doc.styles['Heading 1']
            except KeyError:
                pass
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.left_indent = Cm(0)
            pf.space_before = Pt(Config.HEADING_1_SPACE_BEFORE_PT)
            pf.space_after = Pt(Config.HEADING_1_SPACE_AFTER_PT)
            self._set_outline_level(paragraph, '0')
            _format_heading_runs(Config.HEADING_1_FONT_SIZE, is_italic=False)

        elif heading_type == ParagraphType.HEADING_2 or level == 2:
            try:
                paragraph.style = doc.styles['Heading 2']
            except KeyError:
                pass
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf.left_indent = Cm(Config.HEADING_INDENT_PER_LEVEL_CM * 1)
            pf.space_before = Pt(Config.HEADING_SPACE_BEFORE_PT)
            pf.space_after = Pt(Config.HEADING_SPACE_AFTER_PT)
            self._set_outline_level(paragraph, '1')
            _format_heading_runs(Config.HEADING_2_FONT_SIZE, is_italic=False)

        elif heading_type == ParagraphType.HEADING_3 or level == 3:
            try:
                paragraph.style = doc.styles['Heading 3']
            except KeyError:
                pass
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf.left_indent = Cm(Config.HEADING_INDENT_PER_LEVEL_CM * 2)
            pf.space_before = Pt(Config.HEADING_SPACE_BEFORE_PT)
            pf.space_after = Pt(Config.HEADING_SPACE_AFTER_PT)
            self._set_outline_level(paragraph, '2')
            _format_heading_runs(Config.HEADING_3_FONT_SIZE, is_italic=True)

        elif heading_type == ParagraphType.HEADING_4 or level == 4:
            try:
                paragraph.style = doc.styles['Heading 4']
            except KeyError:
                try:
                    paragraph.style = doc.styles['Heading 3']
                except KeyError:
                    pass
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf.left_indent = Cm(Config.HEADING_INDENT_PER_LEVEL_CM * 3)
            pf.space_before = Pt(Config.HEADING_SPACE_BEFORE_PT)
            pf.space_after = Pt(Config.HEADING_SPACE_AFTER_PT)
            self._set_outline_level(paragraph, '3')
            _format_heading_runs(Config.HEADING_4_FONT_SIZE, is_italic=True)

        else: # Cấp sâu hơn (> 4), fallback về định dạng chuẩn cấp 4
            try:
                paragraph.style = doc.styles['Heading 4']
            except KeyError:
                try:
                    paragraph.style = doc.styles['Heading 3']
                except KeyError:
                    pass
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf.left_indent = Cm(Config.HEADING_INDENT_PER_LEVEL_CM * (level - 1))
            pf.space_before = Pt(Config.HEADING_SPACE_BEFORE_PT)
            pf.space_after = Pt(Config.HEADING_SPACE_AFTER_PT)
            self._set_outline_level(paragraph, str(min(level - 1, 8)))
            _format_heading_runs(Config.HEADING_4_FONT_SIZE, is_italic=True)

    def _apply_body_formatting(self, doc, paragraph):
        """
        Áp dụng định dạng Body Text:
        - Xóa tab thừa, khoảng trắng thụt lề thừa ở đầu
        - Căn đều (Justify)
        - Thụt lề đầu dòng: 1.27cm
        - Giãn dòng 1.3, Space after 6pt
        - Reset style inheritance nhưng bảo tồn triệt để bold/italic nội tại
        """
        RunManipulator.clean_leading_junk_and_symbols(paragraph, remove_bullets=False)

        try:
            paragraph.style = doc.styles['Normal']
        except Exception:
            try:
                paragraph.style = 'Normal'
            except Exception:
                pass

        pf = paragraph.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf.first_line_indent = Cm(Config.FIRST_LINE_INDENT_CM)
        pf.left_indent = Cm(0)
        pf.right_indent = Cm(0)
        pf.line_spacing = Config.LINE_SPACING
        pf.space_before = Pt(Config.SPACE_BEFORE_PT)
        pf.space_after = Pt(Config.SPACE_AFTER_PT)

        RunManipulator.reset_run_style_inheritance(paragraph, keep_bold=True, keep_italic=True)
