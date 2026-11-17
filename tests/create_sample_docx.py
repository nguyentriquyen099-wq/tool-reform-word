"""
Script tạo file Word .docx mẫu giả lập báo cáo chưa chuẩn hóa thực tế
để kiểm thử toàn diện toàn bộ Pipeline.
"""

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def create_sample_messy_docx(output_path: str = "mau_bao_cao_chua_chuan_hoa.docx"):
    doc = Document()

    # 1. TRANG BÌA (Cover Zone)
    p0 = doc.add_paragraph("BỘ GIÁO DỤC VÀ ĐÀO TẠO")
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1 = doc.add_paragraph("TRƯỜNG ĐẠI HỌC BÁCH KHOA HÀ NỘI")
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2 = doc.add_paragraph()
    p3 = doc.add_paragraph("LUẬN VĂN THẠC SĨ KHOA HỌC")
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p4 = doc.add_paragraph("NGHIÊN CỨU ỨNG DỤNG TRÍ TUỆ NHÂN TẠO TRONG XỬ LÝ VĂN BẢN")
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p5 = doc.add_paragraph("Học viên thực hiện: Nguyễn Văn A")
    p6 = doc.add_paragraph("Người hướng dẫn: PGS. TS. Trần Văn B")
    p7 = doc.add_paragraph("Hà Nội, 2026")
    p7.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Section break ngắt bìa
    pPr = p7._element.get_or_add_pPr()
    sectPr = OxmlElement('w:sectPr')
    pPr.append(sectPr)

    # 2. PHẦN ĐẦU KHÔNG ĐÁNH SỐ (LỜI CẢM ƠN -> Non-TOC; LỜI MỞ ĐẦU -> In TOC)
    doc.add_paragraph("LỜI CẢM ƠN")
    doc.add_paragraph("Em xin chân thành cảm ơn các thầy cô giáo đã tận tình hướng dẫn trong suốt quá trình học tập...")

    doc.add_paragraph("LỜI MỞ ĐẦU")
    doc.add_paragraph("Trong thời đại bùng nổ thông tin và chuyển đổi số, việc tự động hóa chuẩn hóa tài liệu văn bản đóng vai trò vô cùng thiết yếu...")

    # 3. CHƯƠNG 1 (HEADING 1)
    doc.add_paragraph("CHƯƠNG 1: TỔNG QUAN HỆ THỐNG")
    p_lead = doc.add_paragraph()
    r_normal = p_lead.add_run("Hệ thống chuẩn hóa văn bản Word nhằm mục tiêu xử lý các định dạng phức tạp. Trong đó, ")
    r_bold = p_lead.add_run("tính toàn vẹn dữ liệu")
    r_bold.bold = True
    r_end = p_lead.add_run(" và ")
    r_italic = p_lead.add_run("trải nghiệm người dùng")
    r_italic.italic = True
    p_lead.add_run(" luôn được đặt lên hàng đầu.")

    # 1.1 HEADING 2
    doc.add_paragraph("1.1 Mục tiêu và phạm vi nghiên cứu")
    doc.add_paragraph("Nghiên cứu tập trung vào các định dạng báo cáo khoa học, đồ án tốt nghiệp, và văn bản hành chính theo tiêu chuẩn Việt Nam.")

    # 1.1.1 HEADING 3
    doc.add_paragraph("1.1.1 Yêu cầu kỹ thuật chi tiết")

    # DANH SÁCH BULLET RÁC & FRAGMENTED RUNS
    p_list1 = doc.add_paragraph()
    p_list1.add_run("\t   •  ")
    p_list1.add_run("Hỗ trợ phân loại thác nước 5 tầng trong ")
    r_algo = p_list1.add_run("thời gian thực O(N)")
    r_algo.bold = True
    p_list1.add_run(" mà không gây trễ.")

    p_list2 = doc.add_paragraph()
    # Ký hiệu bullet ❖
    p_list2.add_run("❖   Tự động phát hiện và khóa lề trang bìa.")

    # Danh sách cấp 2
    p_list3 = doc.add_paragraph()
    p_list3.add_run("    +   Chi tiết kiểm tra thẻ sectPr.")

    # BẪY NUMBERED LIST CON TRONG CHƯƠNG 1 (Phải không bị nhận nhầm làm Heading 1)
    p_trap1 = doc.add_paragraph("1. Bước một: Tiếp nhận file .docx từ người dùng qua giao diện Web kéo thả trực quan.")
    p_trap2 = doc.add_paragraph("2. Bước hai: Giải mã XML và đối chiếu styles.xml, numbering.xml để nhận diện cấu trúc.")

    # 4. CHƯƠNG 2 & BẢNG BIỂU & CAPTION
    doc.add_paragraph("CHƯƠNG 2: THỰC NGHIỆM VÀ KẾT QUẢ")
    doc.add_paragraph("Dưới đây là bảng tổng hợp kết quả thử nghiệm hiệu năng của hệ thống qua các kịch bản kiểm thử.")

    # Bảng biểu
    table = doc.add_table(rows=3, cols=3)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "STT"
    hdr_cells[1].text = "Tên kịch bản"
    hdr_cells[2].text = "Thời gian xử lý"

    row1 = table.rows[1].cells
    row1[0].text = "1"
    row1[1].text = "Tài liệu 50 trang"
    row1[2].text = "1.2s"

    row2 = table.rows[2].cells
    row2[0].text = "2"
    row2[1].text = "Tài liệu 200 trang"
    row2[2].text = "3.8s"

    # Caption Bảng
    doc.add_paragraph("Bảng 2.1: Kết quả đo đạc hiệu năng chuẩn hóa")

    # Hình ảnh / Caption Hình
    doc.add_paragraph("Hình 2.2: Sơ đồ luồng xử lý dữ liệu qua 6 phân hệ")

    # 5. KẾT LUẬN (TOC UNNUMBERED WHITELIST)
    doc.add_paragraph("KẾT LUẬN")
    doc.add_paragraph("Hệ thống đã chứng minh được tính hiệu quả, ổn định và tuân thủ nghiêm ngặt các quy định về văn bản.")

    # 6. TÀI LIỆU THAM KHẢO
    doc.add_paragraph("TÀI LIỆU THAM KHẢO")
    doc.add_paragraph("[1] Nghị định số 30/2020/NĐ-CP của Chính phủ về công tác văn thư.")
    doc.add_paragraph("[2] ISO/IEC 29500: Office Open XML File Formats.")

    doc.save(output_path)
    print(f"Đã tạo file mẫu thành công tại: {output_path}")
    return output_path

if __name__ == "__main__":
    create_sample_messy_docx()
