"""
Giao diện Web Trực quan Chuẩn hóa Báo cáo Word (.docx)
Sử dụng Streamlit: Kéo thả file .docx, cấu hình DeepSeek AI, Cổng Thanh toán QR Banking (VietQR & Local QR), theo dõi log và tải file hoàn thiện.
"""

import os
import io
import time
from pathlib import Path
import streamlit as st
from config import Config, POPULAR_BANKS
from core.pipeline import DocumentStandardizer
from core.toc_manager import TOCManager
from core.payment_verifier import PaymentVerifier, check_sepay_payment

# Cấu hình giao diện Streamlit
st.set_page_config(
    page_title="Hệ thống Chuẩn hóa Báo cáo Word (.docx)",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS cho giao diện hiện đại, chuyên nghiệp
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #4B5563;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 12px;
        text-align: center;
    }
    .metric-val {
        font-size: 1.6rem;
        font-weight: bold;
        color: #2563EB;
    }
    .metric-label {
        font-size: 0.85rem;
        color: #64748B;
    }
    .stAlert {
        border-radius: 8px;
    }
    .payment-box {
        background: linear-gradient(135deg, #F8FAFC 0%, #EFF6FF 100%);
        border: 2px solid #3B82F6;
        border-radius: 12px;
        padding: 20px;
        margin-top: 15px;
        margin-bottom: 20px;
        box-shadow: 0 4px 12px rgba(59, 130, 246, 0.08);
    }
    .payment-header {
        font-size: 1.3rem;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 6px;
    }
    .payment-sub {
        font-size: 0.95rem;
        color: #4B5563;
        margin-bottom: 16px;
    }
</style>
""", unsafe_allow_html=True)

# SIDEBAR: Cấu hình hệ thống & DeepSeek API & Cổng Thanh toán QR
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/ms-word.png", width=64)
    st.title("Cấu hình Hệ thống")
    st.markdown("---")

    st.subheader("🤖 Cấu hình AI DeepSeek")
    env_api_key = Config.DEEPSEEK_API_KEY
    if not env_api_key:
        try:
            if hasattr(st, "secrets") and "DEEPSEEK_API_KEY" in st.secrets:
                env_api_key = str(st.secrets["DEEPSEEK_API_KEY"]).strip()
        except Exception:
            pass

    api_key_input = st.text_input(
        "DeepSeek API Key:",
        value=env_api_key,
        type="password",
        help="API Key dùng để phân loại tự động các đề mục ngắn/nghi vấn bằng DeepSeek AI."
    )

    if api_key_input:
        st.success("✅ Đã sẵn sàng kết nối DeepSeek API")
    else:
        st.info("ℹ️ Chưa có API Key. Hệ thống sẽ sử dụng Heuristic Fallback thông minh (vẫn hoạt động tốt).")

    with st.expander("⚙️ Tùy chọn nâng cao API"):
        base_url_input = st.text_input("DeepSeek Base URL:", value=Config.DEEPSEEK_BASE_URL)
        model_input = st.text_input("DeepSeek Model:", value=Config.DEEPSEEK_MODEL)

    st.markdown("---")
    st.subheader("💳 Cổng Thanh Toán QR Banking")
    enable_payment = st.checkbox(
        "Bật Cổng Thanh toán QR khi xuất file",
        value=Config.ENABLE_PAYMENT_GATEWAY,
        help="Khi chuẩn hóa thành công, hiển thị khung quét mã QR thanh toán/ủng hộ trước khi tải tài liệu hoàn thiện."
    )

    if enable_payment:
        bank_keys = list(POPULAR_BANKS.keys())
        current_cfg_bank = Config.BANK_ID.strip().upper()
        is_in_popular = current_cfg_bank in bank_keys
        default_idx = bank_keys.index(current_cfg_bank) if is_in_popular else 0

        bank_select = st.selectbox(
            "Chọn Ngân hàng:",
            options=bank_keys,
            index=default_idx,
            format_func=lambda k: f"{k} - {POPULAR_BANKS.get(k, '')}"
        )
        custom_bank_default = current_cfg_bank if (current_cfg_bank and not is_in_popular) else ""
        custom_bank = st.text_input(
            "Hoặc nhập Mã Ngân hàng khác:",
            value=custom_bank_default,
            help="Nếu ngân hàng của bạn không có trong danh sách trên, hãy nhập mã (VD: NAB, TIMO, CAKE, BIDC...)"
        ).strip().upper()
        bank_id_input = custom_bank if custom_bank else bank_select

        account_no_input = st.text_input("Số tài khoản nhận tiền:", value=Config.ACCOUNT_NO).strip()
        account_name_input = st.text_input("Tên chủ tài khoản:", value=Config.ACCOUNT_NAME).strip()
        payment_amount_input = int(st.number_input(
            "Số tiền thanh toán / ủng hộ (VNĐ):",
            value=int(Config.PAYMENT_AMOUNT),
            step=1000,
            min_value=0
        ))

        payment_mode_options = ["auto", "manual"]
        default_mode_idx = 0 if Config.PAYMENT_MODE == "auto" else 1
        payment_mode_input = st.radio(
            "Chế độ Xác nhận Thanh toán:",
            options=payment_mode_options,
            index=default_mode_idx,
            format_func=lambda m: "⚡ Cấp độ 2: Tự động check tiền vào (SePay API)" if m == "auto" else "🖐️ Cấp độ 1: Bấm nút xác nhận thủ công",
            help="Cấp độ 2: Tự động kết nối ngân hàng qua SePay API, tiền vào tài khoản là tự động mở khóa nút tải file."
        )

        if payment_mode_input == "auto":
            sepay_token_input = st.text_input(
                "SePay API Token:",
                value=Config.SEPAY_API_TOKEN,
                type="password",
                help="Token lấy tại https://my.sepay.vn -> Cấu hình công ty -> Truy cập API."
            ).strip()
            test_mode_input = st.checkbox(
                "🧪 Bật Chế độ Thử nghiệm (Sandbox/Mock)",
                value=Config.SEPAY_TEST_MODE,
                help="Cho phép mô phỏng thanh toán thành công để kiểm tra toàn bộ luồng mà không cần nạp tiền thật."
            )
            with st.expander("📖 Hướng dẫn 3 bước kết nối SePay (2 phút)"):
                st.markdown("""
                **3 bước để tiền về thẳng tài khoản ngân hàng của bạn:**
                1. **Đăng ký tài khoản miễn phí:** Truy cập [my.sepay.vn](https://my.sepay.vn) và liên kết tài khoản ngân hàng của bạn.
                2. **Tạo API Token:** Vào **Cấu hình công ty** ➡️ **Truy cập API** ➡️ Bấm **+ Thêm API** và sao chép mã Token.
                3. **Dán Token vào đây:** Dán mã Token vào ô **SePay API Token** bên trên (hoặc lưu vào file `.env`). Xong! Hệ thống sẽ tự động quét biến động số dư và mở khóa tải file khi tiền về tài khoản.
                """)
        else:
            sepay_token_input = ""
            test_mode_input = False

        local_qr_found = Config.find_local_qr_image()
        if local_qr_found:
            st.success(f"🖼️ Nhận diện ảnh QR: `{Path(local_qr_found).name}`")
        else:
            st.caption("💡 Mẹo: Bỏ file ảnh `qr.png` hoặc `qr.jpg` vào thư mục gốc để hiển thị ảnh QR cá nhân của bạn.")
    else:
        bank_id_input = Config.BANK_ID
        account_no_input = Config.ACCOUNT_NO
        account_name_input = Config.ACCOUNT_NAME
        payment_amount_input = int(Config.PAYMENT_AMOUNT)
        payment_mode_input = "manual"
        sepay_token_input = ""
        test_mode_input = False

    st.markdown("---")
    st.subheader("⚙️ Tùy chọn Chuẩn hóa")
    normalize_chapter_roman = st.checkbox(
        "Chuẩn hóa số Chương sang số La Mã (CHƯƠNG I, II, III...)",
        value=Config.NORMALIZE_CHAPTER_TO_ROMAN,
        help="Chuyển đổi tiêu đề Chương từ số Ả Rập (CHƯƠNG 1) sang số La Mã (CHƯƠNG I, II, III...). Mặc định tắt theo Mẫu form chuẩn."
    )
    normalize_heading_caps = st.checkbox(
        "Chuẩn hóa chữ hoa Heading cấp con (Sentence Case)",
        value=Config.NORMALIZE_HEADING_CAPS,
        help="Tự động chuẩn hóa Heading 2, 3... nếu bị gõ ALL CAPS về dạng Sentence Case chuẩn (chữ thường, chỉ viết hoa đầu câu), giữ nguyên từ viết tắt (LAN, IP, CNTT...) để đồng nhất kích cỡ thị giác."
    )

    st.markdown("---")
    st.subheader("📐 Tiêu chuẩn Định dạng Mẫu Form")
    st.markdown(f"""
    - **Font chữ:** {Config.FONT_NAME}
    - **Cỡ chữ Body:** {int(Config.BODY_FONT_SIZE)} pt (Justify, Thụt lề {Config.FIRST_LINE_INDENT_CM}cm, Spacing {Config.LINE_SPACING}, After {int(Config.SPACE_AFTER_PT)}pt)
    - **Cỡ chữ Heading 1:** {int(Config.HEADING_1_FONT_SIZE)} pt (In đậm, Căn giữa, {int(Config.HEADING_1_SPACE_BEFORE_PT)}pt before, {int(Config.HEADING_1_SPACE_AFTER_PT)}pt after)
    - **Cỡ chữ Heading 2:** {int(Config.HEADING_2_FONT_SIZE)} pt (In đậm, Căn trái, {int(Config.HEADING_SPACE_BEFORE_PT)}pt before, {int(Config.HEADING_SPACE_AFTER_PT)}pt after)
    - **Cỡ chữ Heading 3:** {int(Config.HEADING_3_FONT_SIZE)} pt (In đậm + In nghiêng, Căn trái, {int(Config.HEADING_SPACE_BEFORE_PT)}pt before, {int(Config.HEADING_SPACE_AFTER_PT)}pt after)
    - **Cỡ chữ Caption:** {int(Config.CAPTION_FONT_SIZE)} pt (In nghiêng, Căn giữa, Style hinhanh/bang)
    - **Cỡ chữ Bảng biểu:** {int(Config.TABLE_FONT_SIZE)} pt (Căn giữa trang, Dòng đầu in đậm)
    - **Khoảng cách dòng:** {Config.LINE_SPACING} lines
    - **Căn lề trang:** Top {Config.MARGIN_TOP_CM}cm, Bottom {Config.MARGIN_BOTTOM_CM}cm, Left {Config.MARGIN_LEFT_CM}cm, Right {Config.MARGIN_RIGHT_CM}cm
    """)

    # Kiểm tra Microsoft Word hoặc LibreOffice trên máy
    word_path = TOCManager.find_word_executable()
    lo_path = TOCManager.find_libreoffice_executable()
    if word_path:
        st.success("✅ Phát hiện Microsoft Word (hỗ trợ tự động kết xuất Mục lục và tính toán số trang chính xác).")
    elif lo_path:
        st.success("✅ Phát hiện LibreOffice Headless (hỗ trợ cập nhật số trang trực tiếp).")
    else:
        st.caption("💡 Mẹo: Cài đặt Microsoft Word hoặc LibreOffice nếu muốn tự động kết xuất số trang offline.")

# MAIN CONTENT
st.markdown('<div class="main-header">📄 Chuẩn Hóa Báo Cáo Word Tự Động (.docx)</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Tự động nhận diện Trang bìa, Bảng biểu, Tiêu đề phân cấp (HSM + AI DeepSeek), Danh sách Bullet, Cập nhật Mục lục và Tích hợp Cổng QR Banking chuẩn Việt Nam.</div>', unsafe_allow_html=True)

# Tải file lên
uploaded_file = st.file_uploader(
    "Kéo và thả file Word (.docx) của bạn vào đây:",
    type=["docx"],
    help="Chỉ chấp nhận định dạng Word .docx tiêu chuẩn"
)

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    file_key = f"{uploaded_file.name}_{len(file_bytes)}"
    if st.session_state.get("current_file_key") != file_key:
        st.session_state["current_file_key"] = file_key
        st.session_state["processed_data"] = None
        st.session_state["payment_confirmed"] = False
        st.session_state["poll_count"] = 0
        st.session_state["balloon_shown"] = False

    file_details = {
        "Tên file": uploaded_file.name,
        "Kích thước": f"{round(len(file_bytes) / 1024, 1)} KB"
    }

    col1, col2 = st.columns([3, 1])
    with col1:
        st.info(f"📁 **File đã chọn:** `{file_details['Tên file']}` ({file_details['Kích thước']})")
    with col2:
        start_btn = st.button("🚀 Bắt đầu Chuẩn hóa", type="primary", use_container_width=True)

    # Tùy chọn chèn Mục lục tự động
    auto_insert_toc = st.checkbox(
        "📑 Tự động chèn Trang Mục lục nếu chưa có",
        value=True,
        help="Tự động phát hiện và chèn trang Mục lục chuẩn OpenXML trước Chương 1 (hoặc dọn dẹp các dòng chấm chấm gõ tay)."
    )

    if start_btn:
        progress_bar = st.progress(0)
        status_box = st.empty()
        log_records = []

        def log_callback(msg: str, progress: float):
            progress_bar.progress(int(progress * 100))
            status_box.markdown(f"**Đang thực hiện:** {msg}")
            log_records.append(f"[{time.strftime('%H:%M:%S')}] {msg}")

        # Khởi tạo pipeline với cấu hình từ giao diện
        current_api_key = api_key_input.strip() if api_key_input else None
        current_base_url = base_url_input.strip() if base_url_input else None
        current_model = model_input.strip() if model_input else None
        standardizer = DocumentStandardizer(
            api_key=current_api_key,
            base_url=current_base_url,
            model=current_model,
            normalize_chapter_to_roman=normalize_chapter_roman,
            normalize_heading_caps=normalize_heading_caps
        )

        try:
            with st.spinner("Hệ thống đang phân tích và chuẩn hóa văn bản..."):
                input_bytes = io.BytesIO(file_bytes)
                result = standardizer.process(
                    input_bytes,
                    output_target=None,
                    log_callback=log_callback,
                    auto_insert_toc=auto_insert_toc,
                    evaluate_toc=True,
                    normalize_chapter_to_roman=normalize_chapter_roman,
                    normalize_heading_caps=normalize_heading_caps
                )

            # Tạo mã đơn giao dịch ngắn gọn và ghi nhận thời điểm tạo
            order_code = f"BC{int(time.time() * 1000) % 900000 + 100000:06d}"
            order_time = time.strftime("%Y-%m-%d %H:%M:%S")
            st.session_state["processed_data"] = {
                "result": result,
                "order_code": order_code,
                "order_time": order_time,
                "log_records": log_records,
                "filename": uploaded_file.name
            }
            st.session_state["payment_confirmed"] = False
            st.session_state["poll_count"] = 0
            st.session_state["balloon_shown"] = False
            progress_bar.empty()
            status_box.empty()
        except Exception as e:
            st.error(f"❌ Đã xảy ra lỗi trong quá trình chuẩn hóa: {str(e)}")
            st.exception(e)

    # Hiển thị kết quả xử lý và Khung thanh toán từ Session State
    if st.session_state.get("processed_data") is not None:
        p_data = st.session_state["processed_data"]
        result = p_data["result"]
        order_code = p_data["order_code"]
        log_records = p_data["log_records"]
        filename = p_data["filename"]
        stats = result["stats"]
        out_buffer = result["buffer"]
        output_filename = f"ChuanHoa_{filename}"

        st.success("🎉 Chuẩn hóa văn bản thành công hoàn toàn!")

        if stats.get("chapters_roman_count", 0) > 0:
            st.info(f"🏛️ **Số Chương La Mã:** Đã tự động chuẩn hóa {stats['chapters_roman_count']} tiêu đề Chương sang số La Mã (CHƯƠNG I, II, III...) theo chuẩn học thuật.")
        if stats.get("headings_casing_normalized_count", 0) > 0:
            st.info(f"🔤 **Chuẩn hóa chữ hoa:** Đã tự động chuẩn hóa {stats['headings_casing_normalized_count']} tiêu đề cấp con từ ALL CAPS về chữ thường chuẩn thị giác.")
        if stats.get("toc_inserted"):
            st.info("📑 **Trang Mục lục:** Đã tự động tạo và chèn trang Mục lục chuẩn OpenXML vào vị trí chuẩn trước Chương 1.")
        if stats.get("toc_evaluated"):
            st.success("✨ **Kết xuất Mục lục:** Đã tự động tính toán toàn bộ số trang và đề mục chuẩn xác bằng Microsoft Word / LibreOffice.")

        # Hiển thị kết quả thống kê dạng Thẻ (Metrics Dashboard)
        st.markdown("### 📊 Thống kê Kết quả Xử lý")
        m_col1, m_col2, m_col3, m_col4 = st.columns(4)
        with m_col1:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-val">{stats['heading_1_count'] + stats['heading_2_count'] + stats['heading_3_count'] + stats.get('heading_4_count', 0) + stats['heading_unnumbered_count']}</div>
                <div class="metric-label">Tiêu đề (Headings)</div>
            </div>
            """, unsafe_allow_html=True)
        with m_col2:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-val">{stats['list_items_count']}</div>
                <div class="metric-label">Danh sách Bullet chuẩn hóa</div>
            </div>
            """, unsafe_allow_html=True)
        with m_col3:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-val">{stats['body_paragraphs_count']}</div>
                <div class="metric-label">Đoạn văn Body ({Config.FIRST_LINE_INDENT_CM}cm)</div>
            </div>
            """, unsafe_allow_html=True)
        with m_col4:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-val">{stats['tables_count']}</div>
                <div class="metric-label">Bảng biểu khóa lề</div>
            </div>
            """, unsafe_allow_html=True)

        m2_col1, m2_col2, m2_col3, m2_col4 = st.columns(4)
        with m2_col1:
            st.markdown(f"""
            <div class="metric-card" style="margin-top:10px;">
                <div class="metric-val">{stats['cover_paragraphs']}</div>
                <div class="metric-label">Đoạn Trang bìa (Được khóa)</div>
            </div>
            """, unsafe_allow_html=True)
        with m2_col2:
            st.markdown(f"""
            <div class="metric-card" style="margin-top:10px;">
                <div class="metric-val">{stats['caption_count']}</div>
                <div class="metric-label">Caption Hình/Bảng</div>
            </div>
            """, unsafe_allow_html=True)
        with m2_col3:
            st.markdown(f"""
            <div class="metric-card" style="margin-top:10px;">
                <div class="metric-val">{stats['ai_classified_count']}</div>
                <div class="metric-label">Đoạn AI/Heuristic xử lý</div>
            </div>
            """, unsafe_allow_html=True)
        with m2_col4:
            st.markdown(f"""
            <div class="metric-card" style="margin-top:10px;">
                <div class="metric-val">{stats['processing_time']}s</div>
                <div class="metric-label">Thời gian thực thi</div>
            </div>
            """, unsafe_allow_html=True)

        # CỔNG THANH TOÁN QR BANKING / XUẤT FILE
        if enable_payment:
            st.markdown("---")
            if st.session_state.get("payment_confirmed"):
                tx_details = st.session_state.get("payment_details") or {}
                tx_id = tx_details.get("id") or tx_details.get("reference_number") or order_code
                paid_amt = PaymentVerifier.parse_amount(tx_details.get("amount_in", payment_amount_input))
                if paid_amt <= 0:
                    paid_amt = payment_amount_input

                st.markdown(f"""
                <div style="background: linear-gradient(135deg, #ECFDF5 0%, #D1FAE5 100%); border: 2px solid #10B981; border-radius: 12px; padding: 22px; margin-top: 15px; margin-bottom: 20px; box-shadow: 0 4px 12px rgba(16, 185, 129, 0.15);">
                    <div style="font-size: 1.35rem; font-weight: 700; color: #065F46; margin-bottom: 6px;">
                        ✅ ĐÃ NHẬN THANH TOÁN {paid_amt:,.0f} VNĐ THÀNH CÔNG! (Mã giao dịch: {tx_id})
                    </div>
                    <div style="font-size: 0.95rem; color: #047857;">
                        Hệ thống đã tự động nhận diện giao dịch thành công cho đơn hàng <b>{order_code}</b>. Toàn bộ file báo cáo đã được mở khóa!
                    </div>
                </div>
                """, unsafe_allow_html=True)

                if not st.session_state.get("balloon_shown", False):
                    st.balloons()
                    st.session_state["balloon_shown"] = True

                st.download_button(
                    label=f"📥 Tải về Báo cáo Hoàn thiện ({output_filename})",
                    data=out_buffer.getvalue(),
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    use_container_width=True
                )
            else:
                # Nút tải file bị KHÓA / ẨN HOÀN TOÀN!
                st.markdown("""
                <div class="payment-box">
                    <div class="payment-header">💳 Cổng Thanh Toán QR Banking / Mở Khóa Tải Báo Cáo</div>
                    <div class="payment-sub">
                        Quét mã QR bằng mọi ứng dụng Ngân hàng (App Banking) để thanh toán.
                        Hệ thống tự động nhận diện giao dịch và lập tức mở khóa nút tải file.
                    </div>
                </div>
                """, unsafe_allow_html=True)

                pay_col1, pay_col2 = st.columns([1, 1.25])
                transfer_memo = order_code
                local_qr_path = Config.find_local_qr_image()
                has_vietqr = Config.is_vietqr_configured(bank_id_input, account_no_input)
                vietqr_url = Config.get_vietqr_url(
                    bank_id=bank_id_input,
                    account_no=account_no_input,
                    account_name=account_name_input,
                    amount=payment_amount_input,
                    add_info=transfer_memo
                )

                with pay_col1:
                    # Hiển thị mã QR theo 2 cơ chế (VietQR tự động & Ảnh QR cá nhân)
                    if has_vietqr and local_qr_path:
                        tab_vqr, tab_loc = st.tabs(["✨ Mã VietQR Tự Động", "🖼️ Mã QR Cá Nhân"])
                        with tab_vqr:
                            st.image(
                                vietqr_url,
                                caption=f"Mã VietQR tự động: {payment_amount_input:,.0f} VNĐ" if payment_amount_input > 0 else "Mã VietQR tự động",
                                use_container_width=True
                            )
                        with tab_loc:
                            try:
                                st.image(
                                    local_qr_path,
                                    caption=f"Mã QR cá nhân ({Path(local_qr_path).name})",
                                    use_container_width=True
                                )
                            except Exception as img_err:
                                st.error(f"⚠️ Không thể hiển thị ảnh QR cá nhân: {img_err}")
                    elif local_qr_path:
                        try:
                            st.image(
                                local_qr_path,
                                caption=f"Mã QR cá nhân ({Path(local_qr_path).name})",
                                use_container_width=True
                            )
                        except Exception as img_err:
                            st.error(f"⚠️ Không thể hiển thị ảnh QR cá nhân: {img_err}")
                    elif has_vietqr:
                        st.image(
                            vietqr_url,
                            caption=f"Quét bằng App Ngân hàng ({payment_amount_input:,.0f} VNĐ)" if payment_amount_input > 0 else "Quét bằng App Ngân hàng",
                            use_container_width=True
                        )
                    else:
                        st.warning("⚠️ Chưa cấu hình thông tin ngân hàng hoặc chưa có ảnh QR trong thư mục gốc. Vui lòng cấu hình ở cột bên trái.")

                with pay_col2:
                    bank_name = Config.get_bank_display_name(bank_id_input)
                    if bank_name:
                        st.markdown(f"**🏦 Ngân hàng:** `{bank_name}`")
                    if account_no_input:
                        st.markdown(f"**🔢 Số tài khoản:** `{account_no_input}`")
                    if account_name_input:
                        st.markdown(f"**👤 Chủ tài khoản:** `{account_name_input}`")
                    if not account_no_input and local_qr_path:
                        st.markdown("💡 *Đang sử dụng mã QR cá nhân trong thư mục gốc.*")
                    if payment_amount_input > 0:
                        st.markdown(f"**💵 Số tiền:** `{payment_amount_input:,.0f} VNĐ`")
                    else:
                        st.markdown("**💵 Số tiền:** `Tùy tâm ủng hộ`")
                    st.markdown("**📝 Nội dung chuyển khoản (bấm vào ô để sao chép):**")
                    st.code(transfer_memo, language=None)
                    st.caption("⚡ Quét mã bằng App Banking sẽ tự động điền toàn bộ thông tin trên.")

                    if payment_mode_input == "auto":
                        st.markdown(f"""
                        <div style="background-color:#FEF3C7; border:1px solid #F59E0B; border-radius:8px; padding:10px 14px; margin-top:8px; margin-bottom:12px; color:#92400E; font-size:0.95rem;">
                            ⏳ <b>Đang chờ thanh toán...</b> Hệ thống tự động kiểm tra tài khoản của bạn.
                        </div>
                        """, unsafe_allow_html=True)

                        if not sepay_token_input and not test_mode_input:
                            st.info("💡 **Gợi ý Thử nghiệm:** Chưa cấu hình SePay API Token. Bạn có thể bấm nút **Kiểm tra thanh toán ngay** hoặc **Mô phỏng thanh toán** bên dưới để trải nghiệm ngay quy trình mở khóa tải file.")

                        just_checked_in_run = False
                        order_time = p_data.get("order_time")

                        btn_col_a, btn_col_b = st.columns([1.2, 1])
                        with btn_col_a:
                            if st.button("🔄 Kiểm tra thanh toán ngay", use_container_width=True, type="primary"):
                                just_checked_in_run = True
                                is_mock_test = bool(test_mode_input or not sepay_token_input)
                                with st.spinner("Đang kết nối SePay kiểm tra biến động số dư..."):
                                    is_paid, tx_info, msg = check_sepay_payment(
                                        order_code=order_code,
                                        expected_amount=payment_amount_input,
                                        api_token=sepay_token_input,
                                        account_number=account_no_input,
                                        test_mode=is_mock_test,
                                        min_transaction_date=order_time
                                    )
                                    if is_paid:
                                        st.session_state["payment_confirmed"] = True
                                        st.session_state["payment_details"] = tx_info
                                        st.rerun()
                                    else:
                                        st.warning(f"⚠️ {msg}")

                        with btn_col_b:
                            if test_mode_input or not sepay_token_input:
                                if st.button("🧪 Mô phỏng thanh toán (Test)", use_container_width=True):
                                    is_paid, tx_info, msg = check_sepay_payment(
                                        order_code=order_code,
                                        expected_amount=payment_amount_input,
                                        test_mode=True
                                    )
                                    st.session_state["payment_confirmed"] = True
                                    st.session_state["payment_details"] = tx_info
                                    st.rerun()
                            else:
                                st.caption(f"⚡ SePay API Polling ({Config.PAYMENT_CHECK_INTERVAL_SECONDS}s)")

                        st.warning("🔒 **Nút tải file đang bị khóa.** Nút tải file sẽ tự động xuất hiện ngay khi tiền vào tài khoản.")

                        # Tự động kiểm tra (Auto-polling) khi có SePay Token
                        is_test_env = bool(os.environ.get("STREAMLIT_TESTING") or os.environ.get("PYTEST_CURRENT_TEST"))
                        if sepay_token_input and not is_test_env and not test_mode_input and not just_checked_in_run:
                            is_paid, tx_info, msg = check_sepay_payment(
                                order_code=order_code,
                                expected_amount=payment_amount_input,
                                api_token=sepay_token_input,
                                account_number=account_no_input,
                                test_mode=False,
                                timeout=5,
                                min_transaction_date=order_time
                            )
                            if is_paid:
                                st.session_state["payment_confirmed"] = True
                                st.session_state["payment_details"] = tx_info
                                st.rerun()
                            else:
                                cur_polls = st.session_state.get("poll_count", 0)
                                if cur_polls < 60:
                                    st.session_state["poll_count"] = cur_polls + 1
                                    st.caption(f"🔄 Hệ thống tự động làm mới sau {Config.PAYMENT_CHECK_INTERVAL_SECONDS}s... (Lần quét: {st.session_state['poll_count']}/60)")
                                    time.sleep(Config.PAYMENT_CHECK_INTERVAL_SECONDS)
                                    st.rerun()
                                else:
                                    st.caption("ℹ️ Đã tạm dừng tự động kiểm tra sau 3 phút. Bấm nút 'Kiểm tra thanh toán ngay' khi bạn chuyển khoản xong.")

                    else:
                        st.info("🖐️ **Chế độ Xác nhận Thủ công (Cấp độ 1):** Sau khi chuyển khoản, bấm nút bên dưới để mở khóa tải file.")
                        if st.button("✅ Tôi đã chuyển khoản xong", use_container_width=True, type="primary"):
                            st.session_state["payment_confirmed"] = True
                            st.session_state["payment_details"] = {
                                "id": f"MANUAL_{order_code}",
                                "amount_in": payment_amount_input,
                                "transaction_content": f"Chuyen khoan {order_code}"
                            }
                            st.rerun()
                        st.warning("🔒 **Nút tải file đang bị khóa.** Vui lòng bấm xác nhận sau khi chuyển khoản để tải tài liệu.")
        else:
            st.markdown("---")
            # Nếu tắt cổng thanh toán, hiển thị trực tiếp nút tải file
            st.download_button(
                label=f"📥 Tải về Báo cáo Hoàn thiện ({output_filename})",
                data=out_buffer.getvalue(),
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True
            )

        # Hướng dẫn mở file
        st.info("""
        📌 **Lưu ý khi mở tài liệu trong Word:**
        - Tài liệu đã được tích hợp cờ `<w:updateFields/>`. Khi bạn mở file, Microsoft Word có thể hiển thị hộp thoại hỏi có muốn cập nhật các trường liên kết không: hãy nhấn **Yes** để Mục lục (TOC) và số trang tự động cập nhật chuẩn xác 100%!
        """)

        # Nhật ký chi tiết
        with st.expander("📋 Xem chi tiết Nhật ký Xử lý (Processing Logs)"):
            st.code("\n".join(log_records), language="text")

else:
    # Hướng dẫn nhanh khi chưa tải file
    st.markdown("""
    ---
    ### 🌟 Các Tính Năng Nổi Bật của Hệ Thống:
    1. **Phân hệ 1 (Document Zoning):** Tự động phát hiện và bảo vệ Trang bìa không bị lệch layout; tách riêng toàn bộ Bảng biểu không áp đặt thụt lề chuẩn.
    2. **Phân hệ 2 (Waterfall Classifier):** Phân loại 5 tầng chính xác trong 1 vòng lặp O(N): Media/Công thức Toán, Chú thích hình/bảng (Caption), Tiêu đề (Heading), Danh sách (List), Văn bản thân (Body Text).
    3. **Phân hệ 3 (HSM & AI DeepSeek):** State Machine kiểm soát chặt chẽ phân cấp tăng tiến `[1] -> [1, 1]`, tránh bẫy đánh đồng danh sách số con làm Heading 1; tích hợp DeepSeek AI nhận diện tiêu đề khó trong 0.5s.
    4. **Phân hệ 4 (List Normalizer):** Xóa toàn bộ bullet rác, quy đổi chuẩn mực về dấu `- ` (cấp 1) và `+ ` (cấp 2).
    5. **Phân hệ 5 (Run-level Manipulation):** Tuyệt đối không xóa Run text làm mất formatting nội bộ; bảo toàn bold/italic cục bộ và công thức toán học; reset kế thừa Styles chuẩn.
    6. **Phân hệ 6 (Auto-TOC Update):** Kích hoạt thẻ `updateFields` tự động cập nhật Mục lục và số trang khi mở file; chống lệch trang chuẩn Nghị định 30.
    7. **Cổng Thanh toán QR Banking:** Tự động tạo mã VietQR chuẩn quốc tế với số tiền và nội dung chuyển khoản tự sinh, hỗ trợ mã QR cá nhân offline.
    """)
