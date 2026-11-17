# HỆ THỐNG CHUẨN HÓA BÁO CÁO WORD TỰ ĐỘNG (.DOCX)
### Chuẩn Văn bản Nghị định 30/2020/NĐ-CP & Báo cáo Học thuật / Kỹ thuật (Tích hợp AI DeepSeek)

---

## 🚀 Giới Thiệu
Hệ thống chuẩn hóa tự động các tài liệu báo cáo, đồ án, luận văn dạng Microsoft Word (`.docx`). Hệ thống kết hợp giữa giải mã cấu trúc OpenXML ở mức độ thấp, phân loại Thác nước (Waterfall Classifier), Máy trạng thái phân cấp (HSM) và Trí tuệ nhân tạo DeepSeek AI để chuẩn hóa văn bản hoàn toàn tự động, chính xác và bảo tồn 100% dữ liệu gốc.

---

## 🛠️ 6 Phân Hệ Cốt Lõi & 4 Lưu Ý Kỹ Thuật

1. **Phân hệ 1 (Document Zoning):**
   - Khóa vùng **Trang bìa**: Tự động nhận diện ranh giới trang bìa qua Section Break (`<w:sectPr>`) hoặc từ khóa bắt đầu nội dung ("LỜI MỞ ĐẦU", "CHƯƠNG 1",...). Bảo vệ tuyệt đối bố cục trang bìa, không bị áp đặt thụt lề 1.27cm.
   - Tách riêng **Bảng biểu** (`document.tables`): Đảm bảo các ô trong bảng có `first_line_indent = 0`, căn giữa bảng, cỡ chữ 12pt gọn gàng.

2. **Phân hệ 2 (Waterfall Classifier):**
   - 1 vòng lặp $O(N)$ duyệt toàn bộ đoạn văn:
     - **Tầng 1 (Media/Math):** `<w:drawing>`, `<w:pict>`, `<m:oMathPara>`, `<m:oMath>` $\rightarrow$ Căn giữa, lề 0.
     - **Tầng 2 (Caption):** Nhận diện Regex chú thích Hình/Bảng/Sơ đồ $\rightarrow$ Căn giữa, in nghiêng, 11pt, không thụt lề.
     - **Tầng 3 (Heading):** Đề mục phân cấp $\rightarrow$ Chuyển cho HSM & Whitelist.
     - **Tầng 4 (List items):** Bullet symbols (`•`, `●`, `○`, `▪`, `❖`, `➢`, `-`, `+`, `a)`, `b)`) $\rightarrow$ Chuẩn hóa danh sách.
     - **Tầng 5 (Body Text):** Dọn dẹp tab thừa $\rightarrow$ Căn đều hai bên (Justify), thụt lề đầu dòng `1.27cm` (chuẩn 0.5 inch).

3. **Phân hệ 3 (Hierarchical State Machine - HSM & AI Fallback):**
   - **HSM:** Theo dõi đường dẫn phân cấp `current_path = [1, 2]`. Ngăn chặn triệt để "bẫy" danh sách con được đánh số (ví dụ: đang ở mục 2.1 gặp "1. Bước một:...") bị gán nhầm làm Heading 1.
   - **AI Fallback (DeepSeek):** Gom danh sách JSON các dòng ngắn, in đậm nghi vấn không có số mục gửi DeepSeek API (`deepseek-chat`) phân loại trong 0.5s. Tự động chuyển sang Heuristic nâng cao nếu mất kết nối mạng.

4. **Phân hệ 4 (List Normalizer):**
   - Xóa bỏ toàn bộ ký tự bullet rác và icon ẩn.
   - Chuẩn hóa:
     - Cấp 1: Tiền tố `"- "` (`first_line_indent = 1.27cm`, `left_indent = 0`).
     - Cấp 2: Tiền tố `"+ "` (`first_line_indent = 2.0cm`, `left_indent = 0`).
   - Gỡ bỏ thẻ `<w:numPr>` để tránh xung đột bullet gốc trong Word.

5. **Phân hệ 5 (Run-level Manipulation & Style Reset):**
   - **Tuyệt đối không ghi đè `p.text = ...`** $\rightarrow$ Sửa trực tiếp trên từng Run, bảo tồn 100% các từ in đậm (`bold`), in nghiêng (`italic`) và công thức toán học nội tuyến.
   - **Xử lý Run phân mảnh & `<w:sym>`:** Duyệt qua các Run đầu để xóa sạch tab và bullet ký hiệu Wingdings ẩn.
   - **Reset Style Inheritance:** Xóa các thuộc tính ghi đè trực tiếp trên Run (font name, size, color) để kế thừa chuẩn từ Style gốc, giữ nguyên in đậm/nghiêng cục bộ.

6. **Phân hệ 6 (Auto-TOC Update & Chống lệch số trang):**
   - Chèn cờ `<w:updateFields w:val="true"/>` vào `word/settings.xml` $\rightarrow$ Word tự động cập nhật Mục lục và số trang khi mở file.
   - Chuẩn hóa lề trang A4: Top 2cm, Bottom 2cm, Left 3cm, Right 2cm.
   - Hỗ trợ gọi LibreOffice headless (nếu có trên máy tính) để tính toán số trang.

7. **Xử lý Heading không đánh số (Whitelist):**
   - **Nhóm hiển thị trong Mục lục (TOC):** "LỜI MỞ ĐẦU", "KẾT LUẬN", "TÀI LIỆU THAM KHẢO" $\rightarrow$ Style `Heading 1 Unnumbered` có `outlineLevel = 0`.
   - **Nhóm không hiển thị trong Mục lục:** "LỜI CẢM ƠN", "MỤC LỤC", "DANH MỤC HÌNH ẢNH",... $\rightarrow$ Style `Non-TOC Heading` không có `outlineLevel`.

---

## ⚡ Hướng Dẫn Sử Dụng Nhanh

### Cách 1: Khởi động 1-Click bằng file `run.bat` (Khuyên dùng)
1. Nhấp đúp chuột vào file **`run.bat`**.
2. Hệ thống sẽ:
   - Tự động kiểm tra Python trên máy.
   - Tự động tạo và kích hoạt môi trường ảo (`venv`).
   - Tự động cài đặt các thư viện cần thiết (`requirements.txt`).
   - Tự động mở trình duyệt Web giao diện trực quan tại `http://localhost:8501`.

### Cách 2: Cấu Hình Cổng Thanh Toán & Tự Động Nhận Diện Số Dư (SePay Cấp Độ 2)

Hệ thống hỗ trợ 2 Cấp độ vận hành thanh toán chuyên nghiệp:
- **Cấp độ 2 (Mặc định - 100% Tự động):** Kết nối SePay API kiểm tra biến động số dư ngân hàng theo thời gian thực. Khi tiền nổ vào tài khoản ngân hàng, hệ thống tự động xác nhận và **mở khóa nút tải file** hoàn thiện, chống bùng tiền tuyệt đối!
- **Cấp độ 1 (Bấm xác nhận):** Khách quét mã VietQR và bấm nút "Tôi đã chuyển khoản xong" để mở khóa tải file.

#### ⚡ Hướng dẫn 3 bước cực kỳ đơn giản để tiền về thẳng tài khoản của bạn (2 phút):
1. **Bước 1 - Đăng ký tài khoản SePay miễn phí:**
   - Truy cập [my.sepay.vn](https://my.sepay.vn) và đăng ký tài khoản (chỉ mất 1 phút).
   - Thêm tài khoản ngân hàng nhận tiền của bạn (hỗ trợ MB, Vietcombank, Techcombank, ACB, TPBank, MSB, Timo, Cake,...).
2. **Bước 2 - Lấy SePay API Token:**
   - Tại trang quản trị SePay, vào menu **Cấu hình công ty** ➡️ **Truy cập API**.
   - Bấm nút **+ Thêm API** (đặt tên tùy ý, trạng thái Hoạt động) ➡️ Nhấn **Lưu** và **Sao chép API Token**.
3. **Bước 3 - Dán API Token vào hệ thống:**
   - Dán mã Token vào ô **SePay API Token** ở cột Sidebar bên trái trên giao diện Web (hoặc điền vào biến `SEPAY_API_TOKEN` trong file `.env`).
   - Thế là xong! Khách quét mã thanh toán 5.000 VNĐ, tiền về thẳng tài khoản của bạn và hệ thống tự động mở file cho khách tải về ngay tức khắc!

*(💡 Mẹo: Bạn có thể bật **🧪 Chế độ Thử nghiệm (Sandbox/Mock)** trên giao diện để kiểm tra toàn bộ luồng tải file tự động mà không cần nạp tiền thật).*

---

### Cách 3: Điền Thông Tin Ngân Hàng Nhận Tiền (VietQR & Ảnh QR)
Hệ thống hỗ trợ song song 2 cơ chế hiển thị mã QR:
- **Cơ chế 1: Tự động tạo mã VietQR chuẩn quốc tế** (kèm số tiền và nội dung chuyển khoản tự sinh ngắn gọn ví dụ: `BC12345` qua API `img.vietqr.io`).
- **Cơ chế 2: Hiển thị file ảnh QR cá nhân của bạn** (chỉ cần ném file `qr.png` hoặc `qr.jpg` vào thư mục gốc).

Bạn có thể cấu hình tài khoản ngân hàng theo 3 cách:
1. **Cách 1: Nhập trực tiếp trên Web UI (Sidebar bên trái):**
   - Mở Web, nhìn sang cột bên trái tại mục **💳 Cổng Thanh Toán QR Banking**.
   - Chọn ngân hàng từ danh sách thả xuống (hoặc gõ mã tùy ý), nhập Số tài khoản, Tên chủ tài khoản và Số tiền. Khung thanh toán bên phải sẽ đổi mã QR tức thì!
2. **Cách 2: Cấu hình lâu dài trong file `.env`:**
   - Mở file `.env` và cập nhật các dòng:
     ```env
     ENABLE_PAYMENT_GATEWAY=True
     BANK_ID=MB
     ACCOUNT_NO=0382838388
     ACCOUNT_NAME=NGUYEN VAN A
     PAYMENT_AMOUNT=5000
     PAYMENT_MODE=auto
     SEPAY_API_TOKEN=your_sepay_api_token_here
     PAYMENT_CHECK_INTERVAL_SECONDS=3
     SEPAY_TEST_MODE=False
     ```
3. **Cách 3: Sử dụng ảnh QR có sẵn:**
   - Đặt file ảnh QR tải từ App Ngân hàng của bạn vào thư mục gốc dự án với tên `qr.png` hoặc `qr.jpg`. Web sẽ tự nhận diện và cung cấp Tab hiển thị ảnh cá nhân của bạn.

#### 📋 Bảng Mã Ngân Hàng Napas / VietQR Phổ Biến:
| Mã (BANK_ID) | Tên Ngân Hàng | Mã (BANK_ID) | Tên Ngân Hàng |
| :--- | :--- | :--- | :--- |
| **MB** | MBBank - Ngân hàng Quân Đội | **VCB** | Vietcombank - Ngoại Thương |
| **TCB** | Techcombank - Kỹ Thương | **BIDV** | BIDV - Đầu tư & Phát triển |
| **ICB** | VietinBank - Công Thương | **VBA** | Agribank - Nông nghiệp & PTNT |
| **ACB** | ACB - Á Châu | **VPB** | VPBank - Việt Nam Thịnh Vượng |
| **TPB** | TPBank - Tiên Phong | **STB** | Sacombank - Sài Gòn Thương Tín |
| **VIB** | VIB - Quốc tế | **SHB** | SHB - Sài Gòn Hà Nội |
| **HDB** | HDBank - Phát triển TP.HCM | **LPB** | LPBank - Lộc Phát |
| **MSB** | MSB - Hàng Hải | **OCB** | OCB - Phương Đông |
| **SEAB** | SeABank - Đông Nam Á | **NAB** | Nam A Bank - Nam Á |
| **BAB** | Bac A Bank - Bắc Á | **BVB** | BVBank - Bản Việt |
| **PVCB** | PVcomBank - Đại Chúng | **NCB** | NCB - Quốc Dân |
| **CAKE** | CAKE by VPBank | **TIMO** | Timo by BVBank |

---

## 🌐 Hướng Dẫn Đưa Web Lên Cloud 24/7 Miễn Phí (Streamlit Community Cloud)

Bạn có thể đưa trang web lên mạng Internet để bất kỳ ai cũng có thể truy cập 24/7 từ điện thoại hoặc máy tính mà không tốn một đồng chi phí:

### Bước 1: Đẩy mã nguồn lên GitHub
1. Tạo một tài khoản trên [GitHub](https://github.com/) (miễn phí).
2. Tạo một Repository mới (ví dụ: `chuan-hoa-word`).
3. Đẩy toàn bộ mã nguồn của thư mục này lên GitHub (trừ thư mục `venv`).

### Bước 2: Triển khai trên Streamlit Cloud
1. Truy cập [share.streamlit.io](https://share.streamlit.io/) và đăng nhập bằng tài khoản GitHub của bạn.
2. Nhấn nút **"New app"**.
3. Chọn Repository vừa tạo, Branch `main`, và Main file path là `app.py`.
4. Trong phần **Advanced settings** $\rightarrow$ **Secrets**, dán cấu hình từ file `.streamlit/secrets.toml.example`:
   ```toml
   DEEPSEEK_API_KEY = "sk-..."
   ENABLE_PAYMENT_GATEWAY = "True"
   BANK_ID = "MB"
   ACCOUNT_NO = "0382838388"
   ACCOUNT_NAME = "NGUYEN VAN A"
   PAYMENT_AMOUNT = "5000"
   PAYMENT_MODE = "auto"
   SEPAY_API_TOKEN = "your_sepay_api_token"
   PAYMENT_CHECK_INTERVAL_SECONDS = "3"
   ```
5. Nhấn **"Deploy!"**. Trong 1-2 phút, ứng dụng của bạn sẽ có một đường link công khai dạng `https://chuan-hoa-word.streamlit.app` hoạt động 24/7!

---

## 🧪 Chạy Kiểm Thử (Unit & Integration Tests)
Để chạy toàn bộ kiểm thử chuyên sâu (Bao gồm kiểm thử đơn vị từng phân hệ, các trường hợp biên khó, cổng thanh toán QR, xác thực SePay tự động và kiểm thử tích hợp toàn diện):
```bash
python -m unittest discover tests
```

---

## 📂 Cấu Trúc Mã Nguồn
```
├── .env                     # File cấu hình API Key DeepSeek & SePay của người dùng
├── .env.example             # File cấu hình mẫu
├── requirements.txt         # Danh sách thư viện phụ thuộc
├── run.bat                  # File khởi động tự động trên Windows
├── config.py                # Cấu hình quy chuẩn lề, font, whitelists, thanh toán
├── app.py                   # Giao diện Web trực quan (Streamlit)
├── core/
│   ├── __init__.py
│   ├── document_zoner.py    # Phân hệ 1: Document Zoning (Bìa & Bảng)
│   ├── classifier.py        # Phân hệ 2: Waterfall Classifier 5 tầng
│   ├── hsm.py               # Phân hệ 3: HSM & AI Fallback DeepSeek
│   ├── list_normalizer.py   # Phân hệ 4: List Normalizer
│   ├── run_manipulator.py   # Phân hệ 5: Run Manipulation & Symbol Removal
│   ├── toc_manager.py       # Phân hệ 6: updateFields & Margins A4
│   ├── payment_verifier.py  # Cấp độ 2: Tự động xác thực biến động số dư SePay API
│   ├── numbering_parser.py  # Đọc cấu trúc numbering.xml
│   ├── style_manager.py     # Thiết lập styles.xml & rFonts
│   └── pipeline.py          # Master Pipeline kết nối toàn bộ hệ thống
└── tests/
    ├── test_zoner.py
    ├── test_run_manipulator.py
    ├── test_hsm.py
    ├── test_list_normalizer.py
    ├── test_style_and_toc.py
    ├── test_edge_cases.py
    ├── test_pipeline.py
    ├── test_payment_gateway.py
    ├── test_payment_verifier.py
    └── create_sample_docx.py
```
