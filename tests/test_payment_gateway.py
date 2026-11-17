"""
Unit tests cho tính năng Cổng Thanh toán QR Banking (VietQR & Local QR)
"""

import os
import unittest
from unittest.mock import patch, MagicMock
from pathlib import Path
from config import Config, get_config_value, POPULAR_BANKS, normalize_bank_id, parse_payment_amount

class TestPaymentGatewayConfig(unittest.TestCase):
    def setUp(self):
        # Lưu lại biến môi trường gốc
        self.original_env = os.environ.copy()

    def tearDown(self):
        # Khôi phục biến môi trường
        os.environ.clear()
        os.environ.update(self.original_env)
        Config.reload()

    def test_config_attributes_presence(self):
        """Kiểm tra sự tồn tại của các biến cấu hình thanh toán"""
        self.assertTrue(hasattr(Config, "ENABLE_PAYMENT_GATEWAY"))
        self.assertTrue(hasattr(Config, "BANK_ID"))
        self.assertTrue(hasattr(Config, "ACCOUNT_NO"))
        self.assertTrue(hasattr(Config, "ACCOUNT_NAME"))
        self.assertTrue(hasattr(Config, "PAYMENT_AMOUNT"))
        self.assertTrue(hasattr(Config, "PAYMENT_MODE"))
        self.assertTrue(hasattr(Config, "SEPAY_API_TOKEN"))
        self.assertTrue(hasattr(Config, "PAYMENT_CHECK_INTERVAL_SECONDS"))
        self.assertTrue(hasattr(Config, "SEPAY_TEST_MODE"))
        self.assertIsInstance(Config.ENABLE_PAYMENT_GATEWAY, bool)
        self.assertIsInstance(Config.PAYMENT_AMOUNT, int)
        self.assertIsInstance(Config.PAYMENT_CHECK_INTERVAL_SECONDS, int)
        self.assertIsInstance(Config.SEPAY_TEST_MODE, bool)

    def test_get_config_value_os_env(self):
        """Kiểm tra đọc cấu hình từ os.getenv"""
        os.environ["TEST_BANK_KEY"] = "VCB"
        val = get_config_value("TEST_BANK_KEY", "DEFAULT")
        self.assertEqual(val, "VCB")

        val_def = get_config_value("NON_EXISTENT_KEY_123", "FALLBACK")
        self.assertEqual(val_def, "FALLBACK")

    def test_get_config_value_st_secrets_priority(self):
        """Kiểm tra ưu tiên đọc từ st.secrets trước os.getenv khi deploy Streamlit Cloud"""
        os.environ["PAYMENT_SECRET_TEST"] = "FROM_ENV"
        mock_secrets = {"PAYMENT_SECRET_TEST": "FROM_SECRETS"}

        with patch.dict("sys.modules", {"streamlit": MagicMock(secrets=mock_secrets)}):
            val = get_config_value("PAYMENT_SECRET_TEST", "DEFAULT")
            self.assertEqual(val, "FROM_SECRETS")

    def test_vietqr_url_generation_standard(self):
        """Kiểm tra URL VietQR sinh ra đúng định dạng chuẩn API VietQR compact2"""
        url = Config.get_vietqr_url(
            bank_id="MB",
            account_no="0382838388",
            account_name="NGUYEN VAN A",
            amount=5000,
            add_info="CH BAO CAO BC123"
        )
        self.assertTrue(url.startswith("https://img.vietqr.io/image/MB-0382838388-compact2.png"))
        self.assertIn("amount=5000", url)
        self.assertIn("addInfo=CH%20BAO%20CAO%20BC123", url)
        self.assertIn("accountName=NGUYEN%20VAN%20A", url)

    def test_vietqr_url_special_characters_encoded(self):
        """Kiểm tra mã hóa ký tự tiếng Việt và ký tự đặc biệt trong URL"""
        url = Config.get_vietqr_url(
            bank_id="TCB",
            account_no="1900123456",
            account_name="TRẦN VĂN B",
            amount=10000,
            add_info="CH BAO CAO & REPORT #99"
        )
        self.assertIn("https://img.vietqr.io/image/TCB-1900123456-compact2.png", url)
        self.assertIn("amount=10000", url)
        # Ký tự & và # phải được url-encode
        self.assertNotIn(" & ", url)
        self.assertNotIn(" #", url)

    def test_is_vietqr_configured(self):
        """Kiểm tra hàm kiểm tra tính đầy đủ của thông tin VietQR"""
        self.assertTrue(Config.is_vietqr_configured("MB", "0123456789"))
        self.assertFalse(Config.is_vietqr_configured("", "0123456789"))
        self.assertFalse(Config.is_vietqr_configured("MB", ""))
        self.assertFalse(Config.is_vietqr_configured("", ""))

    def test_bank_display_name(self):
        """Kiểm tra ánh xạ tên ngân hàng thân thiện"""
        self.assertEqual(Config.get_bank_display_name("MB"), POPULAR_BANKS["MB"])
        self.assertEqual(Config.get_bank_display_name("VCB"), POPULAR_BANKS["VCB"])
        self.assertEqual(Config.get_bank_display_name("UNKNOWN_BANK"), "Ngân hàng UNKNOWN_BANK")

    def test_find_local_qr_image(self):
        """Kiểm tra tìm kiếm file ảnh QR có sẵn trong thư mục gốc (Cơ chế 2)"""
        test_qr = Path("qr.png")
        if test_qr.exists():
            test_qr.unlink()

        # Ban đầu không có file qr.png
        self.assertIsNone(Config.find_local_qr_image())

        # Tạo file qr.png giả lập
        test_qr.write_bytes(b"dummy_qr_data")
        try:
            found = Config.find_local_qr_image()
            self.assertIsNotNone(found)
            self.assertTrue(found.endswith("qr.png"))
        finally:
            if test_qr.exists():
                test_qr.unlink()

    def test_vietqr_url_account_no_with_spaces(self):
        """Kiểm tra số tài khoản có dấu cách được tự động xóa khoảng trắng"""
        url = Config.get_vietqr_url(
            bank_id="MB",
            account_no=" 038 283 8388 ",
            account_name="NGUYEN VAN A",
            amount=5000,
            add_info="CH BAO CAO"
        )
        self.assertIn("https://img.vietqr.io/image/MB-0382838388-compact2.png", url)

    def test_find_local_qr_image_jpg(self):
        """Kiểm tra nhận diện file qr.jpg"""
        test_qr = Path("qr.jpg")
        if test_qr.exists():
            test_qr.unlink()

        test_qr.write_bytes(b"dummy_jpg_data")
        try:
            found = Config.find_local_qr_image()
            self.assertIsNotNone(found)
            self.assertTrue(found.endswith("qr.jpg"))
        finally:
            if test_qr.exists():
                test_qr.unlink()

    def test_config_invalid_payment_amount(self):
        """Kiểm tra cấu hình PAYMENT_AMOUNT không hợp lệ sẽ tự fallback về 5000"""
        os.environ["PAYMENT_AMOUNT"] = "khong_phai_so"
        Config.reload(load_env=False)
        self.assertEqual(Config.PAYMENT_AMOUNT, 5000)

    def test_streamlit_app_loads(self):
        """Kiểm tra giao diện Streamlit app.py khởi chạy đầy đủ không có lỗi cú pháp hoặc crash"""
        from streamlit.testing.v1 import AppTest
        app_path = Path(__file__).parent.parent / "app.py"
        at = AppTest.from_file(str(app_path))
        at.run(timeout=10)
        self.assertFalse(at.exception, f"AppTest gặp lỗi: {at.exception}")
        # Kiểm tra tồn tại tiêu đề chính
        self.assertGreater(len(at.title), 0)

    def test_config_reload_with_new_env(self):
        """Kiểm tra Config.reload() cập nhật chính xác khi biến môi trường thay đổi"""
        os.environ["BANK_ID"] = "BIDV"
        os.environ["ACCOUNT_NO"] = "9876543210"
        os.environ["ACCOUNT_NAME"] = "LE THI C"
        os.environ["PAYMENT_AMOUNT"] = "20000"
        os.environ["ENABLE_PAYMENT_GATEWAY"] = "False"
        os.environ["PAYMENT_MODE"] = "manual"
        os.environ["SEPAY_API_TOKEN"] = "sepay_sample_tok_999"
        os.environ["PAYMENT_CHECK_INTERVAL_SECONDS"] = "5"
        os.environ["SEPAY_TEST_MODE"] = "True"

        Config.reload(load_env=False)
        self.assertEqual(Config.BANK_ID, "BIDV")
        self.assertEqual(Config.ACCOUNT_NO, "9876543210")
        self.assertEqual(Config.ACCOUNT_NAME, "LE THI C")
        self.assertEqual(Config.PAYMENT_AMOUNT, 20000)
        self.assertFalse(Config.ENABLE_PAYMENT_GATEWAY)
        self.assertEqual(Config.PAYMENT_MODE, "manual")
        self.assertEqual(Config.SEPAY_API_TOKEN, "sepay_sample_tok_999")
        self.assertEqual(Config.PAYMENT_CHECK_INTERVAL_SECONDS, 5)
        self.assertTrue(Config.SEPAY_TEST_MODE)

    def test_get_config_value_lowercase_secrets(self):
        """Kiểm tra đọc st.secrets với key dạng chữ thường (bank_id, enable_payment_gateway)"""
        mock_secrets = {"bank_id": "VCB", "enable_payment_gateway": "false"}
        with patch.dict("sys.modules", {"streamlit": MagicMock(secrets=mock_secrets)}):
            self.assertEqual(get_config_value("BANK_ID"), "VCB")
            self.assertEqual(get_config_value("ENABLE_PAYMENT_GATEWAY"), "false")

    def test_get_config_value_nested_secrets(self):
        """Kiểm tra đọc st.secrets trong section [payment] hoặc [vietqr]"""
        mock_secrets = {
            "payment": {
                "bank_id": "TCB",
                "account_no": "999888777",
                "payment_amount": "10000"
            }
        }
        with patch.dict("sys.modules", {"streamlit": MagicMock(secrets=mock_secrets)}):
            self.assertEqual(get_config_value("BANK_ID"), "TCB")
            self.assertEqual(get_config_value("ACCOUNT_NO"), "999888777")
            self.assertEqual(get_config_value("PAYMENT_AMOUNT"), "10000")

    def test_normalize_bank_id_aliases(self):
        """Kiểm tra hàm chuẩn hóa mã ngân hàng từ bí danh sang mã Napas chuẩn"""
        self.assertEqual(normalize_bank_id("AGRIBANK"), "VBA")
        self.assertEqual(normalize_bank_id("VIETCOMBANK"), "VCB")
        self.assertEqual(normalize_bank_id("MBBANK"), "MB")
        self.assertEqual(normalize_bank_id("techcombank"), "TCB")
        self.assertEqual(normalize_bank_id("vietinbank"), "ICB")
        self.assertEqual(normalize_bank_id("sacombank"), "STB")
        self.assertEqual(normalize_bank_id("vcb"), "VCB")
        self.assertEqual(normalize_bank_id("TIMO"), "TIMO")
        self.assertEqual(normalize_bank_id("CAKE"), "CAKE")
        self.assertEqual(normalize_bank_id(""), "")

    def test_parse_payment_amount_variations(self):
        """Kiểm tra chuyển đổi số tiền với nhiều định dạng khác nhau"""
        self.assertEqual(parse_payment_amount("10,000"), 10000)
        self.assertEqual(parse_payment_amount("10.000"), 10000)
        self.assertEqual(parse_payment_amount("5000.0"), 5000)
        self.assertEqual(parse_payment_amount(" 20000 "), 20000)
        self.assertEqual(parse_payment_amount(15000), 15000)
        self.assertEqual(parse_payment_amount("", default=5000), 5000)
        self.assertEqual(parse_payment_amount(None, default=5000), 5000)
        self.assertEqual(parse_payment_amount("abc", default=5000), 5000)

    def test_find_local_qr_image_ignores_zero_byte_file(self):
        """Kiểm tra hàm tìm ảnh QR cục bộ bỏ qua các file rỗng 0 bytes chống crash"""
        test_qr = Path("qr.png")
        if test_qr.exists():
            test_qr.unlink()

        test_qr.write_bytes(b"")  # 0 byte
        try:
            self.assertIsNone(Config.find_local_qr_image())
        finally:
            if test_qr.exists():
                test_qr.unlink()

    def test_is_vietqr_configured_handles_none_and_whitespace(self):
        """Kiểm tra tính an toàn tuyệt đối khi truyền None hoặc khoảng trắng"""
        # Khi truyền chuỗi rỗng hoặc khoảng trắng
        self.assertFalse(Config.is_vietqr_configured("", ""))
        self.assertFalse(Config.is_vietqr_configured("   ", "   "))
        self.assertFalse(Config.is_vietqr_configured("MB", "   "))
        self.assertFalse(Config.is_vietqr_configured("   ", "0123456789"))
        self.assertTrue(Config.is_vietqr_configured("MB", "0123456789"))
        # Khi không có cấu hình mặc định và truyền None
        Config.BANK_ID = ""
        Config.ACCOUNT_NO = ""
        self.assertFalse(Config.is_vietqr_configured(None, None))

    def test_get_bank_display_name_empty(self):
        """Kiểm tra hiển thị tên ngân hàng khi truyền chuỗi rỗng"""
        self.assertEqual(Config.get_bank_display_name(""), "Ngân hàng")
        self.assertEqual(Config.get_bank_display_name(None), Config.get_bank_display_name(Config.BANK_ID))

    def test_streamlit_app_custom_bank_preserved(self):
        """Kiểm tra UI Streamlit giữ nguyên mã ngân hàng tùy chỉnh không nằm trong danh sách phổ biến"""
        from streamlit.testing.v1 import AppTest
        os.environ["BANK_ID"] = "CUSTOMBANK"
        os.environ["ACCOUNT_NO"] = "999111"
        Config.reload(load_env=False)
        app_path = Path(__file__).parent.parent / "app.py"
        at = AppTest.from_file(str(app_path))
        at.run(timeout=10)
        self.assertFalse(at.exception)
        # Kiểm tra custom_bank text_input chứa CUSTOMBANK
        custom_inputs = [ti for ti in at.text_input if "Mã Ngân hàng khác" in ti.label]
        self.assertEqual(len(custom_inputs), 1)
        self.assertEqual(custom_inputs[0].value, "CUSTOMBANK")

    def test_streamlit_app_popular_bank_selected(self):
        """Kiểm tra UI Streamlit tự động chọn đúng ngân hàng trong selectbox nếu nằm trong danh mục phổ biến"""
        from streamlit.testing.v1 import AppTest
        os.environ["BANK_ID"] = "TIMO"
        os.environ["ACCOUNT_NO"] = "999111"
        Config.reload(load_env=False)
        app_path = Path(__file__).parent.parent / "app.py"
        at = AppTest.from_file(str(app_path))
        at.run(timeout=10)
        self.assertFalse(at.exception)
        self.assertEqual(at.selectbox[0].value, "TIMO")

if __name__ == "__main__":
    unittest.main()
