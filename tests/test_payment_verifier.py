"""
Unit tests chuyên sâu cho module core/payment_verifier.py (Cấp độ 2 - SePay API Automation)
Kiểm thử toàn diện các trường hợp: Mock mode, API response parsing, so khớp mã đơn, xử lý lỗi mạng & HTTP.
"""

import os
import unittest
from unittest.mock import patch, MagicMock
import requests
from config import Config
from core.payment_verifier import PaymentVerifier, check_sepay_payment


class TestPaymentVerifier(unittest.TestCase):
    def setUp(self):
        self.original_env = os.environ.copy()

    def tearDown(self):
        os.environ.clear()
        os.environ.update(self.original_env)
        Config.reload(load_env=False)

    def test_normalize_code(self):
        """Kiểm tra chuẩn hóa mã đơn hàng"""
        self.assertEqual(PaymentVerifier.normalize_code("bc12345"), "BC12345")
        self.assertEqual(PaymentVerifier.normalize_code("  BC99999  "), "BC99999")
        self.assertEqual(PaymentVerifier.normalize_code(""), "")
        self.assertEqual(PaymentVerifier.normalize_code(None), "")

    def test_parse_amount(self):
        """Kiểm tra chuyển đổi số tiền đa dạng định dạng"""
        self.assertEqual(PaymentVerifier.parse_amount(5000), 5000.0)
        self.assertEqual(PaymentVerifier.parse_amount("5000"), 5000.0)
        self.assertEqual(PaymentVerifier.parse_amount("10,000.00"), 10000.0)
        self.assertEqual(PaymentVerifier.parse_amount(" 50000 "), 50000.0)
        self.assertEqual(PaymentVerifier.parse_amount("abc"), 0.0)
        self.assertEqual(PaymentVerifier.parse_amount(None), 0.0)

    def test_check_empty_order_code(self):
        """Kiểm tra truyền mã đơn rỗng"""
        is_paid, details, msg = check_sepay_payment("", 5000)
        self.assertFalse(is_paid)
        self.assertEqual(details, {})
        self.assertIn("không hợp lệ", msg)

        is_paid, details, msg = check_sepay_payment("   ", 5000)
        self.assertFalse(is_paid)
        self.assertIn("không hợp lệ", msg)

    def test_test_mode_mock_verification(self):
        """Kiểm tra chế độ Test/Sandbox (Mock verification) thành công ngay lập tức"""
        is_paid, details, msg = check_sepay_payment("BC12345", 5000, test_mode=True)
        self.assertTrue(is_paid)
        self.assertTrue(details.get("is_mock"))
        self.assertEqual(details.get("code"), "BC12345")
        self.assertEqual(float(details.get("amount_in")), 5000.0)
        self.assertIn("Mô phỏng thanh toán thành công", msg)

    def test_test_mode_from_config(self):
        """Kiểm tra chế độ Test khi cấu hình SEPAY_TEST_MODE=True trong Config"""
        Config.SEPAY_TEST_MODE = True
        is_paid, details, msg = check_sepay_payment("BC99999", 10000)
        self.assertTrue(is_paid)
        self.assertTrue(details.get("is_mock"))
        self.assertEqual(details.get("code"), "BC99999")

    def test_missing_api_token(self):
        """Kiểm tra khi chưa cấu hình SePay API Token và test_mode=False"""
        Config.SEPAY_API_TOKEN = ""
        Config.SEPAY_TEST_MODE = False
        is_paid, details, msg = check_sepay_payment("BC12345", 5000, api_token="", test_mode=False)
        self.assertFalse(is_paid)
        self.assertEqual(details, {})
        self.assertIn("Chưa cấu hình SePay API Token", msg)

    @patch("requests.get")
    def test_sepay_payment_success_v1(self, mock_get):
        """Kiểm tra thanh toán thành công chuẩn SePay API v1 (key 'transactions')"""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "status": 200,
            "error": None,
            "messages": {"success": True},
            "transactions": [
                {
                    "id": "1001",
                    "bank_brand_name": "MBBank",
                    "account_number": "0382838388",
                    "transaction_date": "2026-11-16 20:00:00",
                    "amount_out": "0.00",
                    "amount_in": "5000.00",
                    "accumulated": "1005000.00",
                    "transaction_content": "NGUYEN VAN A chuyen khoan CH BAO CAO BC12345",
                    "reference_number": "FT26320000001",
                    "code": None
                }
            ]
        }
        mock_get.return_value = mock_response

        is_paid, details, msg = check_sepay_payment("BC12345", 5000, api_token="test_token_xyz")
        self.assertTrue(is_paid)
        self.assertEqual(details["id"], "1001")
        self.assertEqual(details["amount_in"], "5000.00")
        self.assertIn("thành công", msg)
        self.assertIn("1001", msg)

        # Kiểm tra headers Authorization Bearer
        mock_get.assert_called_once()
        headers_sent = mock_get.call_args[1]["headers"]
        self.assertEqual(headers_sent["Authorization"], "Bearer test_token_xyz")

    @patch("requests.get")
    def test_sepay_payment_success_v2(self, mock_get):
        """Kiểm tra thanh toán thành công chuẩn SePay API v2 (key 'data')"""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "status": "success",
            "data": [
                {
                    "id": 2002,
                    "amount_in": 10000,
                    "transaction_content": "Thanh toan BC77889",
                    "reference_number": "REF2002",
                    "code": "BC77889"
                }
            ]
        }
        mock_get.return_value = mock_response

        is_paid, details, msg = check_sepay_payment("bc77889", 5000, api_token="Bearer test_token_2")
        self.assertTrue(is_paid)
        self.assertEqual(details["id"], 2002)
        self.assertIn("thành công", msg)

    @patch("requests.get")
    def test_sepay_payment_case_insensitive_matching(self, mock_get):
        """Kiểm tra so khớp không phân biệt hoa thường và so khớp cả trong trường code"""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "transactions": [
                {
                    "id": "3003",
                    "amount_in": "15000",
                    "transaction_content": "chuyen tien bc99999 noi dung",
                    "code": None
                }
            ]
        }
        mock_get.return_value = mock_response

        is_paid, details, msg = check_sepay_payment("BC99999", 5000, api_token="token_ok")
        self.assertTrue(is_paid)
        self.assertEqual(details["id"], "3003")

    @patch("requests.get")
    def test_sepay_payment_underpaid(self, mock_get):
        """Kiểm tra tìm thấy mã đơn nhưng số tiền chuyển vào ít hơn số tiền yêu cầu"""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "transactions": [
                {
                    "id": "4004",
                    "amount_in": "2000.00",
                    "transaction_content": "CH BAO CAO BC12345",
                    "code": "BC12345"
                }
            ]
        }
        mock_get.return_value = mock_response

        is_paid, details, msg = check_sepay_payment("BC12345", 5000, api_token="token_ok")
        self.assertFalse(is_paid)
        self.assertEqual(details["id"], "4004")
        self.assertIn("nhỏ hơn số tiền yêu cầu", msg)

    @patch("requests.get")
    def test_sepay_payment_not_found(self, mock_get):
        """Kiểm tra không tìm thấy giao dịch nào khớp với mã đơn hàng"""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "transactions": [
                {
                    "id": "5005",
                    "amount_in": "50000.00",
                    "transaction_content": "NGUYEN VAN B chuyen khoan tien an trua",
                    "code": None
                }
            ]
        }
        mock_get.return_value = mock_response

        is_paid, details, msg = check_sepay_payment("BC12345", 5000, api_token="token_ok")
        self.assertFalse(is_paid)
        self.assertEqual(details, {})
        self.assertIn("Chưa tìm thấy giao dịch", msg)

    @patch("requests.get")
    def test_sepay_http_401_unauthorized(self, mock_get):
        """Kiểm tra bắt lỗi HTTP 401 khi API Token sai hoặc hết hạn"""
        mock_response = MagicMock()
        mock_response.status_code = 401
        mock_get.return_value = mock_response

        is_paid, details, msg = check_sepay_payment("BC12345", 5000, api_token="invalid_token")
        self.assertFalse(is_paid)
        self.assertIn("401", msg)
        self.assertIn("API Token không chính xác", msg)

    @patch("requests.get")
    def test_sepay_http_403_forbidden(self, mock_get):
        """Kiểm tra bắt lỗi HTTP 403 khi không có quyền truy cập"""
        mock_response = MagicMock()
        mock_response.status_code = 403
        mock_get.return_value = mock_response

        is_paid, details, msg = check_sepay_payment("BC12345", 5000, api_token="token_no_permission")
        self.assertFalse(is_paid)
        self.assertIn("403", msg)

    @patch("requests.get")
    def test_sepay_http_429_rate_limit(self, mock_get):
        """Kiểm tra bắt lỗi HTTP 429 khi vượt quá tần suất gọi API (3 req/s)"""
        mock_response = MagicMock()
        mock_response.status_code = 429
        mock_get.return_value = mock_response

        is_paid, details, msg = check_sepay_payment("BC12345", 5000, api_token="token_ok")
        self.assertFalse(is_paid)
        self.assertIn("429", msg)
        self.assertIn("tần suất", msg)

    @patch("requests.get")
    def test_sepay_http_500_server_error(self, mock_get):
        """Kiểm tra bắt lỗi HTTP 500 từ máy chủ SePay"""
        mock_response = MagicMock()
        mock_response.status_code = 500
        mock_get.return_value = mock_response

        is_paid, details, msg = check_sepay_payment("BC12345", 5000, api_token="token_ok")
        self.assertFalse(is_paid)
        self.assertIn("500", msg)

    @patch("requests.get")
    def test_sepay_timeout_exception(self, mock_get):
        """Kiểm tra xử lý ngoại lệ Timeout kết nối SePay an toàn"""
        mock_get.side_effect = requests.exceptions.Timeout("Connection timed out")

        is_paid, details, msg = check_sepay_payment("BC12345", 5000, api_token="token_ok", timeout=5)
        self.assertFalse(is_paid)
        self.assertIn("Hết thời gian chờ kết nối", msg)

    @patch("requests.get")
    def test_sepay_connection_error_exception(self, mock_get):
        """Kiểm tra xử lý ngoại lệ rớt mạng (ConnectionError) an toàn"""
        mock_get.side_effect = requests.exceptions.ConnectionError("Network disconnected")

        is_paid, details, msg = check_sepay_payment("BC12345", 5000, api_token="token_ok")
        self.assertFalse(is_paid)
        self.assertIn("Không thể kết nối đến máy chủ SePay", msg)

    @patch("requests.get")
    def test_sepay_invalid_json_response(self, mock_get):
        """Kiểm tra xử lý khi máy chủ SePay trả về dữ liệu không phải JSON chuẩn"""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.side_effect = ValueError("Invalid JSON")
        mock_get.return_value = mock_response

        is_paid, details, msg = check_sepay_payment("BC12345", 5000, api_token="token_ok")
        self.assertFalse(is_paid)
        self.assertIn("Không thể giải mã phản hồi JSON", msg)

    @patch("requests.get")
    def test_multiple_transactions_priority_to_full_payment(self, mock_get):
        """Kiểm tra khi có nhiều giao dịch, ưu tiên chọn giao dịch đã trả đủ tiền"""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "transactions": [
                {
                    "id": "TX_1",
                    "amount_in": "1000.00",
                    "transaction_content": "CH BAO CAO BC12345",
                    "code": "BC12345"
                },
                {
                    "id": "TX_2",
                    "amount_in": "50000.00",
                    "transaction_content": "NGUYEN VAN B chuyen tien",
                    "code": None
                },
                {
                    "id": "TX_3",
                    "amount_in": "5000.00",
                    "transaction_content": "CH BAO CAO BC12345",
                    "code": "BC12345"
                }
            ]
        }
        mock_get.return_value = mock_response

        is_paid, details, msg = check_sepay_payment("BC12345", 5000, api_token="tok_123")
        self.assertTrue(is_paid)
        self.assertEqual(details["id"], "TX_3")
        self.assertIn("TX_3", msg)

    @patch("requests.get")
    def test_null_and_malformed_fields_in_transactions(self, mock_get):
        """Kiểm tra xử lý danh sách chứa phần tử bất thường (None, rỗng, không phải dict) mà không bị crash"""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "transactions": [
                None,
                "invalid_str_element",
                12345,
                {},
                {"id": "BAD_1", "transaction_content": None, "amount_in": None, "code": None},
                {"id": "GOOD_1", "transaction_content": "BC12345", "amount_in": "5000", "code": None}
            ]
        }
        mock_get.return_value = mock_response

        is_paid, details, msg = check_sepay_payment("BC12345", 5000, api_token="tok_123")
        self.assertTrue(is_paid)
        self.assertEqual(details["id"], "GOOD_1")

    @patch("requests.get")
    def test_compact_order_code_with_spaces_and_symbols(self, mock_get):
        """Kiểm tra so khớp mã đơn hàng có dấu cách hoặc gạch nối (BC-12345 với BC12345)"""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "transactions": [
                {
                    "id": "SPACED_1",
                    "amount_in": "5000",
                    "transaction_content": "MBVCB.1234.BC12345.CT",
                    "code": None
                }
            ]
        }
        mock_get.return_value = mock_response

        is_paid, details, msg = check_sepay_payment("BC-12345", 5000, api_token="tok_123")
        self.assertTrue(is_paid)
        self.assertEqual(details["id"], "SPACED_1")

    @patch("requests.get")
    def test_zero_and_negative_expected_amount(self, mock_get):
        """Kiểm tra số tiền yêu cầu là 0 hoặc số âm (tùy tâm ủng hộ)"""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "transactions": [
                {
                    "id": "ANY_AMT",
                    "amount_in": "1000",
                    "transaction_content": "BC12345",
                    "code": None
                }
            ]
        }
        mock_get.return_value = mock_response

        is_paid, details, msg = check_sepay_payment("BC12345", 0, api_token="tok_123")
        self.assertTrue(is_paid)

        is_paid, details, msg = check_sepay_payment("BC12345", -100, api_token="tok_123")
        self.assertTrue(is_paid)

    @patch("requests.get")
    def test_rejection_of_false_positive_substring_matches(self, mock_get):
        """Kiểm tra CHỐNG BẪY FALSE POSITIVE: BC12345 không được nhận nhầm BC123456, ABC12345, BC12345X"""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "transactions": [
                {"id": "FALSE_1", "amount_in": 5000, "transaction_content": "Chuyen tien BC123456"},
                {"id": "FALSE_2", "amount_in": 5000, "transaction_content": "Chuyen tien ABC12345"},
                {"id": "FALSE_3", "amount_in": 5000, "transaction_content": "Chuyen tien BC12345X"},
                {"id": "FALSE_4", "amount_in": 5000, "transaction_content": "Chuyen tien BC1234"},
            ]
        }
        mock_get.return_value = mock_response

        is_paid, details, msg = check_sepay_payment("BC12345", 5000, api_token="tok_test")
        self.assertFalse(is_paid)
        self.assertEqual(details, {})
        self.assertIn("Chưa tìm thấy giao dịch", msg)

    def test_vietnamese_dot_amount_parsing(self):
        """Kiểm tra xử lý số tiền định dạng dấu chấm nghìn Việt Nam (5.000, 50.000, 50.000.000)"""
        self.assertEqual(PaymentVerifier.parse_amount("5.000"), 5000.0)
        self.assertEqual(PaymentVerifier.parse_amount("50.000"), 50000.0)
        self.assertEqual(PaymentVerifier.parse_amount("50.000.000"), 50000000.0)
        self.assertEqual(PaymentVerifier.parse_amount("5,000.00"), 5000.0)
        self.assertEqual(PaymentVerifier.parse_amount("123.45"), 123.45)

    def test_is_code_matched_boundary_conditions(self):
        """Kiểm tra hàm so khớp boundary is_code_matched với đa dạng biến thể"""
        # Các trường hợp hợp lệ
        self.assertTrue(PaymentVerifier.is_code_matched("BC12345", "MBVCB.1234.BC12345.CT"))
        self.assertTrue(PaymentVerifier.is_code_matched("BC12345", "Chuyen khoan BC 12345 ngay"))
        self.assertTrue(PaymentVerifier.is_code_matched("BC12345", "Thanh toan BC-12345"))
        self.assertTrue(PaymentVerifier.is_code_matched("BC12345", "BC_12345 xong"))
        self.assertTrue(PaymentVerifier.is_code_matched("BC12345", "Noi dung", tx_code="BC12345"))

        # Các trường hợp không hợp lệ (sai mã, prefix/suffix dính liền)
        self.assertFalse(PaymentVerifier.is_code_matched("BC12345", "Chuyen khoan BC123456"))
        self.assertFalse(PaymentVerifier.is_code_matched("BC12345", "Chuyen khoan ABC12345"))
        self.assertFalse(PaymentVerifier.is_code_matched("BC12345", "Chuyen khoan BC1234"))
        self.assertFalse(PaymentVerifier.is_code_matched("BC12345", "Chuyen khoan BC12345X"))

    @patch("requests.get")
    def test_min_transaction_date_filters_old_transactions(self, mock_get):
        """Kiểm tra chống Replay Attack: Bỏ qua giao dịch cũ hơn min_transaction_date"""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "transactions": [
                {
                    "id": "OLD_TX",
                    "amount_in": "5000",
                    "transaction_date": "2026-11-10 10:00:00",
                    "transaction_content": "BC12345"
                }
            ]
        }
        mock_get.return_value = mock_response

        # Giao dịch tạo lúc 2026-11-16 12:00:00 -> Giao dịch cũ ngày 10/11 phải bị từ chối
        is_paid, details, msg = check_sepay_payment(
            "BC12345",
            5000,
            api_token="tok_test",
            min_transaction_date="2026-11-16 12:00:00"
        )
        self.assertFalse(is_paid)
        self.assertEqual(details, {})

    @patch("requests.get")
    def test_bearer_token_formatting(self, mock_get):
        """Kiểm tra không nhân đôi tiền tố Bearer khi token đã có sẵn Bearer"""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {"transactions": []}
        mock_get.return_value = mock_response

        check_sepay_payment("BC12345", 5000, api_token="Bearer my_secret_token")
        headers = mock_get.call_args[1]["headers"]
        self.assertEqual(headers["Authorization"], "Bearer my_secret_token")

    def test_streamlit_app_payment_mode_radio_presence(self):
        """Kiểm tra giao diện Streamlit app.py hiển thị mục chọn Cấp độ thanh toán (Auto / Manual)"""
        from streamlit.testing.v1 import AppTest
        from pathlib import Path
        app_path = Path(__file__).parent.parent / "app.py"
        at = AppTest.from_file(str(app_path))
        at.run(timeout=10)
        self.assertFalse(at.exception)
        # Kiểm tra tồn tại radio chọn Chế độ thanh toán
        payment_radios = [r for r in at.radio if "Chế độ Xác nhận Thanh toán" in r.label]
        self.assertEqual(len(payment_radios), 1)
        self.assertEqual(payment_radios[0].value, "auto")

    def test_streamlit_app_payment_unlock_end_to_end(self):
        """Kiểm thử luồng mở khóa nút tải file Streamlit end-to-end (Khóa nút tải -> Bấm kiểm tra -> Mở khóa)"""
        import io
        from streamlit.testing.v1 import AppTest
        from pathlib import Path
        from config import Config
        old_test_mode = Config.SEPAY_TEST_MODE
        Config.SEPAY_TEST_MODE = True
        try:
            app_path = Path(__file__).parent.parent / "app.py"
            at = AppTest.from_file(str(app_path))
            at.run(timeout=10)
            # Tải file mẫu lên widget
            at.file_uploader[0].upload("test.docx", b"PK\x03\x04test_docx_data").run(timeout=10)
            # Thiết lập session_state có kết quả xử lý
            at.session_state["processed_data"] = {
                "result": {
                    "buffer": io.BytesIO(b"dummy docx bytes"),
                    "stats": {
                        "chapters_roman_count": 0,
                        "headings_casing_normalized_count": 0,
                        "toc_inserted": False,
                        "toc_evaluated": False,
                        "heading_1_count": 1,
                        "heading_2_count": 1,
                        "heading_3_count": 0,
                        "heading_4_count": 0,
                        "heading_unnumbered_count": 0,
                        "list_items_count": 0,
                        "body_paragraphs_count": 5,
                        "tables_count": 0,
                        "cover_paragraphs": 0,
                        "caption_count": 0,
                        "ai_classified_count": 0,
                        "processing_time": 0.5,
                    }
                },
                "order_code": "BC998877",
                "order_time": "2026-11-16 22:00:00",
                "log_records": ["Done"],
                "filename": "test.docx"
            }
            at.session_state["payment_confirmed"] = False
            at.run(timeout=10)

            # Ban đầu nút tải file bị KHÓA HOÀN TOÀN
            self.assertEqual(len(at.download_button), 0)

            # Tìm và bấm nút 'Kiểm tra thanh toán ngay'
            check_btns = [b for b in at.button if "Kiểm tra thanh toán ngay" in b.label]
            self.assertEqual(len(check_btns), 1)

            check_btns[0].click().run(timeout=10)
            self.assertFalse(at.exception)

            # Sau khi xác nhận thanh toán, payment_confirmed = True và BUNG NÚT TẢI FILE
            self.assertTrue(at.session_state["payment_confirmed"])
            self.assertEqual(len(at.download_button), 1)
            self.assertIn("ChuanHoa_test.docx", at.download_button[0].label)
        finally:
            Config.SEPAY_TEST_MODE = old_test_mode


if __name__ == "__main__":
    unittest.main()
