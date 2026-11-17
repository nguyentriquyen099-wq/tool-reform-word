"""
Module Kiểm tra Thanh toán Tự động qua Cổng SePay API (Cấp độ 2)
Tự động truy vấn biến động số dư ngân hàng theo thời gian thực (Webhook/Polling)
So khớp mã đơn hàng và số tiền thực nhận để mở khóa tải file hoàn thiện.
"""

import re
import time
import logging
from typing import Tuple, Dict, Any, Optional
import requests
from config import Config

logger = logging.getLogger(__name__)

DEFAULT_SEPAY_API_URL = "https://my.sepay.vn/userapi/transactions/list"


class PaymentVerifier:
    """
    Trình xác thực giao dịch chuyển khoản ngân hàng qua API SePay.
    Hỗ trợ chế độ Thử nghiệm (Sandbox/Mock) và Cấp độ 2 (Tự động nhận diện biến động số dư).
    """

    @staticmethod
    def normalize_code(raw_code: str) -> str:
        """Chuẩn hóa mã đơn hàng để so khớp (viết hoa, loại bỏ khoảng trắng thừa)"""
        if not raw_code:
            return ""
        return str(raw_code).strip().upper()

    @staticmethod
    def parse_amount(raw_amount: Any) -> float:
        """Chuyển đổi số tiền an toàn từ chuỗi/số sang float, hỗ trợ cả định dạng quốc tế và dấu chấm nghìn Việt Nam"""
        if raw_amount is None:
            return 0.0
        if isinstance(raw_amount, (int, float)):
            return max(0.0, float(raw_amount))
        s = str(raw_amount).strip().replace(" ", "")
        if not s:
            return 0.0
        # Cả dấu phẩy và dấu chấm: 5,000.00 hoặc 5.000,00
        if "," in s and "." in s:
            last_comma = s.rfind(",")
            last_dot = s.rfind(".")
            if last_comma < last_dot:  # 5,000.00
                s = s.replace(",", "")
            else:  # 5.000,00
                s = s.replace(".", "").replace(",", ".")
        elif "," in s:  # Chỉ có dấu phẩy: 5,000 hoặc 5,00
            parts = s.split(",")
            if len(parts[-1]) == 3:  # Dấu phẩy ngăn cách hàng nghìn: 5,000
                s = s.replace(",", "")
            else:
                s = s.replace(",", ".")
        elif "." in s:  # Chỉ có dấu chấm: 50.000.000 hoặc 5.000 hoặc 5000.00
            parts = s.split(".")
            if len(parts) > 2:  # 50.000.000
                s = s.replace(".", "")
            elif len(parts) == 2:
                # Nếu phần sau dấu chấm có 3 chữ số (vd: 5.000, 50.000), đây là phân cách nghìn kiểu VN
                if len(parts[1]) == 3:
                    s = s.replace(".", "")
                # Nếu có 2 chữ số thập phân (vd: 5000.00), giữ nguyên float chuẩn
        try:
            return max(0.0, float(s))
        except (ValueError, TypeError):
            return 0.0

    @classmethod
    def is_code_matched(cls, order_code: str, content: str, tx_code: Optional[str] = None) -> bool:
        """
        So khớp mã đơn hàng trong nội dung giao dịch hoặc trường code của SePay.
        Sử dụng boundary matching để chống bẫy false positive (ví dụ: BC12345 không được khớp BC123456 hay ABC12345).
        Hỗ trợ các biến thể người dùng gõ: BC12345, BC 12345, BC-12345, BC_12345, MBVCB.1234.BC12345.CT.
        """
        code_clean = cls.normalize_code(order_code)
        if not code_clean:
            return False

        # Kiểm tra trường code SePay tự bóc tách nếu có
        if tx_code:
            clean_tx_code = cls.normalize_code(tx_code)
            if clean_tx_code == code_clean:
                return True

        # Tách phần chữ và phần số nếu mã có dạng Chữ + Số (ví dụ: BC12345, BC-12345)
        m = re.match(r"^([A-Z]+)[\s\-_.]*([0-9]+)$", code_clean)
        if m:
            letters, digits = m.group(1), m.group(2)
            pattern = rf"(?<![A-Za-z0-9]){re.escape(letters)}[\s\-_.]*{re.escape(digits)}(?![A-Za-z0-9])"
        else:
            pattern = rf"(?<![A-Za-z0-9]){re.escape(code_clean)}(?![A-Za-z0-9])"

        content_str = str(content or "")
        if re.search(pattern, content_str, re.IGNORECASE):
            return True

        if tx_code and re.search(pattern, str(tx_code), re.IGNORECASE):
            return True

        return False

    @classmethod
    def check_sepay_payment(
        cls,
        order_code: str,
        expected_amount: int | float = 5000,
        api_token: Optional[str] = None,
        account_number: Optional[str] = None,
        test_mode: Optional[bool] = None,
        limit: int = 50,
        timeout: int = 10,
        base_url: Optional[str] = None,
        min_transaction_date: Optional[str] = None,
    ) -> Tuple[bool, Dict[str, Any], str]:
        """
        Kiểm tra biến động số dư ngân hàng từ SePay API.

        Args:
            order_code: Mã đơn hàng cần kiểm tra (ví dụ: 'BC12345').
            expected_amount: Số tiền tối thiểu yêu cầu (mặc định: 5000 VNĐ).
            api_token: Token SePay API. Nếu None sẽ đọc từ Config.SEPAY_API_TOKEN.
            account_number: Số tài khoản ngân hàng cần lọc (tùy chọn).
            test_mode: Bật/Tắt chế độ Sandbox mô phỏng. Nếu None sẽ đọc từ Config.SEPAY_TEST_MODE.
            limit: Số lượng giao dịch gần nhất cần quét (10 - 100, mặc định 50).
            timeout: Thời gian chờ tối đa khi gọi API (giây, mặc định 10s).
            base_url: Endpoint SePay API (mặc định lấy từ Config.SEPAY_API_URL hoặc DEFAULT_SEPAY_API_URL).
            min_transaction_date: Thời gian sớm nhất chấp nhận giao dịch (YYYY-MM-DD HH:MM:SS), chống khớp nhầm giao dịch cũ.

        Returns:
            Tuple[bool, Dict[str, Any], str]:
                - is_paid (bool): True nếu đã nhận đủ tiền và đúng nội dung.
                - transaction_details (dict): Thông tin chi tiết giao dịch khớp (hoặc rỗng).
                - message (str): Thông báo chi tiết trạng thái để hiển thị cho người dùng.
        """
        code_clean = cls.normalize_code(order_code)
        if not code_clean:
            return False, {}, "Mã đơn hàng không hợp lệ hoặc để trống."

        exp_amt = max(0.0, cls.parse_amount(expected_amount))

        # Xác định chế độ Test / Sandbox
        is_test = test_mode if test_mode is not None else getattr(Config, "SEPAY_TEST_MODE", False)

        # Lấy token từ tham số hoặc từ Config
        token = (api_token if api_token is not None else getattr(Config, "SEPAY_API_TOKEN", "")).strip()

        # Xử lý chế độ Test / Mock verification
        if is_test:
            mock_id = f"MOCK_{code_clean}_{int(time.time())}"
            mock_tx = {
                "id": mock_id,
                "bank_brand_name": getattr(Config, "BANK_ID", "MB"),
                "account_number": account_number or getattr(Config, "ACCOUNT_NO", "0382838388"),
                "transaction_date": time.strftime("%Y-%m-%d %H:%M:%S"),
                "amount_out": "0.00",
                "amount_in": f"{exp_amt:.2f}",
                "accumulated": "1000000.00",
                "transaction_content": f"Thanh toan test don hang {code_clean}",
                "reference_number": f"REF_{mock_id}",
                "code": code_clean,
                "is_mock": True,
            }
            return (
                True,
                mock_tx,
                f"Mô phỏng thanh toán thành công (Chế độ Thử nghiệm / Sandbox) cho đơn hàng {code_clean}."
            )

        # Nếu không ở chế độ Test mà chưa có token
        if not token:
            return (
                False,
                {},
                "Chưa cấu hình SePay API Token. Vui lòng cấu hình token trong file .env hoặc trên giao diện để tự động kiểm tra số dư."
            )

        # Xác định URL SePay API
        api_url = base_url or getattr(Config, "SEPAY_API_URL", DEFAULT_SEPAY_API_URL)

        # Chuẩn bị Headers & Query Parameters
        auth_header = token if token.lower().startswith("bearer ") else f"Bearer {token}"
        headers = {
            "Authorization": auth_header,
            "Content-Type": "application/json",
            "User-Agent": "Report-Standardizer-PaymentVerifier/2.0",
        }

        query_limit = min(max(int(limit), 10), 100)
        params: Dict[str, Any] = {"limit": query_limit}

        if min_transaction_date:
            params["transaction_date_min"] = str(min_transaction_date).strip()

        acc_num = account_number if account_number is not None else getattr(Config, "ACCOUNT_NO", "")
        if acc_num:
            clean_acc = str(acc_num).strip().replace(" ", "")
            if clean_acc:
                params["account_number"] = clean_acc

        # Gọi SePay API
        try:
            response = requests.get(api_url, headers=headers, params=params, timeout=timeout)
        except requests.exceptions.Timeout:
            logger.warning("Gọi SePay API bị timeout sau %ds", timeout)
            return False, {}, f"Hết thời gian chờ kết nối máy chủ SePay ({timeout}s). Vui lòng thử lại sau giây lát."
        except requests.exceptions.ConnectionError:
            logger.warning("Không thể kết nối đến máy chủ SePay: Lỗi mạng")
            return False, {}, "Không thể kết nối đến máy chủ SePay. Vui lòng kiểm tra đường truyền Internet."
        except requests.exceptions.RequestException as e:
            logger.error("Lỗi yêu cầu HTTP SePay: %s", str(e))
            return False, {}, f"Lỗi kết nối API SePay: {str(e)}"
        except Exception as e:
            logger.error("Lỗi không xác định khi kết nối SePay: %s", str(e))
            return False, {}, f"Lỗi không xác định khi kết nối SePay: {str(e)}"

        # Xử lý các mã lỗi HTTP từ SePay
        if response.status_code == 401:
            return False, {}, "Lỗi xác thực SePay (HTTP 401): API Token không chính xác hoặc đã hết hạn."
        elif response.status_code == 403:
            return False, {}, "Lỗi quyền truy cập SePay (HTTP 403): Token không có quyền truy vấn danh sách giao dịch."
        elif response.status_code == 429:
            return False, {}, "SePay API bị giới hạn tần suất gọi (HTTP 429 - Tối đa 3 req/s). Vui lòng chờ vài giây."
        elif response.status_code >= 500:
            return False, {}, f"Máy chủ SePay đang bảo trì hoặc gặp sự cố (HTTP {response.status_code})."
        elif response.status_code != 200:
            return False, {}, f"Máy chủ SePay phản hồi mã lỗi HTTP {response.status_code}."

        # Parse JSON
        try:
            payload = response.json()
        except Exception as e:
            return False, {}, f"Không thể giải mã phản hồi JSON từ SePay: {str(e)}"

        # Lấy danh sách giao dịch (hỗ trợ cả chuẩn v1 và v2)
        transactions = []
        if isinstance(payload, dict):
            if "transactions" in payload and isinstance(payload["transactions"], list):
                transactions = payload["transactions"]
            elif "data" in payload and isinstance(payload["data"], list):
                transactions = payload["data"]
            elif "data" in payload and isinstance(payload["data"], dict) and "transactions" in payload["data"]:
                transactions = payload["data"]["transactions"]
        elif isinstance(payload, list):
            transactions = payload

        if not isinstance(transactions, list):
            return False, {}, "Dữ liệu trả về từ SePay không đúng định dạng danh sách giao dịch."

        # So khớp giao dịch
        partial_match_tx = None
        partial_match_amt = 0.0

        for tx in transactions:
            if not isinstance(tx, dict):
                continue

            # Bỏ qua giao dịch cũ hơn mốc thời gian tối thiểu (nếu có cấu hình)
            if min_transaction_date:
                tx_date = str(tx.get("transaction_date") or "").strip()
                if tx_date and tx_date < str(min_transaction_date).strip():
                    continue

            content = str(tx.get("transaction_content") or "")
            tx_code = str(tx.get("code") or "")

            # Kiểm tra xem mã đơn hàng có xuất hiện trong nội dung hoặc trường code không (strict boundary matching)
            if cls.is_code_matched(code_clean, content, tx_code):
                amt_in = cls.parse_amount(tx.get("amount_in", 0))
                if amt_in >= exp_amt:
                    tx_id = tx.get("id") or tx.get("reference_number") or "N/A"
                    return (
                        True,
                        tx,
                        f"Đã nhận thanh toán thành công {amt_in:,.0f} VNĐ cho đơn hàng {code_clean} (Mã GD: {tx_id})."
                    )
                else:
                    if amt_in > partial_match_amt:
                        partial_match_amt = amt_in
                        partial_match_tx = tx

        # Nếu tìm thấy giao dịch nhưng số tiền chưa đủ
        if partial_match_tx is not None:
            tx_id = partial_match_tx.get("id") or partial_match_tx.get("reference_number") or "N/A"
            return (
                False,
                partial_match_tx,
                f"Đã phát hiện giao dịch mã {code_clean} (Mã GD: {tx_id}) nhưng số tiền nhận ({partial_match_amt:,.0f} VNĐ) nhỏ hơn số tiền yêu cầu ({exp_amt:,.0f} VNĐ)."
            )

        return (
            False,
            {},
            f"Chưa tìm thấy giao dịch thanh toán cho mã đơn {code_clean} trên tài khoản ngân hàng."
        )


def check_sepay_payment(
    order_code: str,
    expected_amount: int | float = 5000,
    api_token: Optional[str] = None,
    account_number: Optional[str] = None,
    test_mode: Optional[bool] = None,
    limit: int = 50,
    timeout: int = 10,
    base_url: Optional[str] = None,
    min_transaction_date: Optional[str] = None,
) -> Tuple[bool, Dict[str, Any], str]:
    """Hàm helper tiện lợi gọi PaymentVerifier.check_sepay_payment"""
    return PaymentVerifier.check_sepay_payment(
        order_code=order_code,
        expected_amount=expected_amount,
        api_token=api_token,
        account_number=account_number,
        test_mode=test_mode,
        limit=limit,
        timeout=timeout,
        base_url=base_url,
        min_transaction_date=min_transaction_date,
    )
