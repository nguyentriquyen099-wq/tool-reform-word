"""
Phân hệ 3: Hierarchical State Machine (HSM) & AI Fallback
- State Machine kiểm tra tính tăng tiến phân cấp (Current_Path = [1, 2],
  tránh bẫy numbered list con bị gán nhầm làm Heading 1).
- AI Fallback: gom danh sách JSON các dòng ngắn/in đậm nghi vấn không có số mục,
  gửi DeepSeek API (OpenAI-compatible) để phân loại nhanh và chính xác.
"""

import re
import json
import logging
import unicodedata
from typing import List, Dict, Optional, Tuple, Any
import requests
from config import Config

logger = logging.getLogger(__name__)

class HierarchicalStateMachine:
    def __init__(self):
        """
        Khởi tạo State Machine lưu trữ đường dẫn phân cấp hiện tại.
        Ví dụ: [1] -> [1, 1] -> [1, 2] -> [1, 2, 1] -> [2]
        """
        self.current_path: List[int] = []
        self._in_numbered_list: bool = False
        self._list_item_number: int = 0

    def reset(self):
        self.current_path = []
        self._in_numbered_list = False
        self._list_item_number = 0

    def validate_and_update(self, candidate_path: List[int], text: str, is_bold: bool = False) -> Tuple[bool, int]:
        """
        Kiểm tra tính hợp lệ của candidate_path đối với current_path:
        - candidate_path: danh sách các số, ví dụ "1.2.3" -> [1, 2, 3]
        - text: nội dung đề mục để kiểm tra dấu câu / độ dài
        - is_bold: có in đậm không

        Trả về: (is_valid_heading: bool, heading_level: int)
        """
        if not candidate_path:
            return False, 0

        cand_len = len(candidate_path)
        stripped_text = text.strip()
        is_chapter_kw = bool(re.match(r'^(chương|phần|bài)\s+', stripped_text, re.IGNORECASE))

        # Quy tắc 1: Đề mục không quá dài (< 160 ký tự)
        if len(stripped_text) > 160:
            return False, 0

        # Quy tắc 2: Đề mục tiếng Việt chuẩn không kết thúc bằng dấu hai chấm (:), chấm phẩy (;)
        if (stripped_text.endswith(":") or stripped_text.endswith(";")) and not is_chapter_kw:
            return False, 0

        # Nếu đoạn văn có nhiều hơn 2 câu hoặc dài với dấu chấm câu (câu phức), thường là numbered list
        sentences = [s for s in re.split(r'[\.\?\!]\s+', stripped_text) if len(s.strip()) > 8]
        if len(sentences) >= 2 and not is_bold and not is_chapter_kw:
            return False, 0

        # Quy tắc 3: Bẫy Numbered List Con (Cực kỳ quan trọng) & Heading 1
        if cand_len == 1:
            cand_num = candidate_path[0]

            # Kiểm tra nếu là một phần tử trong danh sách liệt kê số (1., 2., 3...)
            if self._in_numbered_list and not is_chapter_kw:
                # Nếu tiếp tục dãy số của list (ví dụ: đang ở bước 1 mà gặp bước 2, 3...)
                if cand_num == self._list_item_number + 1:
                    # Nếu không phải tiêu đề in đậm viết hoa hoàn toàn -> Chắc chắn là item tiếp theo của list!
                    if not (is_bold and stripped_text.isupper()):
                        self._list_item_number = cand_num
                        return False, 0

            # Đối với đề mục cấp 1 chỉ có 1 số ("1.", "2.") không có chữ "Chương/Phần":
            # Bắt buộc phải là in đậm (is_bold) hoặc viết hoa toàn bộ (isupper),
            # không thể là một câu văn thường không in đậm!
            if not is_chapter_kw and not is_bold and not stripped_text.isupper():
                self._in_numbered_list = True
                self._list_item_number = cand_num
                return False, 0

            if self.current_path:
                curr_major = self.current_path[0]

                # Nếu cand_num < curr_major (ví dụ đang ở Chương 2 mà gặp "1. Làm gì đó"):
                # Đây là danh sách số 1, 2, 3 bên trong Chương 2, KHÔNG PHẢI Heading 1!
                if cand_num < curr_major and not is_chapter_kw:
                    self._in_numbered_list = True
                    self._list_item_number = cand_num
                    return False, 0

                # Nếu cand_num == curr_major:
                # Trong cùng 1 tài liệu, không thể có 2 Heading 1 có cùng số mục trừ khi là "Chương X"
                if cand_num == curr_major and not is_chapter_kw:
                    self._in_numbered_list = True
                    self._list_item_number = cand_num
                    return False, 0

                # Nếu đang ở cấp sâu (ví dụ [1, 2] hoặc [1, 1, 1]):
                # Nhảy lên cand_num == curr_major + 1 mà không có từ khóa "Chương":
                # Phải là in đậm và ngắn (< 80 ký tự), không phải câu hành động thường!
                if len(self.current_path) > 1 and not is_chapter_kw:
                    if not is_bold or len(stripped_text) > 80:
                        self._in_numbered_list = True
                        self._list_item_number = cand_num
                        return False, 0

                # Nếu cand_num > curr_major + 1 (ví dụ từ chương 1 nhảy cóc lên chương 5):
                if cand_num > curr_major + 1:
                    return False, 0

            # Hợp lệ là Heading 1
            self.current_path = [cand_num]
            self._in_numbered_list = False
            self._list_item_number = 0
            return True, 1

        # Quy tắc 4: Kiểm tra cấp con (cand_len > 1)
        # candidate_path phải khớp tiền tố với current_path
        major_cand = candidate_path[0]
        if self.current_path:
            curr_major = self.current_path[0]
            if major_cand != curr_major and major_cand != curr_major + 1:
                # Tiền tố không khớp với phân cấp hiện tại
                return False, 0

        # Kiểm tra bước nhảy cấp (Level Jump): không được nhảy cóc quá 1 cấp
        curr_len = len(self.current_path) if self.current_path else 0
        if cand_len > curr_len + 1:
            if not self.current_path and cand_len == 2 and candidate_path[0] == 1:
                pass # Hợp lệ: tài liệu bắt đầu ngay từ mục 1.1 (implicit Chương 1)
            else:
                return False, 0

        # Cập nhật state path
        self.current_path = list(candidate_path)
        self._in_numbered_list = False
        self._list_item_number = 0
        heading_level = min(cand_len, 4) # Giới hạn tối đa Heading 1-4
        return True, heading_level


_SENTINEL = object()

class AIFallbackClassifier:
    def __init__(self, api_key: Any = _SENTINEL, base_url: Optional[str] = None, model: Optional[str] = None):
        if api_key is _SENTINEL:
            self.api_key = Config.DEEPSEEK_API_KEY
        else:
            self.api_key = api_key or ""
        self.base_url = (base_url or Config.DEEPSEEK_BASE_URL).rstrip('/')
        self.model = model or Config.DEEPSEEK_MODEL

    def is_configured(self) -> bool:
        return bool(self.api_key and len(self.api_key) > 5)

    def classify_candidates(self, candidates: List[Dict[str, Any]]) -> Dict[int, str]:
        """
        Gom danh sách JSON gửi DeepSeek API phân loại trong 1 request batch duy nhất.
        candidates: [
            {"id": 0, "text": "...", "is_bold": True, "prev_text": "...", "next_text": "..."},
            ...
        ]
        Trả về dict: {id: "HEADING_1"|"HEADING_2"|"HEADING_3"|"BODY"|"CAPTION"|"LIST"}
        """
        results: Dict[int, str] = {}
        if not candidates:
            return results

        # Nếu không có API Key, sử dụng Heuristic Fallback thông minh
        if not self.is_configured():
            logger.info("DeepSeek API Key chưa được cấu hình. Sử dụng bộ phân loại Heuristic nâng cao.")
            return self._heuristic_fallback(candidates)

        try:
            # Chuẩn bị payload theo chuẩn OpenAI-compatible Chat Completion
            prompt_content = (
                "Bạn là chuyên gia phân loại cấu trúc văn bản báo cáo khoa học/kỹ thuật tiếng Việt.\n"
                "Dưới đây là danh sách các đoạn văn bản nghi vấn trong tài liệu Word cần phân loại.\n"
                "Hãy phân loại từng mục vào một trong các nhãn sau:\n"
                "- HEADING_1: Tiêu đề cấp 1 lớn (như Mở đầu, Chương, Kết luận, v.v.)\n"
                "- HEADING_2: Tiêu đề cấp 2 (tiểu mục lớn trong chương)\n"
                "- HEADING_3: Tiêu đề cấp 3 (tiểu mục nhỏ)\n"
                "- HEADING_4: Tiêu đề cấp 4 (tiểu mục rất chi tiết như 1.2.2.1)\n"
                "- CAPTION: Chú thích hình ảnh, bảng biểu, sơ đồ\n"
                "- LIST: Đầu dòng danh sách liệt kê\n"
                "- BODY: Đoạn văn nội dung thông thường\n\n"
                "Trả về KẾT QUẢ DUY NHẤT dưới dạng JSON Array (không kèm markdown thừa, không ```json):\n"
                '[{"id": 0, "type": "HEADING_2"}, {"id": 1, "type": "BODY"}]\n\n'
                f"Danh sách cần phân loại:\n{json.dumps(candidates, ensure_ascii=False, indent=2)}"
            )

            headers = {
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json"
            }

            payload = {
                "model": self.model,
                "messages": [
                    {"role": "system", "content": "You are a precise document structural classifier that only outputs valid JSON."},
                    {"role": "user", "content": prompt_content}
                ],
                "temperature": 0.1,
                "max_tokens": 1500
            }

            url = f"{self.base_url}/chat/completions"
            response = requests.post(url, headers=headers, json=payload, timeout=12)

            if response.status_code == 200:
                res_json = response.json()
                raw_text = res_json['choices'][0]['message']['content'].strip()

                # Trích xuất JSON Array một cách an toàn và chống lỗi định dạng
                json_match = re.search(r'\[[\s\S]*\]', raw_text)
                if json_match:
                    raw_text = json_match.group(0)
                else:
                    raw_text = raw_text.strip()
                    if "```" in raw_text:
                        raw_text = re.sub(r'^```json\s*', '', raw_text, flags=re.MULTILINE)
                        raw_text = re.sub(r'^```\s*', '', raw_text, flags=re.MULTILINE)
                        raw_text = raw_text.strip()

                parsed = json.loads(raw_text)
                for item in parsed:
                    c_id = item.get("id")
                    c_type = item.get("type", "BODY").upper()
                    if c_id is not None:
                        results[c_id] = c_type
                logger.info(f"AI Fallback DeepSeek đã phân loại thành công {len(results)} đoạn nghi vấn.")
                return results
            else:
                logger.warning(f"DeepSeek API trả về mã lỗi {response.status_code}: {response.text}")
                return self._heuristic_fallback(candidates)

        except Exception as e:
            logger.warning(f"Lỗi khi gọi DeepSeek API ({e}). Tự động chuyển sang Heuristic Fallback.")
            return self._heuristic_fallback(candidates)

    def _heuristic_fallback(self, candidates: List[Dict[str, Any]]) -> Dict[int, str]:
        """
        Bộ quy tắc dự phòng khi không có Internet hoặc chưa nhập API Key
        """
        from core.classifier import CAPTION_PATTERN

        results = {}
        for c in candidates:
            c_id = c["id"]
            text = unicodedata.normalize('NFC', c["text"].strip())
            is_bold = c.get("is_bold", False)

            # 1. Kiểm tra Caption (hỗ trợ cả có và không có số thứ tự)
            if CAPTION_PATTERN.match(text):
                results[c_id] = "CAPTION"
                continue

            # 2. Kiểm tra Unnumbered Heading theo từ khóa hoặc chữ hoa toàn bộ
            clean_upper = text.upper().rstrip(':. ')
            if clean_upper in Config.TOC_UNNUMBERED_WHITELIST:
                results[c_id] = "HEADING_1"
                continue
            if clean_upper in Config.NON_TOC_WHITELIST:
                results[c_id] = "HEADING_NON_TOC"
                continue

            # 3. Dòng ngắn, in đậm, không kết thúc bằng dấu câu -> Heading 2 hoặc 3
            if is_bold and len(text) < 90 and not text.endswith((".", ":", ";", ",")):
                # Nếu viết hoa toàn bộ -> Heading 2
                if text.isupper():
                    results[c_id] = "HEADING_2"
                else:
                    results[c_id] = "HEADING_3"
            else:
                results[c_id] = "BODY"

        return results
