"""
Unit tests cho Phân hệ 3: Hierarchical State Machine (HSM) & AI Fallback
"""

import unittest
from core.hsm import HierarchicalStateMachine, AIFallbackClassifier

class TestHSM(unittest.TestCase):
    def setUp(self):
        self.hsm = HierarchicalStateMachine()

    def test_valid_progression(self):
        # [1] -> valid Heading 1
        valid, lvl = self.hsm.validate_and_update([1], "Chương 1 Tổng quan", is_bold=True)
        self.assertTrue(valid)
        self.assertEqual(lvl, 1)

        # [1, 1] -> valid Heading 2
        valid, lvl = self.hsm.validate_and_update([1, 1], "1.1 Bối cảnh đề tài", is_bold=True)
        self.assertTrue(valid)
        self.assertEqual(lvl, 2)

        # [1, 2] -> valid Heading 2
        valid, lvl = self.hsm.validate_and_update([1, 2], "1.2 Mục tiêu nghiên cứu", is_bold=True)
        self.assertTrue(valid)
        self.assertEqual(lvl, 2)

        # [2] -> valid Heading 1
        valid, lvl = self.hsm.validate_and_update([2], "Chương 2 Cơ sở lý thuyết", is_bold=True)
        self.assertTrue(valid)
        self.assertEqual(lvl, 1)

    def test_numbered_list_trap_avoided(self):
        """
        Bẫy quan trọng: Đang ở Chương 2 / mục 2.1, xuất hiện chuỗi "1. Bước một...", "2. Bước hai...", "3. Bước ba..."
        TUYỆT ĐỐI không mục nào được bị gán nhầm làm Heading 1!
        """
        # Thiết lập state đang ở Chương 2
        self.hsm.validate_and_update([2], "Chương 2 Cơ sở lý thuyết", is_bold=True)
        self.hsm.validate_and_update([2, 1], "2.1 Khái niệm cơ bản", is_bold=True)

        # Gặp câu danh sách con số "1."
        valid1, lvl1 = self.hsm.validate_and_update(
            [1],
            "1. Bước một: Thu thập và chuẩn bị toàn bộ tập dữ liệu huấn luyện cho hệ thống.",
            is_bold=False
        )
        self.assertFalse(valid1)
        self.assertEqual(lvl1, 0)
        self.assertEqual(self.hsm.current_path, [2, 1])

        # Gặp câu danh sách con số "2." tiếp theo ngay sau đó (bẫy mà attempt trước bị lỗi)
        valid2, lvl2 = self.hsm.validate_and_update(
            [2],
            "2. Bước hai: Xử lý làm sạch và chuẩn hóa nhãn dữ liệu theo định dạng chuẩn.",
            is_bold=False
        )
        self.assertFalse(valid2, "Mục 2. của danh sách con tuyệt đối không được thành Heading 1!")
        self.assertEqual(lvl2, 0)
        self.assertEqual(self.hsm.current_path, [2, 1])

        # Gặp câu danh sách con số "3." tiếp theo
        valid3, lvl3 = self.hsm.validate_and_update(
            [3],
            "3. Bước ba: Huấn luyện mô hình và đánh giá độ chính xác thực tế.",
            is_bold=False
        )
        self.assertFalse(valid3, "Mục 3. của danh sách con tuyệt đối không được thành Heading 1!")
        self.assertEqual(lvl3, 0)
        self.assertEqual(self.hsm.current_path, [2, 1])

        # Khi gặp Chương 3 thực sự -> Phải chấp nhận Heading 1 bình thường
        valid_ch3, lvl_ch3 = self.hsm.validate_and_update(
            [3],
            "CHƯƠNG 3: KẾT QUẢ THỬ NGHIỆM VÀ ĐÁNH GIÁ",
            is_bold=True
        )
        self.assertTrue(valid_ch3)
        self.assertEqual(lvl_ch3, 1)
        self.assertEqual(self.hsm.current_path, [3])

    def test_level_jump_rejected(self):
        # Đang ở [1], nhảy cóc lên [1, 1, 1, 1]
        self.hsm.validate_and_update([1], "Chương 1", is_bold=True)
        valid, lvl = self.hsm.validate_and_update([1, 1, 1, 1], "1.1.1.1 Chi tiết sâu", is_bold=True)
        self.assertFalse(valid)

    def test_ai_fallback_heuristics(self):
        ai = AIFallbackClassifier(api_key=None) # Chế độ offline
        candidates = [
            {"id": 1, "text": "KẾT LUẬN", "is_bold": True},
            {"id": 2, "text": "Hình 2.1: Sơ đồ hoạt động của mạng neural", "is_bold": False},
            {"id": 3, "text": "Đặc điểm kiến trúc hệ thống", "is_bold": True},
            {"id": 4, "text": "Đây là một đoạn văn bản bình thường không phải tiêu đề.", "is_bold": False}
        ]
        res = ai.classify_candidates(candidates)
        self.assertEqual(res[1], "HEADING_1")
        self.assertEqual(res[2], "CAPTION")
        self.assertEqual(res[3], "HEADING_3") # Hoặc HEADING_2
        self.assertEqual(res[4], "BODY")

if __name__ == '__main__':
    unittest.main()
