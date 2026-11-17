"""
Bộ phân tích numbering.xml của tài liệu Word (.docx)
Giải quyết Lưu ý kỹ thuật số 3: Xử lý Auto-Numbering thông minh
- Tra cứu numId/ilvl trong numbering.xml
- Phân biệt decimal vs bullet
- Dựa vào ilvl để xác định Style Heading tương ứng
"""

import logging
from typing import Dict, Optional, Tuple, Any
from lxml import etree

logger = logging.getLogger(__name__)

# OpenXML Namespaces
NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}

def _xpath(elem, xpath_expr: str):
    """
    Hỗ trợ gọi xpath an toàn cho cả lxml.etree._Element (cần namespaces)
    và BaseOxmlElement của python-docx (không hỗ trợ kwarg namespaces).
    """
    try:
        return elem.xpath(xpath_expr, namespaces=NS)
    except TypeError:
        return elem.xpath(xpath_expr)

class NumberingParser:
    def __init__(self, doc):
        """
        Khởi tạo parser với đối tượng docx Document
        """
        self.doc = doc
        self.num_to_abstract: Dict[int, int] = {}
        self.abstract_num_levels: Dict[int, Dict[int, Dict[str, Any]]] = {}
        self._parsed = False
        self._parse()

    def _parse(self):
        """
        Đọc và lập chỉ mục cấu trúc numbering.xml
        """
        try:
            # Truy xuất numbering part từ docx package
            numbering_part = None
            for rel in self.doc.part.rels.values():
                if "numbering" in rel.reltype:
                    numbering_part = rel.target_ref
                    break

            if numbering_part is None:
                # Tìm trong part-related items
                try:
                    num_part = self.doc.part.numbering_part
                    numbering_xml = num_part._element
                except Exception:
                    numbering_xml = None
            else:
                try:
                    num_part = self.doc.part.related_parts[numbering_part]
                    numbering_xml = num_part._element
                except Exception:
                    numbering_xml = None

            if numbering_xml is None:
                # Kiểm tra fallback qua element trực tiếp nếu có
                try:
                    num_part = getattr(self.doc.part, 'numbering_part', None)
                    if num_part:
                        numbering_xml = num_part._element
                except Exception:
                    pass

            if numbering_xml is None:
                return

            # 1. Map w:num -> w:abstractNumId
            num_elements = _xpath(numbering_xml, '//w:num')
            for num in num_elements:
                num_id_val = num.get(f"{{{NS['w']}}}numId")
                if not num_id_val:
                    continue
                try:
                    num_id = int(num_id_val)
                except ValueError:
                    continue

                abs_ref = _xpath(num, './w:abstractNumId')
                if abs_ref:
                    abs_val = abs_ref[0].get(f"{{{NS['w']}}}val")
                    if abs_val is not None:
                        try:
                            self.num_to_abstract[num_id] = int(abs_val)
                        except ValueError:
                            pass

            # 2. Map w:abstractNum -> các cấp độ w:lvl
            abstract_elements = _xpath(numbering_xml, '//w:abstractNum')
            for abs_elem in abstract_elements:
                abs_id_val = abs_elem.get(f"{{{NS['w']}}}abstractNumId")
                if not abs_id_val:
                    continue
                try:
                    abs_id = int(abs_id_val)
                except ValueError:
                    continue

                levels = {}
                lvl_elements = _xpath(abs_elem, './w:lvl')
                for lvl in lvl_elements:
                    ilvl_val = lvl.get(f"{{{NS['w']}}}ilvl")
                    if not ilvl_val:
                        continue
                    try:
                        ilvl = int(ilvl_val)
                    except ValueError:
                        continue

                    # numFmt (decimal, bullet, lowerLetter, etc.)
                    num_fmt_elem = _xpath(lvl, './w:numFmt')
                    num_fmt = num_fmt_elem[0].get(f"{{{NS['w']}}}val") if num_fmt_elem else "decimal"

                    # lvlText (%1., %1.%2, etc.)
                    lvl_text_elem = _xpath(lvl, './w:lvlText')
                    lvl_text = lvl_text_elem[0].get(f"{{{NS['w']}}}val") if lvl_text_elem else ""

                    # pStyle (liên kết style nếu có)
                    pstyle_elem = _xpath(lvl, './w:pStyle')
                    p_style = pstyle_elem[0].get(f"{{{NS['w']}}}val") if pstyle_elem else None

                    levels[ilvl] = {
                        "numFmt": num_fmt.lower() if num_fmt else "decimal",
                        "lvlText": lvl_text,
                        "pStyle": p_style
                    }

                self.abstract_num_levels[abs_id] = levels

            self._parsed = True
        except Exception as e:
            logger.debug(f"Lỗi khi parse numbering.xml: {e}")

    def get_info(self, num_id: Optional[int], ilvl: Optional[int]) -> Dict[str, Any]:
        """
        Tra cứu thông tin numbering dựa trên numId và ilvl.
        Trả về dict:
        {
            "is_numbered": bool,
            "is_bullet": bool,
            "is_decimal": bool,
            "num_fmt": str,
            "heading_level": Optional[int], # 1, 2, 3 nếu là phân cấp số
            "lvl_text": str
        }
        """
        res = {
            "is_numbered": False,
            "is_bullet": False,
            "is_decimal": False,
            "num_fmt": "none",
            "heading_level": None,
            "lvl_text": "",
            "p_style": ""
        }

        if num_id is None:
            return res

        res["is_numbered"] = True

        abs_id = self.num_to_abstract.get(num_id)
        if abs_id is None:
            # Không tìm thấy abstractNumId, suy đoán mặc định
            return res

        levels = self.abstract_num_levels.get(abs_id, {})
        ilvl_key = ilvl if ilvl is not None else 0
        lvl_info = levels.get(ilvl_key, {})

        num_fmt = lvl_info.get("numFmt", "")
        lvl_text = lvl_info.get("lvlText", "")
        p_style = lvl_info.get("pStyle", "")

        res["num_fmt"] = num_fmt
        res["lvl_text"] = lvl_text
        res["p_style"] = p_style

        if num_fmt == "bullet":
            res["is_bullet"] = True
        elif num_fmt in ["decimal", "decimalzeropadded", "upperroman", "lowerroman"]:
            res["is_decimal"] = True
            # Nếu có liên kết tới style heading hoặc định dạng số phân cấp
            if p_style and "heading" in p_style.lower():
                # Suy ra heading_level từ tên style
                for i in range(1, 10):
                    if str(i) in p_style:
                        res["heading_level"] = i
                        break
            elif res["is_decimal"]:
                # ilvl 0 -> Heading 1, ilvl 1 -> Heading 2, ilvl 2 -> Heading 3
                res["heading_level"] = ilvl_key + 1

        return res

    def extract_num_pr(self, paragraph) -> Tuple[Optional[int], Optional[int]]:
        """
        Trích xuất numId và ilvl từ thẻ <w:numPr> của paragraph
        """
        p_elem = paragraph._element
        num_pr_list = _xpath(p_elem, './w:pPr/w:numPr')
        if not num_pr_list:
            return None, None

        num_pr = num_pr_list[0]
        num_id_elems = _xpath(num_pr, './w:numId')
        ilvl_elems = _xpath(num_pr, './w:ilvl')

        num_id = None
        ilvl = 0

        if num_id_elems:
            val = num_id_elems[0].get(f"{{{NS['w']}}}val")
            if val is not None:
                try:
                    num_id = int(val)
                except ValueError:
                    pass

        if ilvl_elems:
            val = ilvl_elems[0].get(f"{{{NS['w']}}}val")
            if val is not None:
                try:
                    ilvl = int(val)
                except ValueError:
                    ilvl = 0

        return num_id, ilvl
