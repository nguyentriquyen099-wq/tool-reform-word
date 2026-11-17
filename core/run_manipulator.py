"""
Phân hệ 5 & Lưu ý kỹ thuật số 1: Run-level Manipulation & Style Reset
- Cấm ghi đè paragraph.text = ... làm mất Run formatting (bold/italic cục bộ, OMML math).
- Xử lý Run phân mảnh: duyệt qua các Run đầu tiên cho đến khi xóa hết rác/tab.
- Xử lý ký tự Symbol ẩn: phát hiện và xử lý thẻ <w:sym> (font Wingdings/Symbol) trong XML.
- Reset Style Inheritance: reset override trên Run (set font.name, font.size, font.color = None)
  để kế thừa Style gốc, đồng thời BẢO TỒN triệt để bold/italic cục bộ.
"""

import re
import logging
import unicodedata
from typing import List, Optional, Any
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from config import Config

logger = logging.getLogger(__name__)

# Regex các ký tự bullet rác thường gặp và marker danh sách
BULLET_CHARS_PATTERN = re.compile(
    r'^[\s\t\xa0\u2002\u2003\u2009\u200a\u3000•●○▪❖➢★–—\-\+\*►\.\>]+'
    r'|'
    r'^[\s\t\xa0\u2002\u2003\u2009\u200a\u3000]*(\([a-zA-Z0-9]+\)|\[[a-zA-Z0-9]+\]|[a-zA-Z0-9]+\))\s*'
)

# Regex nhận diện Chương có số Ả Rập ở đầu (CHƯƠNG 1, Chương 2, CHAPTER 3...)
CHAPTER_ARABIC_PATTERN = re.compile(
    r'^\s*(chương|chapter)\s+(\d+)\b',
    re.IGNORECASE
)

# Regex tách tiền tố số đề mục (1.2., 1.2.1., A.,...) và phần nội dung text
HEADING_PREFIX_PATTERN = re.compile(
    r'^(\s*(?:\d+(?:\.\d+)*\.?|[A-Za-z]\.|\([A-Za-z0-9]+\))\s*)(.*)$',
    re.DOTALL
)

# Danh sách từ viết tắt kỹ thuật / học thuật cần giữ nguyên chữ hoa
TECH_ACRONYMS = {
    'LAN', 'WAN', 'VLAN', 'WLAN', 'MAN', 'SAN', 'IP', 'IPV4', 'IPV6',
    'TCP', 'UDP', 'DNS', 'DHCP', 'HTTP', 'HTTPS', 'FTP', 'TFTP', 'SSH', 'OSI',
    'VPN', 'NAT', 'ACL', 'RIP', 'OSPF', 'BGP', 'EIGRP', 'ICMP', 'IGMP', 'ARP', 'RARP',
    'SNMP', 'SMTP', 'POP3', 'IMAP', 'SSL', 'TLS', 'PKI', 'CA', 'DMZ', 'IDS', 'IPS',
    'SIEM', 'SOC', 'NOC', 'STP', 'RSTP', 'MSTP', 'VTP', 'HSRP', 'VRRP', 'GLBP',
    'LLDP', 'CDP', 'LACP', 'PAGP', 'POE', 'QOS', 'COS', 'TOS', 'DSCP', 'MTU', 'MSS', 'TTL',
    'SDN', 'SD-WAN', 'MPLS', 'VXLAN', 'VNI', 'AP', 'SSID', 'BSSID', 'ESSID', 'WEP',
    'WPA', 'WPA2', 'WPA3', 'RADIUS', 'TACACS', 'AAA', 'ISP', 'SLA', 'CDN',
    'CNTT', 'ATTT', 'CSDL', 'AI', 'ML', 'DL', 'IOT', 'IIOT', 'API', 'SDK', 'UI', 'UX',
    'CPU', 'RAM', 'ROM', 'SSD', 'HDD', 'PC', 'MAC', 'ID', 'RFID', 'NFC', 'IT',
    'WIFI', 'WI-FI', 'B2B', 'B2C', 'ERP', 'CRM', 'SCM', 'HRM',
    'SQL', 'NOSQL', 'RDBMS', 'HTML', 'CSS', 'JS', 'JSON', 'XML', 'YAML',
    'IEEE', 'ISO', 'IETF', 'W3C', 'ITU', 'BKHN', 'ĐHBK', 'CNPM', 'KTPM', 'HTTT', 'KHMT', 'PCCC'
}

ROMAN_NUMERALS_SET = {
    'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X',
    'XI', 'XII', 'XIII', 'XIV', 'XV', 'XVI', 'XVII', 'XVIII', 'XIX', 'XX'
}

# Các từ tiếng Việt thường đi liền với lan / san / ai để phân biệt từ vựng với từ viết tắt
VIETNAMESE_WORDS_PREV_LAN = {'lây', 'tràn', 'lan', 'mở', 'bành', 'hoa', 'cây', 'bình', 'nhánh'}
VIETNAMESE_WORDS_NEXT_LAN = {'tỏa', 'rộng', 'can', 'truyền', 'nhanh', 'khắp', 'ra'}
VIETNAMESE_WORDS_NEXT_SAN = {'bằng', 'sẻ', 'lấp', 'hô', 'phẳng', 'hợp', 'sát'}
VIETNAMESE_WORDS_MODAL_SAN = {'được', 'đang', 'đã', 'sẽ', 'phải', 'để', 'cần'}
VIETNAMESE_WORDS_NEXT_AI = {'cũng', 'đó', 'nào', 'là', 'sẽ', 'phải', 'trong', 'được', 'muốn', 'bảo'}
VIETNAMESE_WORDS_PREV_AI = {'bất', 'kỳ', 'kể', 'cho', 'của', 'hỏi', 'biết', 'do'}

def is_predominantly_upper(text: str) -> bool:
    """
    Kiểm tra xem chuỗi có phải được gõ chủ yếu bằng chữ IN HOA (ALL CAPS) hay không.
    Bảo đảm bắt trúng cả trường hợp người dùng gõ Caps Lock nhưng có xen kẽ từ viết tắt/mã chứa chữ thường (IPv6, Wi-Fi...).
    """
    alphas = [c for c in text if c.isalpha()]
    if not alphas:
        return False
    uppers = [c for c in alphas if c.isupper()]
    if len(alphas) == len(uppers):
        return True
    return len(alphas) >= 6 and (len(uppers) / len(alphas)) >= 0.80

def to_sentence_case(title: str) -> str:
    """
    Chuyển đổi chuỗi ALL CAPS sang Sentence Case chuẩn tiếng Việt.
    - Chữ cái đầu câu viết hoa, viết hoa chữ cái đầu sau dấu hai chấm (:), gạch ngang (–, —), dấu chấm (.).
    - Các từ thông thường còn lại viết thường.
    - Bảo tồn các từ viết tắt kỹ thuật (LAN, IP, TCP, CNTT, IT, WIFI, IPv6...) và số La Mã (I, II, III...).
    - Khử nhập nhằng ngữ cảnh giữa từ viết tắt và từ tiếng Việt (lây lan vs mạng LAN, san bằng vs mạng SAN, ai đó vs mô hình AI).
    - Bảo tồn các mã định danh chứa cả chữ và số (ví dụ: 302-A9, A9, v1...).
    """
    if not title or not is_predominantly_upper(title):
        return title

    chars = list(title.lower())
    res = ''.join(chars)

    # Viết hoa chữ cái đầu câu và chữ cái đầu sau dấu hai chấm (:), gạch ngang (–, —), dấu chấm (.)
    res = re.sub(r'(^|[:\.–—]\s+)([a-zà-ỹ])', lambda m: m.group(1) + m.group(2).upper(), res)

    # Thu thập danh sách token với vị trí bắt đầu chính xác để tra cứu O(1)
    token_matches = list(re.finditer(r'\b[a-zA-Z0-9\-_À-ỹ]+\b', res))
    token_count = len(token_matches)
    start_to_idx = {tm.start(): i for i, tm in enumerate(token_matches)}

    def restore_special(m):
        w = m.group(0)
        w_up = w.upper()

        if w_up in TECH_ACRONYMS or w_up in ROMAN_NUMERALS_SET:
            # Khử nhập nhằng cho các từ đồng âm tiếng Việt
            idx = start_to_idx.get(m.start(), -1)
            prev_w = token_matches[idx - 1].group(0).lower() if idx > 0 else ''
            next_w = token_matches[idx + 1].group(0).lower() if idx >= 0 and idx + 1 < token_count else ''

            if w_up == 'LAN' and (prev_w in VIETNAMESE_WORDS_PREV_LAN or next_w in VIETNAMESE_WORDS_NEXT_LAN):
                return w
            if w_up == 'SAN' and (next_w in VIETNAMESE_WORDS_NEXT_SAN or prev_w in VIETNAMESE_WORDS_MODAL_SAN):
                return w
            if w_up == 'AI' and (next_w in VIETNAMESE_WORDS_NEXT_AI or prev_w in VIETNAMESE_WORDS_PREV_AI):
                return w

            return w_up

        if any(c.isdigit() for c in w) and any(c.isalpha() for c in w):
            return m.string[m.start():m.end()].upper()

        return w

    res = re.sub(r'\b[a-zA-Z0-9\-_À-ỹ]+\b', restore_special, res)
    return res

def int_to_roman(num: Any) -> str:
    """
    Chuyển đổi số nguyên dương sang số La Mã (1 -> I, 2 -> II, 3 -> III, 4 -> IV, 5 -> V,... 10 -> X...).
    Nếu đầu vào không phải số nguyên dương hợp lệ (hoặc là bool), trả về chuỗi str(num).
    """
    if isinstance(num, bool):
        return str(num)
    try:
        val_int = int(num)
    except (ValueError, TypeError):
        return str(num)

    if val_int <= 0 or val_int > 3999:
        return str(num)

    val_map = [
        (1000, 'M'), (900, 'CM'), (500, 'D'), (400, 'CD'),
        (100, 'C'), (90, 'XC'), (50, 'L'), (40, 'XL'),
        (10, 'X'), (9, 'IX'), (5, 'V'), (4, 'IV'), (1, 'I')
    ]
    res = []
    rem = val_int
    for val, roman in val_map:
        while rem >= val:
            res.append(roman)
            rem -= val
    return "".join(res)

class RunManipulator:
    @staticmethod
    def clean_leading_junk_and_symbols(paragraph, remove_bullets: bool = True) -> str:
        """
        Duyệt qua các Run đầu tiên của đoạn văn để:
        1. Xóa thẻ <w:sym> (ký tự symbol ẩn dạng Wingdings/Symbol) nếu ở đầu đoạn.
        2. Xóa các tab thừa (\t), khoảng trắng thừa (\s, \xa0) và ký tự bullet rác.
        3. Tuyệt đối KHÔNG gán paragraph.text = ..., giữ nguyên các Run còn lại.
        Trả về: Chuỗi bullet hoặc prefix đã được loại bỏ (nếu có).
        """
        removed_prefix = ""
        runs = paragraph.runs
        if not runs:
            # Kiểm tra trường hợp đặc biệt: paragraph có thẻ w:sym trực tiếp trong XML
            syms = paragraph._element.xpath('.//w:sym')
            if syms and remove_bullets:
                for s in syms:
                    parent = s.getparent()
                    if parent is not None:
                        parent.remove(s)
            return removed_prefix

        # Duyệt từng Run từ đầu
        for run in runs:
            r_elem = run._element

            # 1. Kiểm tra và loại bỏ thẻ <w:sym> (thường là bullet icon ẩn)
            sym_elems = r_elem.xpath('./w:sym')
            if sym_elems and remove_bullets:
                for s in sym_elems:
                    r_elem.remove(s)
                    removed_prefix += "•"

            # 2. Xử lý text trong Run
            text = run.text
            if not text:
                # Nếu run rỗng sau khi bỏ sym, bỏ qua tiếp tục run sau
                continue

            if remove_bullets:
                # Tìm xem có khoảng trắng hoặc bullet ở đầu không
                match = BULLET_CHARS_PATTERN.match(text)
                if match:
                    matched_str = match.group(0)
                    removed_prefix += matched_str
                    remaining_text = text[len(matched_str):]
                    run.text = remaining_text

                    # Nếu sau khi cắt mà run vẫn còn ký tự nội dung -> dừng duyệt
                    if remaining_text:
                        break
                else:
                    # Run này bắt đầu bằng chữ cái bình thường -> dừng duyệt
                    break
            else:
                # Chỉ xóa tab và khoảng trắng thụt lề thừa (kể cả unicode spaces) ở đầu đoạn
                stripped = text.lstrip('\t \xa0\u2002\u2003\u2009\u200a\u3000')
                if len(stripped) != len(text):
                    removed_prefix += text[:len(text) - len(stripped)]
                    run.text = stripped
                if stripped:
                    break

        return removed_prefix

    @staticmethod
    def prepend_text_to_paragraph(paragraph, prefix_text: str):
        """
        Thêm text vào đầu đoạn văn một cách an toàn (trên Run đầu tiên có nội dung).
        Nếu chưa có Run, tạo Run mới.
        """
        if not prefix_text:
            return

        runs = paragraph.runs
        if runs:
            target_run = None
            for r in runs:
                if r.text:
                    target_run = r
                    break
            if target_run is None:
                target_run = runs[0]

            target_run.text = prefix_text + target_run.text
            # Đảm bảo target_run không giữ font Wingdings hoặc Symbol rác
            rPr = target_run._element.find(qn('w:rPr'))
            if rPr is not None:
                rfonts = rPr.find(qn('w:rFonts'))
                if rfonts is not None:
                    ascii_font = (rfonts.get(qn('w:ascii')) or '').lower()
                    if 'wingdings' in ascii_font or 'symbol' in ascii_font:
                        rPr.remove(rfonts)
        else:
            new_run = paragraph.add_run(prefix_text)
            new_run.font.name = Config.FONT_NAME

    @staticmethod
    def clean_trailing_whitespace(paragraph):
        """
        Xóa khoảng trắng thừa / tab / non-breaking space (\xa0) ở cuối đoạn văn.
        Đặc biệt quan trọng cho Heading để tránh rác khoảng trắng trong Mục lục và Navigation Pane.
        """
        runs = paragraph.runs
        for run in reversed(runs):
            if not run.text:
                continue
            stripped = run.text.rstrip('\t \xa0\u2002\u2003\u2009\u200a\u3000')
            if len(stripped) != len(run.text):
                run.text = stripped
            if run.text:
                break

    @staticmethod
    def reset_run_style_inheritance(paragraph, keep_bold: bool = True, keep_italic: bool = True):
        """
        Reset Style Inheritance:
        Xóa các override trực tiếp trên Run (font name, font size, font color)
        và trong paragraph-level pPr/rPr để Run kế thừa chuẩn từ Style gốc (styles.xml).
        BẢO TỒN triệt để:
        - bold cục bộ (nếu keep_bold=True)
        - italic cục bộ (nếu keep_italic=True)
        - math tags (<m:oMath>) không bị đụng tới
        """
        # Xóa các thuộc tính rPr trong paragraph-level pPr ghi đè cỡ chữ/màu kế thừa
        pPr = paragraph._element.find(qn('w:pPr'))
        if pPr is not None:
            p_rPr = pPr.find(qn('w:rPr'))
            if p_rPr is not None:
                for tag in ['sz', 'szCs', 'color']:
                    elem = p_rPr.find(qn(f'w:{tag}'))
                    if elem is not None:
                        p_rPr.remove(elem)

        for run in paragraph.runs:
            # Lưu lại trạng thái bold và italic cục bộ của run
            orig_bold = run.bold
            orig_italic = run.italic

            # Reset override font name và size để kế thừa từ paragraph style
            run.font.name = None
            run.font.size = None
            if run.font.color and run.font.color.rgb:
                run.font.color.rgb = None

            # Xóa các thuộc tính rFonts ghi đè trực tiếp trong XML của Run
            rPr = run._element.find(qn('w:rPr'))
            if rPr is not None:
                rfonts = rPr.find(qn('w:rFonts'))
                if rfonts is not None:
                    rPr.remove(rfonts)

                # Xóa color ghi đè nếu có
                color = rPr.find(qn('w:color'))
                if color is not None:
                    rPr.remove(color)

                # Xóa sz (size) ghi đè nếu có
                sz = rPr.find(qn('w:sz'))
                if sz is not None:
                    rPr.remove(sz)
                szCs = rPr.find(qn('w:szCs'))
                if szCs is not None:
                    rPr.remove(szCs)

            # Khôi phục bold / italic nếu người dùng có định dạng riêng
            if keep_bold and orig_bold is True:
                run.bold = True
            elif keep_bold and orig_bold is False:
                # Nếu rõ ràng là False thì giữ nguyên
                run.bold = False

            if keep_italic and orig_italic is True:
                run.italic = True
            elif keep_italic and orig_italic is False:
                run.italic = False

    @staticmethod
    def convert_chapter_number_to_roman(paragraph) -> bool:
        """
        Chuẩn hóa số Chương từ số Ả Rập sang số La Mã (CHƯƠNG 1 -> CHƯƠNG I, CHƯƠNG 2 -> CHƯƠNG II...)
        theo đúng quy chuẩn báo cáo / đồ án học thuật Việt Nam.
        - Tuân thủ nghiêm ngặt Phân hệ 5: Thao tác trên Run (không dùng paragraph.text = ... làm mất run formatting).
        - Giữ nguyên các run khác và định dạng bold/italic/color/math nội bộ.
        - Nếu tiêu đề đã là số La Mã (CHƯƠNG I, II, III...) hoặc không phải Chương thì giữ nguyên.
        - Giữ nguyên các tiểu mục con dạng số Ả Rập (1.1, 1.2, 1.2.1...).
        - Chuẩn hóa Unicode NFC cho các run để tương thích hoàn toàn chuẩn tiếng Việt và chống lệch slice offset.
        Trả về True nếu có thực hiện chuyển đổi, False nếu không.
        """
        runs = paragraph.runs
        # Chuẩn hóa NFC cho từng Run để xử lý triệt để mã NFD tổ hợp và chống lệch offset
        for r in runs:
            if r.text:
                r.text = unicodedata.normalize('NFC', r.text)

        full_text = "".join(r.text for r in runs) if runs else unicodedata.normalize('NFC', paragraph.text)
        if not full_text:
            return False

        match = CHAPTER_ARABIC_PATTERN.search(full_text)
        if not match:
            return False

        target_start = match.start(1)
        target_end = match.end(2)
        kw = match.group(1).upper()
        num = int(match.group(2))
        roman = int_to_roman(num)
        replacement = f"{kw} {roman}"

        if not runs:
            paragraph.text = full_text[:target_start] + replacement + full_text[target_end:]
            return True

        curr_pos = 0
        first_overlap = True
        for r in runs:
            r_len = len(r.text)
            r_start = curr_pos
            r_end = curr_pos + r_len
            curr_pos = r_end

            if r_end <= target_start or r_start >= target_end:
                continue

            p_slice_start = max(0, target_start - r_start)
            s_slice_start = max(0, target_end - r_start)

            prefix = r.text[:p_slice_start]
            suffix = r.text[s_slice_start:]

            if first_overlap:
                r.text = prefix + replacement + suffix
                first_overlap = False
            else:
                r.text = prefix + suffix

        return True

    @staticmethod
    def normalize_heading_casing(paragraph) -> bool:
        """
        Chuẩn hóa chữ hoa/thường cho Heading cấp con (Heading 2, 3...):
        - Nếu đề mục bị gõ HOA TOÀN BỘ (ALL CAPS), chuyển về dạng Sentence Case chuẩn
          (chỉ viết hoa chữ cái đầu sau số đề mục, bảo tồn từ viết tắt như LAN, IP, CNTT...).
        - Giữ nguyên số đề mục (1.2., 1.2.1., A....).
        - Tuyệt đối tuân thủ Phân hệ 5: Thao tác an toàn trên các Run, bảo tồn đối tượng Run,
          bold, italic, style inheritance và không làm mất run formatting.
        - Chuẩn hóa NFC cho các run để chống lệch encoding tiếng Việt.
        Trả về True nếu có thực hiện chuẩn hóa, False nếu giữ nguyên.
        """
        runs = paragraph.runs
        if not runs:
            full_text = unicodedata.normalize('NFC', paragraph.text)
            if not full_text:
                return False
            m = HEADING_PREFIX_PATTERN.match(full_text)
            if m:
                prefix = m.group(1)
                title = m.group(2)
            else:
                prefix = ""
                title = full_text
            new_title = to_sentence_case(title)
            if new_title != title:
                paragraph.text = prefix + new_title
                return True
            return False

        for r in runs:
            if r.text:
                r.text = unicodedata.normalize('NFC', r.text)

        full_text = "".join(r.text for r in runs)
        if not full_text:
            return False

        m = HEADING_PREFIX_PATTERN.match(full_text)
        if m:
            prefix = m.group(1)
            title = m.group(2)
        else:
            prefix = ""
            title = full_text

        new_title = to_sentence_case(title)
        if new_title == title:
            return False

        new_full_text = prefix + new_title

        # Cập nhật lại text vào các runs bảo tồn cấu trúc runs
        if len(new_full_text) == len(full_text):
            pos = 0
            for r in runs:
                r_len = len(r.text)
                r.text = new_full_text[pos : pos + r_len]
                pos += r_len
        else:
            pos = 0
            for i, r in enumerate(runs):
                if i == len(runs) - 1:
                    r.text = new_full_text[pos:]
                else:
                    r_len = len(r.text)
                    r.text = new_full_text[pos : pos + r_len]
                    pos += r_len

        return True
