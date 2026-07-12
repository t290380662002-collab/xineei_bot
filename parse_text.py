# -*- coding: utf-8 -*-
"""
解析使用者貼上的「訂房文字格式」，例如：

    飯店：巴黎人酒店
    入住：7月15日
    退房：7月17日
    房型：香槟套房/里昂套房（双床）
    件數：1间2晚
    備注：高楼层
    是否吸煙：吸烟房
    入住者中文：吴婷婷
    入住者英文：WU, TINGTING
    出生年月日：1982.07.04
    證件號碼：C73960683

特性：
  - 支援全形/半形冒號與 = 號；欄位順序不拘
  - 容錯貼上時換行消失導致的「標籤併接」（如「飯店：入住：7月15日」）
  - 日期自動正規化（7月15日 / 1982.07.04 → 2026/07/15 / 1982/07/04）
  - 內建常見別名（酒店→飯店、姓名→入住者中文 …）
輸出結構與 fill_booking() 完全相容。
"""
import re
from datetime import datetime

# 標準欄位（對應 fill_booking 的 booking dict key）
STANDARD = [
    "入住", "退房", "飯店", "房型", "件數", "備注", "是否吸煙",
    "入住者中文", "入住者英文", "出生年月日", "證件號碼",
]

# 別名 → 標準欄位
ALIASES = {
    "入住日期": "入住", "checkin": "入住", "check in": "入住", "check-in": "入住",
    "退房日期": "退房", "checkout": "退房", "check out": "退房", "check-out": "退房",
    "酒店": "飯店", "饭店": "飯店", "旅館": "飯店", "hotel": "飯店",
    "房間數": "件數", "房数": "件數", "數量": "件數", "数量": "件數",
    "間數": "件數", "房": "件數", "rooms": "件數",
    "備註": "備注", "备注": "備注", "note": "備注", "remark": "備注",
    "吸煙": "是否吸煙", "吸烟": "是否吸煙", "smoking": "是否吸煙",
    "中文姓名": "入住者中文", "姓名": "入住者中文", "住客中文": "入住者中文", "名字": "入住者中文",
    "英文姓名": "入住者英文", "住客英文": "入住者英文", "ename": "入住者英文",
    "出生": "出生年月日", "生日": "出生年月日", "出生日期": "出生年月日",
    "dob": "出生年月日", "birth": "出生年月日",
    "證件": "證件號碼", "证件": "證件號碼", "證件號": "證件號碼", "证件号": "證件號碼",
    "id": "證件號碼", "護照": "證件號碼", "passport": "證件號碼",
}

# 所有可用作「標籤」的字串（標準欄位 + 別名），用於整段掃描
_ALL = sorted(set(STANDARD + list(ALIASES.keys())), key=len, reverse=True)
_LABEL_ALT = "|".join(re.escape(k) for k in _ALL)
# 掃描整段：捕捉 標籤[:：=]值，值到「下一個標籤」或文末為止（容錯換行消失）
_PAIR_RE = re.compile(
    r"(?P<lab>" + _LABEL_ALT + r")[:：=＝]\s*(?P<val>.*?)(?=(?P<lab2>" + _LABEL_ALT + r")[:：=＝]|\Z)",
    re.DOTALL,
)


def _match_field(label):
    key = label.strip()
    if key in STANDARD:
        return key
    low = key.lower()
    for alias, std in ALIASES.items():
        if alias.lower() == low:
            return std
    return None


def _extract_pairs(text):
    """回傳 [(field, value), ...]，順序依文中出現順序。"""
    pairs = []
    for m in _PAIR_RE.finditer(text or ""):
        field = _match_field(m.group("lab"))
        if field is None:
            continue
        pairs.append((field, m.group("val").strip()))
    return pairs


def looks_like_booking(text):
    """是否像一筆訂房文字（至少含 4 個可辨識欄位才當作訂房）。"""
    if not text:
        return False
    found = sum(1 for f, _ in _extract_pairs(text))
    return found >= 4


def _norm_smoking(value):
    s = (value or "").strip()
    if any(k in s for k in ("不", "無", "无", "no", "NO", "N", "non")):
        return "不吸煙"
    return "吸煙"


def _norm_rooms(value):
    s = str(value or "").strip()
    if not s:
        return ""
    # 先抓「數字 + 间/房/件」前的數字（房數）
    m = re.search(r"(\d+)\s*(?:间|間|房|件|室)", s)
    if m:
        return m.group(1)
    m = re.search(r"\d+", s)
    return m.group(0) if m else s


def _norm_date(value):
    s = re.sub(r"\s+", "", str(value or "")).strip()
    if not s:
        return ""
    # 完整日期 YYYY/M/D 或 YYYY.M.D
    m = re.match(r"^(\d{4})[./\-](\d{1,2})[./\-](\d{1,2})$", s)
    if m:
        y, mo, d = m.groups()
        return f"{int(y):04d}/{int(mo):02d}/{int(d):02d}"
    # 中文 M月D日（無年份）→ 補今年
    m = re.match(r"^(\d{1,2})月(\d{1,2})日?$", s)
    if m:
        mo, d = int(m.group(1)), int(m.group(2))
        return f"{datetime.now().year:04d}/{mo:02d}/{d:02d}"
    # M/D 或 M.D（無年份）→ 補今年
    m = re.match(r"^(\d{1,2})[./\-](\d{1,2})$", s)
    if m:
        mo, d = int(m.group(1)), int(m.group(2))
        return f"{datetime.now().year:04d}/{mo:02d}/{d:02d}"
    return s  # 無法識別就原樣保留


def parse_booking_text(text):
    """把訂房文字解析成 fill_booking() 相容的 dict。"""
    raw = {k: "" for k in STANDARD}
    for field, value in _extract_pairs(text):
        if field == "是否吸煙":
            value = _norm_smoking(value)
        elif field == "件數":
            value = _norm_rooms(value)
        elif field in ("入住", "退房", "出生年月日"):
            value = _norm_date(value)
        # 不要讓「空值」覆寫掉先前已解析出的非空值
        if value != "" or raw[field] == "":
            raw[field] = value

    guests = []
    if raw.get("入住者中文") or raw.get("入住者英文") or raw.get("證件號碼"):
        guests.append({
            "cn_name": raw.get("入住者中文", ""),
            "en_name": raw.get("入住者英文", ""),
            "dob": raw.get("出生年月日", ""),
            "idno": raw.get("證件號碼", ""),
        })

    return {
        "飯店": raw.get("飯店", ""),
        "入住": raw.get("入住", ""),
        "退房": raw.get("退房", ""),
        "房型": raw.get("房型", ""),
        "件數": raw.get("件數", ""),
        "備注": raw.get("備注", ""),
        "是否吸煙": raw.get("是否吸煙", ""),
        "guests": guests,
    }
