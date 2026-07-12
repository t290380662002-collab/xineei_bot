# -*- coding: utf-8 -*-
"""
解析使用者貼上的「訂房文字格式」，例如：

    入住：2026/07/20
    退房：2026/07/22
    飯店：名匯
    房型：RK
    件數：2
    備注：高樓層
    是否吸煙：不吸煙
    入住者中文：渠慎重
    入住者英文：QU,SHENZHONG
    出生年月日：1961/06/11
    證件號碼：E12345678

支援全形/半形冒號與 = 號；欄位順序不拘；並內建常見別名。
輸出結構與 fill_booking() 完全相容。
"""
import re

# 標準欄位（對應 fill_booking 的 booking dict key）
STANDARD = [
    "入住", "退房", "飯店", "房型", "件數", "備注", "是否吸煙",
    "入住者中文", "入住者英文", "出生年月日", "證件號碼",
]

# 別名 → 標準欄位（部分使用者可能用不同寫法）
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


def _split_label_value(line):
    """用 全形/半形冒號 或 = 號 拆成 (label, value)。"""
    m = re.split(r"[:：=＝]\s*", line, maxsplit=1)
    if len(m) == 2:
        return m[0].strip(), m[1].strip()
    return None, None


def _match_field(label):
    key = label.strip()
    if key in STANDARD:
        return key
    low = key.lower()
    for alias, std in ALIASES.items():
        if alias.lower() == low:
            return std
    return None


def looks_like_booking(text):
    """是否像一筆訂房文字（至少含 4 個可辨識欄位才當作訂房）。"""
    if not text:
        return False
    found = 0
    for line in text.splitlines():
        label, value = _split_label_value(line)
        if label is None or not value:
            continue
        if _match_field(label):
            found += 1
    return found >= 4


def _norm_smoking(value):
    s = (value or "").strip()
    if any(k in s for k in ("不", "無", "无", "no", "NO", "N", "non")):
        return "不吸煙"
    return "吸煙"


def _norm_rooms(value):
    m = re.search(r"\d+", str(value or ""))
    return m.group(0) if m else (str(value).strip() if value else "")


def parse_booking_text(text):
    """把訂房文字解析成 fill_booking() 相容的 dict。"""
    raw = {k: "" for k in STANDARD}
    for line in text.splitlines():
        label, value = _split_label_value(line)
        if label is None or not value:
            continue
        field = _match_field(label)
        if field is None:
            continue
        if field == "是否吸煙":
            raw[field] = _norm_smoking(value)
        elif field == "件數":
            raw[field] = _norm_rooms(value)
        else:
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
