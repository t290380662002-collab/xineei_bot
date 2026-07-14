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
    "代理", "訂單編號",
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
    "代理": "代理", "agent": "代理", "代理商": "代理", "渠道": "代理",
    "訂單編號": "訂單編號", "訂單": "訂單編號", "單號": "訂單編號",
    "order": "訂單編號", "orderno": "訂單編號", "order no": "訂單編號",
}

# 所有可用作「標籤」的字串（標準欄位 + 別名），用於整段掃描
_ALL = sorted(set(STANDARD + list(ALIASES.keys())), key=len, reverse=True)
_LABEL_ALT = "|".join(re.escape(k) for k in _ALL)
# 單行掃描：捕捉 標籤[:：=]值，值到「同行下一個標籤」或「行尾」為止。
# 注意：不跨行，避免最後一個欄位把文末多餘文字整段吞進去、污染欄位（如證件號碼）。
_PAIR_RE = re.compile(
    r"(?P<lab>" + _LABEL_ALT + r")[:：=＝]\s*(?P<val>.*?)(?=(?P<lab2>" + _LABEL_ALT + r")[:：=＝]|\n|\r|\Z)",
    re.MULTILINE,
)
# 判斷一整行是否「純標籤字」（沒有冒號、沒有值，例如單獨一行的「备注」）
_BARE_LABEL_RE = re.compile(r"^\s*(?P<lab>" + _LABEL_ALT + r")\s*$")

# 可容納「換行續行」的欄位（值可能被使用者換行拆開，需把續行併回）：
#   房型 / 飯店 / 姓名 直接接續；備注 用「；」分隔。
_WRAPPABLE = {"房型", "飯店", "入住者中文", "入住者英文", "備注"}

# 訂房文字中常出現的「代理」代碼（也是過去會被從備注剔除的 SS/AT/WW/MM 等）。
# 現在改為辨識成「代理」欄位，不再當作備注雜訊剔除。
AGENT_CODES = {"AT", "SS", "私域", "WW", "MM", "M", "ALEN"}

# 備注黑名單（已無需剔除的代理代碼，此處留空；如需剔除其他雜訊再加回）。
SKIP_TOKENS = set()
_SKIP_RE = re.compile(
    r"(?<![A-Za-z])(" + "|".join(re.escape(t) for t in SKIP_TOKENS) + r")(?![A-Za-z])",
    re.IGNORECASE,
)


def _match_agent(seg):
    """若整段文字就是一個代理代碼（AT/SS/私域/WW/MM/M/ALEN），回傳該代碼，否則 None。"""
    if not seg:
        return None
    s = re.sub(r"\s+", "", str(seg)).upper()
    for code in AGENT_CODES:
        if s == code.upper():
            return code
    return None


def _clean_remark(s):
    """移除備注中的黑名單代碼（目前無），並正規化空白與分隔符。"""
    if SKIP_TOKENS:
        s = _SKIP_RE.sub("", s or "")
    s = re.sub(r"[\s；;]+", "；", s).strip("；; ").strip()
    return s



def _match_field(label):
    key = label.strip()
    if key in STANDARD:
        return key
    low = key.lower()
    for alias, std in ALIASES.items():
        if alias.lower() == low:
            return std
    return None


def _append_value(pairs, field, text, sep=""):
    """把續行文字併到最近一筆同欄位的值後面；若無則新增一筆。"""
    for i in range(len(pairs) - 1, -1, -1):
        if pairs[i][0] == field:
            old = pairs[i][1]
            joined = (old + (sep if old else "") + text).strip(sep) if sep else (old + text)
            pairs[i] = (field, joined.strip())
            return
    pairs.append((field, text.strip()))


def _extract_pairs(text):
    """逐行解析，回傳 (pairs, extra_lines)。
    - pairs = [(field, value), ...]（依出現順序）
    - extra_lines = 未歸屬任何欄位、且非純標籤字的額外文字（如 AT），供併入備注
    續行處理：沒有標籤的行，若前一欄位可容納換行（_WRAPPABLE）則併回該欄位；
    否則視為額外文字。純標籤字（如單獨的「备注」）會開啟該欄位以承接下一行的值。"""
    pairs = []
    extra = []
    last_field = None  # 最近一個可承接續行的欄位
    for line in (text or "").splitlines():
        if not line.strip():
            last_field = None  # 空行切斷續行
            continue
        matches = [m for m in _PAIR_RE.finditer(line) if _match_field(m.group("lab"))]
        if matches:
            for m in matches:
                field = _match_field(m.group("lab"))
                pairs.append((field, m.group("val").strip()))
                last_field = field
            continue
        # 整行是純標籤字（如單獨的「备注」）→ 開啟該欄位承接下一行，本行不留值
        bare = _BARE_LABEL_RE.match(line)
        if bare:
            f = _match_field(bare.group("lab"))
            last_field = f if f in _WRAPPABLE else None
            continue
        # 一般續行文字
        seg = line.strip()
        # 代理代碼獨立成行（AT/SS/私域/WW/MM/ALEN）→ 視為「代理」欄位。
        # 優先於「續行併回上一欄位」判斷，避免被誤併進 入住者英文/飯店 等欄位。
        agent = _match_agent(seg)
        if agent and not any(f == "代理" for f, _ in pairs):
            pairs.append(("代理", agent))
            continue
        if last_field in _WRAPPABLE:
            if last_field == "備注":
                cleaned = _clean_remark(seg)
                if cleaned:
                    _append_value(pairs, last_field, cleaned, sep="；")
            else:
                _append_value(pairs, last_field, seg, sep="")
        else:
            cleaned = _clean_remark(seg)
            if cleaned:
                extra.append(cleaned)
    return pairs, extra


def looks_like_booking(text):
    """是否像一筆訂房文字（至少含 4 個可辨識欄位才當作訂房）。"""
    if not text:
        return False
    pairs, _ = _extract_pairs(text)
    found = sum(1 for f, _ in pairs)
    return found >= 4


def _norm_smoking(value):
    s = (value or "").strip()
    if any(k in s for k in ("不", "無", "无", "否", "没", "沒", "no", "NO", "N", "non")):
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
    pairs, extra_lines = _extract_pairs(text)
    for field, value in pairs:
        if field == "是否吸煙":
            value = _norm_smoking(value)
        elif field == "件數":
            value = _norm_rooms(value)
        elif field in ("入住", "退房", "出生年月日"):
            value = _norm_date(value)
        elif field == "備注":
            value = _clean_remark(value)
        # 不要讓「空值」覆寫掉先前已解析出的非空值
        if value != "" or raw[field] == "":
            raw[field] = value

    # 未歸屬任何欄位、且非純標籤字的額外文字（如 AT），併入備注，避免被悄悄丟掉。
    if extra_lines:
        extra = "；".join(extra_lines)
        raw["備注"] = (raw["備注"] + "；" + extra).strip("；") if raw["備注"] else extra

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
        "代理": raw.get("代理", ""),
        "訂單編號": raw.get("訂單編號", ""),
        "guests": guests,
    }
