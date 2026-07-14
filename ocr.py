# -*- coding: utf-8 -*-
"""
證件照片 OCR 與欄位提取 / 比對。

使用 rapidocr_onnxruntime（純 pip 安裝，自帶中英文 ONNX 模型，無需系統安裝 Tesseract）。
限制：
  - 首次執行會自動下載 ONNX 模型（需網路，約數十 MB，之後快取）。
  - OCR 準確率取決於拍照清晰度，約 85–95%；下方比對邏輯已做容錯（去空白、
    日期歸一化、拼音核對），但仍建議人工確認。
"""
import re
import logging
import datetime

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# 低階 OCR（rapidocr_onnxruntime，純 pip，含中英文模型）
# ---------------------------------------------------------------------------
_ENGINE = None


def _get_engine():
    """懶加載並快取 RapidOCR 引擎（首次會下載模型）。"""
    global _ENGINE
    if _ENGINE is None:
        from rapidocr_onnxruntime import RapidOCR
        _ENGINE = RapidOCR()
    return _ENGINE


def ocr_image_bytes(data: bytes) -> str:
    """對圖片 bytes 做 OCR，回傳辨識文字（每行一筆）。"""
    engine = _get_engine()
    try:
        result, _ = engine(data)
    except Exception as e:
        raise RuntimeError(f"證件照片辨識失敗：{e}")
    # result: list of [box, text, score] 或 None
    lines = [item[1] for item in (result or [])]
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# 欄位提取（從 OCR 文本）
# ---------------------------------------------------------------------------
def _norm_cn(s):
    """中文去所有空白。"""
    return re.sub(r"\s+", "", (s or ""))


def _norm_date(s):
    """把 1981.08.26 / 1981-08-26 / 19810826 歸一為 1981.08.26；無法解析回 None。"""
    s = (s or "").strip()
    m = re.search(r"(\d{4})[.\-/年](\d{1,2})[.\-/月](\d{1,2})", s)
    if m:
        y, mo, d = m.groups()
        try:
            return f"{int(y):04d}.{int(mo):02d}.{int(d):02d}"
        except Exception:
            return None
    m2 = re.search(r"(?:19|20)\d{2}[01]\d[0123]\d", s)
    if m2:
        v = m2.group(0)
        return f"{v[0:4]}.{v[4:6]}.{v[6:8]}"
    return None


def _extract_cn_name(text):
    # 優先抓「（中文）姓名」標籤後的中文（容許中間空白，如「陈 翠 翠」）
    m = re.search(r"(?:中文)?姓名[：: ]*([一-鿿 ]{2,8})", text)
    if m:
        return _norm_cn(m.group(1))
    # 退而求其次：任一行純中文 2–5 字
    for line in text.splitlines():
        line = _norm_cn(line)
        if 2 <= len(line) <= 5 and re.fullmatch(r"[一-鿿]+", line):
            return line
    return ""


def _extract_en_name(text):
    # 優先抓「（英文）姓名」標籤後的大寫字母序列
    m = re.search(r"(?:英文)?姓名[：: ]*([A-Z][A-Z ,.\-]{2,40})", text)
    if m:
        return m.group(1).strip()
    # 退而求其次：找 字母,字母 形式（如 CHEN, CUICUI）
    m2 = re.search(r"\b([A-Z]{2,})\s*,\s*([A-Z]{2,})\b", text)
    if m2:
        return f"{m2.group(1)},{m2.group(2)}"
    return ""


def _extract_idno(text):
    # 證件號：1–2 字母 + 6–9 數字（排除純日期型 19xxxxxx）
    cands = re.findall(r"\b[A-Z]{1,2}[0-9]{6,9}\b", text)
    for c in cands:
        if re.fullmatch(r"(?:19|20)\d{6}", c):
            continue
        return c
    return ""


def _extract_dob(text):
    # 優先抓「出生」標籤附近的日期
    m = re.search(r"出生[年月日]*[：: ]*(\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2})", text)
    if m:
        return _norm_date(m.group(1))
    m2 = re.search(r"出生[年月日]*[：: ]*(\d{8})", text)
    if m2:
        return _norm_date(m2.group(1))
    # 否則取所有日期中「年份距今 1 年以上、且在 1900 之後」的第一個（通常即出生）
    dates = re.findall(r"(\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2})", text)
    year = datetime.date.today().year
    for d in dates:
        nd = _norm_date(d)
        if not nd:
            continue
        y = int(nd[:4])
        if 1900 <= y <= year - 1:
            return nd
    return ""


def extract_fields(text: str) -> dict:
    """從 OCR 文本提取四欄位，回傳 {cn_name, en_name, dob, idno}。"""
    return {
        "cn_name": _extract_cn_name(text),
        "en_name": _extract_en_name(text),
        "dob": _extract_dob(text),
        "idno": _extract_idno(text),
    }


# ---------------------------------------------------------------------------
# 比對 / 回顯
# ---------------------------------------------------------------------------
def _en_norm(s):
    return re.sub(r"[^A-Z]", "", (s or "").upper())


def verify_ocr_vs_booking(ocr: dict, booking: dict) -> list:
    """比對 OCR 欄位 與 文字 booking 欄位，回傳不一致提示清單（空=完全相符）。"""
    from fill import verify_name_match

    warns = []
    g0 = (booking.get("guests") or [{}])[0]
    b_cn = _norm_cn(g0.get("cn_name", ""))
    b_en = g0.get("en_name", "")
    b_dob = _norm_date(g0.get("dob", ""))
    b_id = (g0.get("idno", "") or "").strip()

    o_cn = _norm_cn(ocr.get("cn_name", ""))
    o_en = ocr.get("en_name", "")
    o_dob = _norm_date(ocr.get("dob", ""))
    o_id = (ocr.get("idno", "") or "").strip()

    # 中文姓名（去空白後比對）
    if b_cn and o_cn and b_cn != o_cn:
        warns.append(f"⚠️ 中文姓名不符：填寫「{b_cn}」／證件「{o_cn}」")

    # 英文姓名：先比對去分隔符後的字母；不同再用拼音核對（容許姓在後等）
    if b_en and o_en:
        if _en_norm(b_en) != _en_norm(o_en):
            if b_cn:
                ok, _ = verify_name_match(b_cn, o_en)
                if not ok:
                    warns.append(f"⚠️ 英文姓名不符：填寫「{b_en}」／證件「{o_en}」")
            else:
                warns.append(f"⚠️ 英文姓名不符：填寫「{b_en}」／證件「{o_en}」")

    # 出生日期（歸一化後比對）
    if b_dob and o_dob and b_dob != o_dob:
        warns.append(f"⚠️ 出生日期不符：填寫「{b_dob}」／證件「{o_dob}」")

    # 證件號碼（不分大小寫）
    if b_id and o_id and b_id.upper() != o_id.upper():
        warns.append(f"⚠️ 證件號碼不符：填寫「{b_id}」／證件「{o_id}」")

    return warns


def format_fields(ocr: dict) -> str:
    """把識別結果排版成回顯文字。"""
    cn = ocr.get("cn_name") or "（未辨識）"
    en = ocr.get("en_name") or "（未辨識）"
    dob = ocr.get("dob") or "（未辨識）"
    idno = ocr.get("idno") or "（未辨識）"
    return (
        "📄 證件識別結果（請人工確認準確性）：\n"
        f"中文姓名：{cn}\n"
        f"英文姓名：{en}\n"
        f"出生日期：{dob}\n"
        f"證件號碼：{idno}"
    )
