# -*- coding: utf-8 -*-
"""
證件照片 OCR 與欄位提取 / 比對。

使用本地 Tesseract（證件照片不出伺服器，隱私最佳）。
限制：
  - Tesseract 必須已安裝（含 chi_tra / chi_sim 中文包），否則 ocr_image_bytes
    會拋出 TesseractNotInstalled。Docker 部署已內建安裝。
  - OCR 準確率取決於拍照清晰度，約 85–95%；下方比對邏輯已做容錯（去空白、
    日期歸一化、拼音核對），但仍建議人工確認。
"""
import re
import logging
import datetime

logger = logging.getLogger(__name__)


class TesseractNotInstalled(Exception):
    """Tesseract 未安裝或不在 PATH 時拋出。"""


def tesseract_ready():
    """檢查本機 Tesseract 是否可用（供健康檢查用）。"""
    try:
        import pytesseract
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# 低階 OCR（本機 Tesseract，含 chi_tra/chi_sim/eng）
# ---------------------------------------------------------------------------
def _detect_and_crop_card(arr):
    """自動偵測最大前景（卡片）區域並裁剪。回傳裁剪後的 numpy array 或 None。"""
    import cv2
    import numpy as np
    H, W = arr.shape[:2]
    gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
    # 大模糊 + Otsu 找前景（深色 = 卡片，淺色 = 木桌/牆壁）
    blur = cv2.GaussianBlur(gray, (15, 15), 0)
    _, mask = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    foreground = (mask == 0).astype(np.uint8) * 255
    # 形態學閉運算連接斷裂
    kernel = np.ones((15, 15), np.uint8)
    closed = cv2.morphologyEx(foreground, cv2.MORPH_CLOSE, kernel)
    contours, _ = cv2.findContours(closed, cv2.RETR_EXTERNAL,
                                   cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return None
    contours = sorted(contours, key=cv2.contourArea, reverse=True)
    x, y, w, h = cv2.boundingRect(contours[0])
    # 過濾太小或太大
    img_area = H * W
    if w * h < img_area * 0.05 or w * h > img_area * 0.95:
        return None
    # 加 5% padding
    pad = int(0.05 * max(w, h))
    x1, y1 = max(0, x - pad), max(0, y - pad)
    x2, y2 = min(W, x + w + pad), min(H, y + h + pad)
    return arr[y1:y2, x1:x2]


def _ocr_attempt(img, lang: str, psm: int = 6) -> str:
    """對 PIL Image 跑一次 OCR，回傳文字。"""
    import pytesseract
    return pytesseract.image_to_string(
        img, lang=lang,
        config=f"--psm {psm} --oem 3",
    )


def ocr_image_bytes(data: bytes) -> str:
    """對圖片 bytes 做 OCR，回傳辨識文字（含換行）。

    預處理流程（依測試最佳）：
      1. 自動偵測並裁剪卡片區域（用 Otsu 找最大前景）
      2. 灰階
      3. 2x 放大（小字才抓得到）
      4. 不做二值化（adaptive threshold 對中文證件照會打散字形）
      5. 跑多組 (lang, psm) 嘗試，合併結果
    """
    from io import BytesIO
    from PIL import Image, ImageOps
    import numpy as np
    import pytesseract

    try:
        img = Image.open(BytesIO(data))
    except Exception as e:
        raise ValueError(f"圖片無法開啟：{e}")

    img = img.convert("RGB")
    arr = np.array(img)

    # 自動裁剪到卡片
    cropped = _detect_and_crop_card(arr)
    if cropped is not None and cropped.size > 0:
        arr = cropped

    # 灰階 + 2x 放大
    img = Image.fromarray(arr).convert("L")
    w, h = img.size
    # 若裁剪後仍很大，不放大；否則 2x
    if max(w, h) < 1500:
        img = img.resize((w * 2, h * 2), Image.LANCZOS)

    # 多組嘗試，合併結果（單一語言包缺失時不中斷，容錯繼續下一組）
    chunks = []
    attempts = [
        ("chi_tra+eng", 6),   # 繁中 + 英文，PSM 6 單一區塊
        ("chi_tra+eng", 11),  # PSM 11 對稀疏文字（姓名行）有時更好
        ("eng", 6),           # 純英文，抓 ID/MRZ
    ]
    last_err = None
    for lang, psm in attempts:
        try:
            chunks.append(_ocr_attempt(img, lang, psm=psm))
        except Exception as e:
            # 單一 lang/psm 失敗不影響其他嘗試（可能是語言包缺失）
            logger.debug("OCR attempt lang=%s psm=%s failed: %s", lang, psm, e)
            last_err = e
            continue

    if not chunks:
        # 全部嘗試都失敗
        if last_err and ("Tesseract" in type(last_err).__name__
                          or "tesseract" in str(last_err).lower()):
            raise TesseractNotInstalled(
                "伺服器未安裝 Tesseract，無法辨識證件照片。")
        raise last_err if last_err else RuntimeError("OCR 全部嘗試失敗")

    # 合併：用換行連接，extract_fields 內部 regex 會自行挑
    text = "\n".join(chunks)
    return text


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
