# -*- coding: utf-8 -*-
"""
證件照片 OCR 與欄位提取 / 比對。

使用 RapidOCR（ONNX Runtime 版，輕量高效，記憶體約 80~120MB）。
中文識別率佳，適合 512MB Render Starter 環境。
首次調用會自動下載 ONNX 模型（約 30MB），之後快取。
"""
import re
import logging
import datetime

logger = logging.getLogger(__name__)

# RapidOCR 引擎單例（懶加載）
_ENGINE = None


def _get_engine():
    """懶加載 RapidOCR 引擎，回傳實例。"""
    global _ENGINE
    if _ENGINE is None:
        from rapidocr_onnxruntime import RapidOCR
        _ENGINE = RapidOCR()
    return _ENGINE


def engine_ready():
    """檢查 OCR 引擎是否可初始化（供健康檢查用）。"""
    try:
        _get_engine()
        return True
    except Exception as e:
        logger.warning("RapidOCR init failed: %s", e)
        return False


# 保留舊名向後相容（bot.py /health 有引用）
paddleocr_ready = engine_ready


# ---------------------------------------------------------------------------
# 卡片區域自動裁剪
# ---------------------------------------------------------------------------
def _detect_and_crop_card(arr):
    """自動偵測最大前景（卡片）區域並裁剪。回傳裁剪後的 numpy array 或 None。"""
    import cv2
    import numpy as np
    H, W = arr.shape[:2]
    gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
    blur = cv2.GaussianBlur(gray, (15, 15), 0)
    _, mask = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    foreground = (mask == 0).astype(np.uint8) * 255
    kernel = np.ones((15, 15), np.uint8)
    closed = cv2.morphologyEx(foreground, cv2.MORPH_CLOSE, kernel)
    contours, _ = cv2.findContours(closed, cv2.RETR_EXTERNAL,
                                   cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return None
    contours = sorted(contours, key=cv2.contourArea, reverse=True)
    x, y, w, h = cv2.boundingRect(contours[0])
    img_area = H * W
    if w * h < img_area * 0.05 or w * h > img_area * 0.95:
        return None
    pad = int(0.05 * max(w, h))
    x1, y1 = max(0, x - pad), max(0, y - pad)
    x2, y2 = min(W, x + w + pad), min(H, y + h + pad)
    return arr[y1:y2, x1:x2]


# ---------------------------------------------------------------------------
# OCR（RapidOCR）
# ---------------------------------------------------------------------------
def ocr_image_bytes(data: bytes) -> str:
    """對圖片 bytes 做 OCR（RapidOCR），回傳辨識文字（每行用換行分開）。"""
    from io import BytesIO
    from PIL import Image
    import numpy as np

    try:
        img = Image.open(BytesIO(data)).convert("RGB")
    except Exception as e:
        raise ValueError(f"圖片無法開啟：{e}")

    arr = np.array(img)

    # 自動裁剪到卡片區域（減少背景雜訊）
    cropped = _detect_and_crop_card(arr)
    if cropped is not None and cropped.size > 0:
        arr = cropped

    engine = _get_engine()
    result, _ = engine(arr)  # RapidOCR 回 (list, elapse)

    if not result:
        return ""

    lines = []
    for item in result:
        # RapidOCR 格式：[bbox, text, confidence]
        if isinstance(item, (list, tuple)) and len(item) >= 2:
            text = str(item[1]).strip()
            if text:
                lines.append(text)

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
    m = re.search(r"(?:中文)?姓名[：: ]*([一-鿿 ]{2,8})", text)
    if m:
        return _norm_cn(m.group(1))
    for line in text.splitlines():
        line = _norm_cn(line)
        if 2 <= len(line) <= 5 and re.fullmatch(r"[一-鿿]+", line):
            return line
    return ""


def _extract_en_name(text):
    m = re.search(r"(?:英文)?姓名[：: ]*([A-Z][A-Z ,.\-]{2,40})", text)
    if m:
        return m.group(1).strip()
    m2 = re.search(r"\b([A-Z]{2,})\s*,\s*([A-Z]{2,})\b", text)
    if m2:
        return f"{m2.group(1)},{m2.group(2)}"
    return ""


def _extract_idno(text):
    cands = re.findall(r"\b[A-Z]{1,2}[0-9]{6,9}\b", text)
    for c in cands:
        if re.fullmatch(r"(?:19|20)\d{6}", c):
            continue
        return c
    return ""


def _extract_dob(text):
    m = re.search(r"出生[年月日]*[：: ]*(\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2})", text)
    if m:
        return _norm_date(m.group(1))
    m2 = re.search(r"出生[年月日]*[：: ]*(\d{8})", text)
    if m2:
        return _norm_date(m2.group(1))
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

    if b_cn and o_cn and b_cn != o_cn:
        warns.append(f"⚠️ 中文姓名不符：填寫「{b_cn}」／證件「{o_cn}」")
    if b_en and o_en:
        if _en_norm(b_en) != _en_norm(o_en):
            if b_cn:
                ok, _ = verify_name_match(b_cn, o_en)
                if not ok:
                    warns.append(f"⚠️ 英文姓名不符：填寫「{b_en}」／證件「{o_en}」")
            else:
                warns.append(f"⚠️ 英文姓名不符：填寫「{b_en}」／證件「{o_en}」")
    if b_dob and o_dob and b_dob != o_dob:
        warns.append(f"⚠️ 出生日期不符：填寫「{b_dob}」／證件「{o_dob}」")
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
