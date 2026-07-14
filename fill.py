# -*- coding: utf-8 -*-
"""
把一筆訂房資料填進對應的 Excel 模板，回傳填好的 Excel（BytesIO）。
每一筆都會產生一個全新的獨立檔案，不會動到原始模板。
"""
import re
import openpyxl
from openpyxl.styles import Alignment
from io import BytesIO
from datetime import datetime, timezone, timedelta
from config import HOTELS, TEMPLATES_DIR, resolve_hotel

# 台灣時區（GMT+8，無日光節約時間）；Render 伺服器為 UTC，
# 填表日期必須用台灣時間，否則會慢一天。
TAIWAN_TZ = timezone(timedelta(hours=8))

# 簡體/繁體 與常見異體字正規化（讓使用者打的簡稱對到模板正式名）
_CHAR_MAP = {
    "槟": "檳", "双": "雙", "牀": "床", "烟": "煙",
    "达": "達", "台": "臺", "伦": "倫", "汇": "匯",
    "门": "門", "个": "個", "东": "東", "厅": "廳",
    "园": "園",
}
# 房型名中的泛詞（去掉後比對核心字，提升簡稱命中率）
# 注意：不要用會把整個名稱拆光的詞（如「客房」「房」「铁塔」），
#       否則核心變成空字串，空字串是任何字串的子串會全部誤命中。
# 也不放「豪華/尊貴」等前綴詞：否則「梅費爾套房」與「豪華梅費爾套房」
#       會被摺成同一核心字而誤命中（交由第三輪 exact-first 處理）。
_GENERIC = [
    "套房", "大床", "雙床", "雙人床", "人", "典雅",
    "景觀", "金光景", "尊貴", "泳池", "有泳池", "-",
]


def _norm(s):
    s = str(s)
    s = re.sub(r"[（）()]", "", s)  # 去掉全/半形括號（房型註記如「天御别墅（四卧室）」）
    for a, b in _CHAR_MAP.items():
        s = s.replace(a, b)
    return re.sub(r"\s+", "", s).lower()


def _core(s):
    s = _norm(s)
    for g in _GENERIC:
        s = s.replace(g, "")
    return s


# ---------------------------------------------------------------------------
# 訂單摘要 Sheet：固定在產出的 Excel 加一張「訂單摘要」，列出
# 代理 / 訂單編號 / 英文姓名 / 中文姓名 / 入住日期 / 退房日期 / 晚數。
# （與原本的客人清單 Sheet1 互不干擾）
# ---------------------------------------------------------------------------
SUMMARY_SHEET = "訂單摘要"
SUMMARY_HEADERS = ["代理", "訂單編號", "英文姓名", "中文姓名", "入住日期", "退房日期", "晚數"]
_SUMMARY_ALIGN = Alignment(horizontal="center", vertical="center")


def _nights(checkin, checkout):
    """由 入住 / 退房 計算晚數；失敗或無效回傳空字串。"""
    try:
        a = datetime.strptime(_norm_date(checkin), "%Y/%m/%d")
        b = datetime.strptime(_norm_date(checkout), "%Y/%m/%d")
        n = (b - a).days
        return n if n > 0 else ""
    except Exception:
        return ""


def _text_width(s):
    """估算字串顯示寬度：中文/全形字約 2 寬，英文/數字約 1 寬。"""
    if s is None or s == "":
        return 0
    return sum(2 if ord(ch) > 127 else 1 for ch in str(s))


def _fill_summary_sheet(wb, booking, guests, primary):
    if SUMMARY_SHEET in wb.sheetnames:
        ws = wb[SUMMARY_SHEET]
    else:
        ws = wb.create_sheet(SUMMARY_SHEET)

    for c, h in enumerate(SUMMARY_HEADERS, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.alignment = _SUMMARY_ALIGN

    sur, fir = split_en_name(primary.get("en_name", ""))
    en_full = ",".join(x for x in (sur, fir) if x)
    row = [
        booking.get("代理", ""),
        booking.get("訂單編號", ""),
        en_full,
        primary.get("cn_name", ""),
        _norm_date(booking.get("入住", "")),
        _norm_date(booking.get("退房", "")),
        _nights(booking.get("入住", ""), booking.get("退房", "")),
    ]
    for c, v in enumerate(row, start=1):
        cell = ws.cell(row=2, column=c, value=v)
        cell.alignment = _SUMMARY_ALIGN

    # 自動調整欄寬：依標題與資料中最寬者；中文約 2 寬、英文/數字約 1 寬，加 2 安全邊距。
    for c, h in enumerate(SUMMARY_HEADERS, start=1):
        col_letter = openpyxl.utils.get_column_letter(c)
        need = max(_text_width(h), _text_width(row[c - 1])) + 2
        ws.column_dimensions[col_letter].width = max(need, 8)


def _norm_date(value):
    s = re.sub(r"\s+", "", str(value or "")).strip()
    if not s:
        return ""
    m = re.match(r"^(\d{4})[./\-](\d{1,2})[./\-](\d{1,2})$", s)
    if m:
        y, mo, d = m.groups()
        return f"{int(y):04d}/{int(mo):02d}/{int(d):02d}"
    m = re.match(r"^(\d{1,2})月(\d{1,2})日?$", s)
    if m:
        return f"{datetime.now(TAIWAN_TZ).year:04d}/{int(m.group(1)):02d}/{int(m.group(2)):02d}"
    m = re.match(r"^(\d{1,2})[./\-](\d{1,2})$", s)
    if m:
        return f"{datetime.now(TAIWAN_TZ).year:04d}/{int(m.group(1)):02d}/{int(m.group(2)):02d}"
    return s


def split_en_name(en: str):
    """把英文姓名拆成 (姓, 名)。
    範例：'QU,SHENZHONG'  -> ('QU', 'SHENZHONG')   （逗號分隔）
          'TANG/QINGPING' -> ('TANG', 'QINGPING')  （斜線分隔，護照常見格式）
          'ZHOU.YINHUI'   -> ('ZHOU', 'YINHUI')    （點號分隔）
          'SHEN DAN'      -> ('SHEN', 'DAN')        （中文習慣：第一個字是姓）
    """
    en = (en or "").strip()
    if not en:
        return "", ""
    # 逗號 / 斜線 / 點號分隔：前=姓 後=名
    for sep in (",", "/", "."):
        if sep in en:
            a, b = en.split(sep, 1)
            return a.strip(), b.strip()
    toks = en.split()
    if len(toks) >= 2:
        return toks[0], " ".join(toks[1:])
    return "", en


def _set_merged_cell(ws, coord, value):
    """安全寫入儲存格；若 coord 位於合併儲存格內，則改寫該合併範圍左上角那一格。"""
    r, c = openpyxl.utils.coordinate_to_tuple(coord)
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= r <= rng.max_row and rng.min_col <= c <= rng.max_col:
            coord = f"{openpyxl.utils.get_column_letter(rng.min_col)}{rng.min_row}"
            break
    ws[coord] = value


def _write_labeled(ws, coord, value):
    """寫入數值；若該格原本就含標籤文字（如康萊德「人數 Pax(位):」、
    「房數 No.of Rooms (間):」標籤與填空合一），保留標籤並接上數值；
    其他飯店此格為空白填空格，直接寫入數值。"""
    existing = ws[coord].value
    if existing is not None and str(existing).strip():
        txt = str(existing).strip()
        if re.search(r"\d\s*$", txt):
            ws[coord] = txt          # 已含數值就不重複加
        else:
            ws[coord] = f"{txt}{value}"
    else:
        ws[coord] = value


def fill_booking(booking: dict) -> BytesIO:
    """
    booking 結構：
    {
      "飯店": "名匯",
      "入住": "2026/07/20",
      "退房": "2026/07/22",
      "房型": "RK" 或 "豪華大床房",
      "件數": "2",            # 房數
      "備注": "高樓層",
      "是否吸煙": "不吸煙",
      "guests": [
         {"cn_name":"渠慎重", "en_name":"QU,SHENZHONG", "dob":"1961/06/11", "idno":"M41646681"},
         ...
      ]
    }
    """
    hotel_key = resolve_hotel(booking.get("飯店", ""))
    if not hotel_key:
        raise ValueError(f"找不到對應飯店：{booking.get('飯店')}")
    cfg = HOTELS[hotel_key]

    wb = openpyxl.load_workbook(TEMPLATES_DIR / cfg["file"])
    mws = wb[cfg["main_sheet"]]
    mc = cfg["main_cells"]
    guests = booking.get("guests", []) or []
    primary = guests[0] if guests else {}

    sur, fir = split_en_name(primary.get("en_name", ""))
    mws[mc["surname"]] = sur
    mws[mc["firstname"]] = fir
    mws[mc["idno"]] = primary.get("idno", "")
    mws[mc["dob"]] = _norm_date(primary.get("dob", ""))
    mws[mc["checkin"]] = _norm_date(booking.get("入住", ""))
    mws[mc["checkout"]] = _norm_date(booking.get("退房", ""))
    # 房數 / 人數：若該格原本含標籤文字（如康萊德「人數 Pax(位):」、
    # 「房數 No.of Rooms (間):」標籤與填空合一），保留標籤並接上數值；
    # 其他飯店此格為空白填空格，直接寫入數值。
    _write_labeled(mws, mc["rooms"], booking.get("件數", ""))
    _write_labeled(mws, mc["pax"], len(guests))

    # 填表日期（右下角 Date: 後面的底線格；只填日期值，保留模板藍色字體格式）
    if "date" in mc:
        _set_merged_cell(mws, mc["date"], datetime.now(TAIWAN_TZ).strftime('%Y/%m/%d'))

    # 備注 + 吸煙
    remark = (booking.get("備注", "") or "").strip()
    smoking = (booking.get("是否吸煙", "") or "").strip()
    if smoking and hotel_key != "名匯":
        # 名匯有專屬「不吸煙」欄，其他家把吸煙資訊併入備注（只寫「不吸煙」/「吸煙」）
        remark = (remark + f"；{smoking}").strip("；")

    # 寫入備注：若該格原本就含欄位標籤（如「特別要求 Special request :」），
    # 保留標籤並把實際備注接在後面，避免覆蓋掉欄位名稱。
    existing = mws[mc["remark"]].value
    if existing and str(existing).strip():
        mws[mc["remark"]] = f"{existing}{remark}" if remark else existing
    else:
        mws[mc["remark"]] = remark

    # ---- 客人清單（部分飯店無獨立客人清單頁，例如康萊德，直接跳過）----
    if cfg.get("guest_sheet") and cfg.get("guest_cols"):
        gws = wb[cfg["guest_sheet"]]
        gc = cfg["guest_cols"]
        first = cfg["guest_first_row"]
        maxr = gws.max_row

        # 清空舊的範例資料
        for r in range(first, maxr + 1):
            for col in gc.values():
                gws[f"{col}{r}"].value = None

        # 客人數超過現有列數 -> 往下加列
        existing = maxr - first + 1
        needed = max(len(guests), 1)
        if needed > existing:
            gws.insert_rows(maxr + 1, needed - existing)

        for i, g in enumerate(guests):
            r = first + i
            if "cn_name" in gc:
                gws[f"{gc['cn_name']}{r}"] = g.get("cn_name", "")
            if "en_name" in gc:
                gws[f"{gc['en_name']}{r}"] = g.get("en_name", "")
            if "dob" in gc:
                gws[f"{gc['dob']}{r}"] = g.get("dob", "")
            if "idno" in gc:
                gws[f"{gc['idno']}{r}"] = g.get("idno", "")
            if "roomtype" in gc:
                gws[f"{gc['roomtype']}{r}"] = booking.get("房型", "")
            if "smoking" in gc:
                gws[f"{gc['smoking']}{r}"] = smoking

    # ---- 訂單摘要 Sheet（固定新增，與客人清單並存）----
    _fill_summary_sheet(wb, booking, guests, primary)

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


def output_filename(booking: dict) -> str:
    hotel = resolve_hotel(booking.get("飯店", "")) or "訂房"
    g0 = (booking.get("guests") or [{}])[0]
    name = g0.get("cn_name") or g0.get("en_name") or ""
    return f"訂房_{hotel}_{name}.xlsx"
