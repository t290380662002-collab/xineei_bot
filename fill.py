# -*- coding: utf-8 -*-
"""
把一筆訂房資料填進對應的 Excel 模板，回傳填好的 Excel（BytesIO）。
每一筆都會產生一個全新的獨立檔案，不會動到原始模板。
"""
import re
import openpyxl
from io import BytesIO
from config import HOTELS, TEMPLATES_DIR, resolve_hotel


def _norm(s):
    return re.sub(r"\s+", "", str(s)).lower()


def find_room_cell(hotel_cfg, room_input):
    """根據代碼或中文名，找出要打勾的方框儲存格。"""
    if not room_input:
        return None
    inp = _norm(room_input)
    # 第一輪：用代碼比對（最精準）
    for cell, code, cn in hotel_cfg["room_types"]:
        c = _norm(code)
        if c == inp or c in inp or inp in c:
            return cell
    # 第二輪：用中文名比對
    for cell, code, cn in hotel_cfg["room_types"]:
        if cn and (cn in room_input or room_input in cn):
            return cell
    return None


def check_box(ws, cell_ref):
    """把 '(   )' 方框打勾成 '(✓)'。"""
    c = ws[cell_ref]
    if c.value is None:
        c.value = "(✓)"
        return
    txt = str(c.value)
    if "(✓)" in txt or "(X)" in txt:
        return
    c.value = re.sub(r"\([\s]*\)", "(✓)", txt, count=1)


def split_en_name(en: str):
    """把英文姓名拆成 (姓, 名)。
    範例：'QU,SHENZHONG' -> ('QU', 'SHENZHONG')
          'SHEN DAN'     -> ('SHEN', 'DAN')  （中文習慣：第一個字是姓）
    """
    en = (en or "").strip()
    if not en:
        return "", ""
    if "," in en:
        a, b = en.split(",", 1)
        return a.strip(), b.strip()
    toks = en.split()
    if len(toks) >= 2:
        return toks[0], " ".join(toks[1:])
    return "", en


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
    mws[mc["dob"]] = primary.get("dob", "")
    mws[mc["checkin"]] = booking.get("入住", "")
    mws[mc["checkout"]] = booking.get("退房", "")
    mws[mc["rooms"]] = booking.get("件數", "")
    mws[mc["pax"]] = len(guests)

    # 備注 + 吸煙
    remark = (booking.get("備注", "") or "").strip()
    smoking = (booking.get("是否吸煙", "") or "").strip()
    if smoking and hotel_key != "名匯":
        # 名匯有專屬「不吸煙」欄，其他家把吸煙資訊併入備注
        remark = (remark + f"；吸煙狀態：{smoking}").strip("；")
    mws[mc["remark"]] = remark

    # 房型方框打勾
    room_cell = find_room_cell(cfg, booking.get("房型", ""))
    if room_cell:
        check_box(mws, room_cell)
    else:
        # 找不到對應房型，記到備注提醒
        rt = booking.get("房型", "")
        if rt:
            mws[mc["remark"]] = (remark + f"；[房型未對應：{rt}]").strip("；")

    # ---- 客人清單 ----
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

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


def output_filename(booking: dict) -> str:
    hotel = resolve_hotel(booking.get("飯店", "")) or "訂房"
    g0 = (booking.get("guests") or [{}])[0]
    name = g0.get("cn_name") or g0.get("en_name") or ""
    return f"訂房_{hotel}_{name}.xlsx"
