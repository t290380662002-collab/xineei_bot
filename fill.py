# -*- coding: utf-8 -*-
"""
把一筆訂房資料填進對應的 Excel 模板，回傳填好的 Excel（BytesIO）。
每一筆都會產生一個全新的獨立檔案，不會動到原始模板。
"""
import re
import openpyxl
from io import BytesIO
from datetime import datetime
from config import HOTELS, TEMPLATES_DIR, resolve_hotel

# 簡體/繁體 與常見異體字正規化（讓使用者打的簡稱對到模板正式名）
_CHAR_MAP = {
    "槟": "檳", "双": "雙", "牀": "床", "烟": "煙",
    "达": "達", "台": "臺",
}
# 房型名中的泛詞（去掉後比對核心字，提升簡稱命中率）
# 注意：不要用會把整個名稱拆光的詞（如「客房」「房」「铁塔」），
#       否則核心變成空字串，空字串是任何字串的子串會全部誤命中。
_GENERIC = [
    "套房", "大床", "雙床", "雙人床", "人", "典雅",
    "景觀", "金光景", "尊貴", "豪華", "泳池", "有泳池", "-",
]


def _norm(s):
    s = str(s)
    for a, b in _CHAR_MAP.items():
        s = s.replace(a, b)
    return re.sub(r"\s+", "", s).lower()


def _core(s):
    s = _norm(s)
    for g in _GENERIC:
        s = s.replace(g, "")
    return s


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
        return f"{datetime.now().year:04d}/{int(m.group(1)):02d}/{int(m.group(2)):02d}"
    m = re.match(r"^(\d{1,2})[./\-](\d{1,2})$", s)
    if m:
        return f"{datetime.now().year:04d}/{int(m.group(1)):02d}/{int(m.group(2)):02d}"
    return s


def find_room_cell(hotel_cfg, room_input):
    """根據代碼或中文名（含簡稱/異體字），找出要打勾的方框儲存格。"""
    if not room_input:
        return None
    room_input = re.sub(r"[（）()]", "", str(room_input))  # 去掉括號註記
    inp = _norm(room_input)
    # 第一輪：用代碼比對（最精準）
    for cell, code, cn in hotel_cfg["room_types"]:
        c = _norm(code)
        if c == inp or c in inp or inp in c:
            return cell
    # 第二輪：用中文名「核心字」比對（去泛詞 + 異體字正規化）
    inpc = _core(room_input)
    if inpc:
        for cell, code, cn in hotel_cfg["room_types"]:
            if not cn:
                continue
            ccore = _core(cn)
            if not ccore:
                continue
            if inpc in ccore or ccore in inpc:
                return cell
    # 第三輪：床型關鍵字容錯（双床/雙人床/twin → 雙床套房；大床/king → 大床套房）
    bed_rules = [
        (("雙床", "雙人床", "twin", "二人"), ("雙床", "雙人床", "twin")),
        (("大床", "king", "特大床"), ("大床", "king")),
    ]
    for keys, hits in bed_rules:
        if any(k in inp for k in keys):
            for cell, code, cn in hotel_cfg["room_types"]:
                cn_n = _norm(cn or "")
                if any(h in cn_n for h in hits):
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
    mws[mc["dob"]] = _norm_date(primary.get("dob", ""))
    mws[mc["checkin"]] = _norm_date(booking.get("入住", ""))
    mws[mc["checkout"]] = _norm_date(booking.get("退房", ""))
    mws[mc["rooms"]] = booking.get("件數", "")
    mws[mc["pax"]] = len(guests)

    # 備注 + 吸煙
    remark = (booking.get("備注", "") or "").strip()
    smoking = (booking.get("是否吸煙", "") or "").strip()
    if smoking and hotel_key != "名匯":
        # 名匯有專屬「不吸煙」欄，其他家把吸煙資訊併入備注
        remark = (remark + f"；吸煙狀態：{smoking}").strip("；")
    mws[mc["remark"]] = remark

    # 房型方框打勾（支援 / 、、, 分隔的多房型同時打勾）
    room_raw = booking.get("房型", "")
    segs = [s.strip("（）() ") for s in re.split(r"[/、,，;；]", room_raw) if s.strip()]
    matched = False
    for seg in segs:
        cell = find_room_cell(cfg, seg)
        if cell:
            check_box(mws, cell)
            matched = True
    if not matched and room_raw:
        # 找不到對應房型，記到備注提醒
        mws[mc["remark"]] = (remark + f"；[房型未對應：{room_raw}]").strip("；")

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
