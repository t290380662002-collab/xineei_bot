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

    sur, fir = split_en_name(primary.get("en_name", ""), primary.get("cn_name", ""))
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


# 常見中文姓氏 → 護照拼音（大寫）。多字姓放前面，優先匹配。
_SURNAMES_PINYIN = {
    # 複姓
    "歐陽": "OUYANG", "欧阳": "OUYANG", "太史": "TAISHI", "上官": "SHANGGUAN",
    "東方": "DONGFANG", "东方": "DONGFANG", "諸葛": "ZHUGE", "诸葛": "ZHUGE",
    "司馬": "SIMA", "司马": "SIMA", "皇甫": "HUANGFU", "尉遲": "YUCHI",
    "公孙": "GONGSUN", "公孫": "GONGSUN", "慕容": "MURONG", "司徒": "SITU",
    "司空": "SIKONG", "令狐": "LINGHU", "軒轅": "XUANYUAN", "轩辕": "XUANYUAN",
    "南宮": "NANGONG", "南宫": "NANGONG", "夏侯": "XIAHOU", "聞人": "WENREN",
    "鲜于": "XIANYU", "鮮于": "XIANYU", "贺兰": "HELAN", "賀蘭": "HELAN",
    "宇文": "YUWEN", "呼延": "HUYAN", "西門": "XIMEN", "西门": "XIMEN",
    "東郭": "DONGGUO", "东郭": "DONGGUO", "南門": "NANMEN", "南门": "NANMEN",
    "百里": "BAILI", "完顏": "WANYAN", "完颜": "WANYAN", "獨孤": "DUGU", "独孤": "DUGU",
    # 單姓（依頻率與常用度排序，不完整但覆蓋絕大多數酒店訂單）
    "趙": "ZHAO", "赵": "ZHAO", "錢": "QIAN", "钱": "QIAN", "孫": "SUN", "孙": "SUN",
    "李": "LI", "周": "ZHOU", "吳": "WU", "吴": "WU", "鄭": "ZHENG", "郑": "ZHENG",
    "王": "WANG", "馮": "FENG", "冯": "FENG", "陳": "CHEN", "陈": "CHEN", "褚": "CHU",
    "衛": "WEI", "卫": "WEI", "蔣": "JIANG", "蒋": "JIANG", "沈": "SHEN", "韓": "HAN",
    "韩": "HAN", "楊": "YANG", "杨": "YANG", "朱": "ZHU", "秦": "QIN", "尤": "YOU",
    "許": "XU", "许": "XU", "何": "HE", "呂": "LYU", "吕": "LYU", "施": "SHI", "張": "ZHANG",
    "张": "ZHANG", "孔": "KONG", "曹": "CAO", "嚴": "YAN", "严": "YAN", "華": "HUA",
    "华": "HUA", "金": "JIN", "魏": "WEI", "陶": "TAO", "姜": "JIANG", "戚": "QI",
    "謝": "XIE", "谢": "XIE", "鄒": "ZOU", "邹": "ZOU", "喻": "YU", "柏": "BAI",
    "水": "SHUI", "竇": "DOU", "窦": "DOU", "章": "ZHANG", "雲": "YUN", "云": "YUN",
    "蘇": "SU", "苏": "SU", "潘": "PAN", "葛": "GE", "奚": "XI", "范": "FAN",
    "彭": "PENG", "郎": "LANG", "魯": "LU", "鲁": "LU", "韋": "WEI", "韦": "WEI",
    "昌": "CHANG", "馬": "MA", "马": "MA", "苗": "MIAO", "鳳": "FENG", "凤": "FENG",
    "花": "HUA", "方": "FANG", "俞": "YU", "任": "REN", "袁": "YUAN", "柳": "LIU",
    "鮑": "BAO", "鲍": "BAO", "史": "SHI", "唐": "TANG", "費": "FEI", "费": "FEI",
    "廉": "LIAN", "岑": "CEN", "薛": "XUE", "雷": "LEI", "賀": "HE", "贺": "HE",
    "倪": "NI", "湯": "TANG", "汤": "TANG", "滕": "TENG", "殷": "YIN", "羅": "LUO",
    "罗": "LUO", "畢": "BI", "毕": "BI", "郝": "HAO", "鄔": "WU", "邬": "WU",
    "安": "AN", "常": "CHANG", "樂": "LE", "乐": "LE", "于": "YU", "時": "SHI",
    "时": "SHI", "傅": "FU", "付": "FU", "皮": "PI", "卞": "BIAN", "齊": "QI",
    "齐": "QI", "康": "KANG", "伍": "WU", "余": "YU", "元": "YUAN", "卜": "BU",
    "顧": "GU", "顾": "GU", "孟": "MENG", "平": "PING", "黃": "HUANG", "黄": "HUANG",
    "和": "HE", "穆": "MU", "蕭": "XIAO", "萧": "XIAO", "尹": "YIN", "姚": "YAO",
    "邵": "SHAO", "湛": "ZHAN", "汪": "WANG", "祁": "QI", "毛": "MAO", "禹": "YU",
    "狄": "DI", "米": "MI", "貝": "BEI", "贝": "BEI", "明": "MING", "臧": "ZANG",
    "計": "JI", "计": "JI", "伏": "FU", "成": "CHENG", "戴": "DAI", "談": "TAN",
    "谈": "TAN", "宋": "SONG", "茅": "MAO", "龐": "PANG", "庞": "PANG", "熊": "XIONG",
    "紀": "JI", "纪": "JI", "舒": "SHU", "屈": "QU", "項": "XIANG", "项": "XIANG",
    "祝": "ZHU", "董": "DONG", "梁": "LIANG", "杜": "DU", "阮": "RUAN", "藍": "LAN",
    "蓝": "LAN", "閔": "MIN", "闵": "MIN", "席": "XI", "季": "JI", "麻": "MA",
    "強": "QIANG", "强": "QIANG", "賈": "JIA", "贾": "JIA", "路": "LU", "婁": "LOU",
    "娄": "LOU", "危": "WEI", "江": "JIANG", "童": "TONG", "顏": "YAN", "颜": "YAN",
    "郭": "GUO", "梅": "MEI", "盛": "SHENG", "林": "LIN", "刁": "DIAO", "鍾": "ZHONG",
    "钟": "ZHONG", "徐": "XU", "邱": "QIU", "駱": "LUO", "骆": "LUO", "高": "GAO",
    "夏": "XIA", "蔡": "CAI", "田": "TIAN", "樊": "FAN", "胡": "HU", "凌": "LING",
    "霍": "HUO", "虞": "YU", "萬": "WAN", "万": "WAN", "支": "ZHI", "柯": "KE",
    "管": "GUAN", "盧": "LU", "卢": "LU", "莫": "MO", "房": "FANG", "裘": "QIU",
    "繆": "MIAO", "缪": "MIAO", "干": "GAN", "應": "YING", "应": "YING", "宗": "ZONG",
    "丁": "DING", "宣": "XUAN", "賁": "BEN", "贲": "BEN", "鄧": "DENG", "邓": "DENG",
    "郁": "YU", "單": "SHAN", "单": "SHAN", "杭": "HANG", "洪": "HONG", "包": "BAO",
    "石": "SHI", "崔": "CUI", "吉": "JI", "鈕": "NIU", "钮": "NIU", "龔": "GONG",
    "龚": "GONG", "程": "CHENG", "嵇": "JI", "邢": "XING", "裴": "PEI", "陸": "LU",
    "陆": "LU", "榮": "RONG", "荣": "RONG", "翁": "WENG", "荀": "XUN", "羊": "YANG",
    "惠": "HUI", "甄": "ZHEN", "麴": "QU", "封": "FENG", "芮": "RUI", "羿": "YI",
    "儲": "CHU", "储": "CHU", "靳": "JIN", "汲": "JI", "邴": "BING", "糜": "MI",
    "松": "SONG", "井": "JING", "段": "DUAN", "富": "FU", "巫": "WU", "烏": "WU",
    "乌": "WU", "焦": "JIAO", "巴": "BA", "弓": "GONG", "牧": "MU", "山": "SHAN",
    "谷": "GU", "車": "CHE", "车": "CHE", "侯": "HOU", "蓬": "PENG", "全": "QUAN",
    "郗": "XI", "班": "BAN", "仰": "YANG", "秋": "QIU", "仲": "ZHONG", "伊": "YI",
    "宮": "GONG", "宫": "GONG", "寧": "NING", "宁": "NING", "仇": "QIU", "欒": "LUAN",
    "栾": "LUAN", "暴": "BAO", "甘": "GAN", "鈄": "TOU", "钭": "TOU", "厲": "LI",
    "厉": "LI", "戎": "RONG", "祖": "ZU", "武": "WU", "符": "FU", "劉": "LIU",
    "刘": "LIU", "景": "JING", "詹": "ZHAN", "束": "SHU", "龍": "LONG", "龙": "LONG",
    "葉": "YE", "叶": "YE", "司": "SI", "韶": "SHAO", "郜": "GAO", "黎": "LI",
    "薊": "JI", "蓟": "JI", "薄": "BO", "印": "YIN", "宿": "SU", "白": "BAI",
    "懷": "HUAI", "怀": "HUAI", "蒲": "PU", "邰": "TAI", "鄂": "E", "索": "SUO",
    "咸": "XIAN", "籍": "JI", "賴": "LAI", "赖": "LAI", "卓": "ZHUO", "藺": "LIN",
    "蔺": "LIN", "屠": "TU", "蒙": "MENG", "池": "CHI", "喬": "QIAO", "乔": "QIAO",
    "陰": "YIN", "阴": "YIN", "鬱": "YU", "郁": "YU", "胥": "XU", "能": "NENG",
    "蒼": "CANG", "苍": "CANG", "雙": "SHUANG", "双": "SHUANG", "聞": "WEN", "闻": "WEN",
    "莘": "SHEN", "黨": "DANG", "党": "DANG", "翟": "ZHAI", "譚": "TAN", "谭": "TAN",
    "貢": "GONG", "贡": "GONG", "勞": "LAO", "劳": "LAO", "逄": "PANG", "姬": "JI",
    "申": "SHEN", "扶": "FU", "堵": "DU", "冉": "RAN", "宰": "ZAI", "酈": "LI",
    "郦": "LI", "雍": "YONG", "郤": "XI", "璩": "QU", "桑": "SANG", "桂": "GUI",
    "濮": "PU", "牛": "NIU", "壽": "SHOU", "寿": "SHOU", "通": "TONG", "邊": "BIAN",
    "边": "BIAN", "扈": "HU", "燕": "YAN", "冀": "JI", "郟": "JIA", "郏": "JIA",
    "浦": "PU", "尚": "SHANG", "農": "NONG", "农": "NONG", "溫": "WEN", "温": "WEN",
    "別": "BIE", "别": "BIE", "莊": "ZHUANG", "庄": "ZHUANG", "晏": "YAN", "柴": "CHAI",
    "瞿": "QU", "閻": "YAN", "阎": "YAN", "充": "CHONG", "慕": "MU", "連": "LIAN",
    "连": "LIAN", "茹": "RU", "習": "XI", "习": "XI", "宦": "HUAN", "艾": "AI",
    "魚": "YU", "鱼": "YU", "容": "RONG", "古": "GU", "易": "YI", "慎": "SHEN",
    "戈": "GE", "廖": "LIAO", "庾": "YU", "終": "ZHONG", "终": "ZHONG", "暨": "JI",
    "居": "JU", "衡": "HENG", "步": "BU", "都": "DU", "耿": "GENG", "滿": "MAN",
    "满": "MAN", "弘": "HONG", "匡": "KUANG", "國": "GUO", "国": "GUO", "文": "WEN",
    "寇": "KOU", "廣": "GUANG", "广": "GUANG", "祿": "LU", "禄": "LU", "闕": "QUE",
    "阙": "QUE", "東": "DONG", "东": "DONG", "歐": "OU", "欧": "OU", "殳": "SHU",
    "沃": "WO", "利": "LI", "越": "YUE", "隆": "LONG", "師": "SHI", "师": "SHI",
    "鞏": "GONG", "巩": "GONG", "聶": "NIE", "聂": "NIE", "晁": "CHAO", "勾": "GOU",
    "敖": "AO", "融": "RONG", "冷": "LENG", "訾": "ZI", "辛": "XIN", "闞": "KAN",
    "阚": "KAN", "那": "NA", "簡": "JIAN", "简": "JIAN", "饒": "RAO", "饶": "RAO",
    "空": "KONG", "曾": "ZENG", "毋": "WU", "沙": "SHA", "乜": "NIE", "養": "YANG",
    "养": "YANG", "鞠": "JU", "須": "XU", "须": "XU", "豐": "FENG", "丰": "FENG",
    "巢": "CHAO", "關": "GUAN", "关": "GUAN", "蒯": "KUAI", "相": "XIANG", "查": "ZHA",
    "後": "HOU", "后": "HOU", "荊": "JING", "荆": "JING", "紅": "HONG", "红": "HONG",
    "游": "YOU", "竺": "ZHU", "權": "QUAN", "权": "QUAN", "逯": "LU", "蓋": "GE",
    "盖": "GE", "益": "YI", "桓": "HUAN", "公": "GONG", "牟": "MOU", "哈": "HA",
    "言": "YAN", "福": "FU", "肖": "XIAO", "區": "OU", "区": "OU", "麥": "MAI",
    "麦": "MAI", "佟": "TONG", "靖": "JING", "湛": "ZHAN", "謀": "MOU", "谋": "MOU",
    "譚": "TAN", "谭": "TAN", "尋": "XUN", "寻": "XUN", "棘": "JI", "銀": "YIN",
    "银": "YIN", "佴": "ER", "伯": "BO", "賞": "SHANG", "赏": "SHANG", "松": "SONG",
    "段": "DUAN", "甄": "ZHEN", "尉": "WEI", "遲": "CHI", "迟": "CHI", "公": "GONG",
    "孫": "SUN", "孙": "SUN", "公": "GONG", "冶": "YE", "淳": "CHUN", "於": "YU",
    "乜": "NIE", "督": "DU", "仉": "ZHANG", "司": "SI", "邛": "QIONG", "僪": "YU",
    "都": "DU", "粲": "CAN", "僧": "SENG", "薩": "SA", "萨": "SA", "隗": "KUI",
    "穰": "RANG", "還": "HUAN", "邴": "BING", "雒": "LUO", "臧": "ZANG", "紅": "HONG",
    "红": "HONG",
}


def split_en_name(en: str, cn_name: str = None):
    """把英文姓名拆成 (姓, 名)，並一律輸出大寫。
    不論輸入是半形逗號、全形逗號、斜線、點號或空白分隔，都會正確拆開。
    若英文無分隔且提供中文姓名，會用中文姓的拼音自動切開（如 ZHANGYITING + 张依婷 -> ZHANG,YITING）。
    範例：'QU,SHENZHONG'   -> ('QU', 'SHENZHONG')
          'TANG/QINGPING'  -> ('TANG', 'QINGPING')
          'ZHOU.YINHUI'    -> ('ZHOU', 'YINHUI')
          'SHEN DAN'       -> ('SHEN', 'DAN')
          'QIU，JIELEI'    -> ('QIU', 'JIELEI')  （全形逗號）
          'qiu,jielei'     -> ('QIU', 'JIELEI')  （強制大寫）
          'ZHANGYITING'    -> ('ZHANG', 'YITING')  （無分隔，按中文姓切）
    """
    en = (en or "").strip()
    if not en:
        return "", ""
    # 全形逗號先統一成半形逗號
    en = en.replace("，", ",")
    # 逗號 / 斜線 / 點號分隔：前=姓 後=名
    for sep in (",", "/", "."):
        if sep in en:
            a, b = en.split(sep, 1)
            return a.strip().upper(), b.strip().upper()
    # 純空白分隔：中文習慣第一個字是姓
    toks = en.split()
    if len(toks) >= 2:
        return toks[0].upper(), " ".join(toks[1:]).upper()
    # 無分隔：嘗試用中文姓的拼音切開
    if cn_name:
        cn = str(cn_name).strip()
        # 先試複姓（2 字），再試單姓（1 字）
        for n in (2, 1):
            if len(cn) >= n:
                py = _SURNAMES_PINYIN.get(cn[:n])
                if py and en.upper().startswith(py):
                    return py, en[len(py):].upper()
    return "", en.upper()


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

    sur, fir = split_en_name(primary.get("en_name", ""), primary.get("cn_name", ""))
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


# ---------------------------------------------------------------------------
# 中文 / 英文姓名拼音自動核對
# ---------------------------------------------------------------------------
def _cn_to_pinyin(cn: str) -> str:
    """把中文姓名轉成無聲調、無空白的大寫拼音（如 张依婷 -> ZHANGYITING）。"""
    from pypinyin import lazy_pinyin
    return "".join(lazy_pinyin(str(cn))).upper()


def _is_subsequence(a: str, b: str) -> bool:
    """a 是否為 b 的子序列（字元順序一致、可跳過 b 中部分字元）。"""
    it = iter(b)
    return all(ch in it for ch in a)


def verify_name_match(cn_name: str, en_name: str):
    """比對中文姓名與英文姓名拼音是否一致。
    回傳 (ok, message)：ok=True 表示相符或無法判斷；ok=False 附上提示訊息。
    判斷規則：
      - 任一方為空 -> 無法比對，視為 ok
      - 中文拼音 與 英文（去分隔符/空白）完全相等 -> ok
      - 兩者為字母重排（順序不同，如 姓在後）-> ok
      - 一方為另一方子序列（容許英文多稱謂/中間名，或中文只給姓）-> ok
      - 其他（含不同字母）-> 視為不符，回傳提示
    """
    cn = (cn_name or "").strip()
    en = (en_name or "").strip()
    if not cn or not en:
        return True, None
    try:
        cn_py = _cn_to_pinyin(cn)
    except Exception:
        return True, None
    en_norm = re.sub(r"[^A-Z]", "", en.upper())
    if not en_norm:
        return True, None

    if cn_py == en_norm:
        return True, None
    if sorted(cn_py) == sorted(en_norm):
        return True, None
    short, long = (cn_py, en_norm) if len(cn_py) <= len(en_norm) else (en_norm, cn_py)
    if _is_subsequence(short, long):
        return True, None

    return False, (
        f"⚠️ 中文姓名「{cn}」與英文姓名拼音「{en}」似乎不符，"
        f"請確認英文拼音是否正確（預期拼音：{cn_py}）。"
    )


def verify_booking_names(booking: dict):
    """核對訂房資料中所有入住者的 中文/英文 姓名拼音。
    回傳 (all_ok, warnings) ；warnings 為不符的提示清單。
    """
    guests = booking.get("guests") or []
    warnings = []
    for i, g in enumerate(guests, 1):
        ok, msg = verify_name_match(g.get("cn_name", ""), g.get("en_name", ""))
        if not ok and msg:
            prefix = f"第{i}位入住者：" if len(guests) > 1 else ""
            warnings.append(prefix + msg)
    return (len(warnings) == 0, warnings)

