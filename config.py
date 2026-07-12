# -*- coding: utf-8 -*-
"""
四間飯店 Excel 模板的欄位對應設定。
每個飯店包含：
  file         : 模板檔名（放在 templates/ 下）
  main_sheet   : 主表單工作表名稱
  guest_sheet  : 客人清單工作表名稱
  main_cells   : 主表單要填的儲存格
  guest_first_row : 客人清單第一筆資料起始列
  guest_cols   : 客人清單各欄的欄位字母
  room_types   : [(儲存格, 代碼, 中文名), ...]  用來對到方框打勾
"""
from pathlib import Path

TEMPLATES_DIR = Path(__file__).parent / "templates"

# 飯店中文名 -> 內部 key（用來比對使用者輸入的「飯店」欄位）
HOTEL_KEYS = ["名匯", "威尼斯", "巴黎人", "倫敦人"]

HOTELS = {
    "名匯": {
        "file": "名匯-空白.xlsx",
        "main_sheet": "Londoner Grand",
        "guest_sheet": "Sheet1",
        "main_cells": {
            "surname": "B16",    # 姓 Surname
            "firstname": "H16",  # 名 First Name
            "patron": "B17",     # 會員號碼 Patron#
            "idno": "H17",       # 證件號碼 Passport/ID#
            "dob": "B18",        # 出生日期 DOB
            "checkin": "B20",    # 入住日期 C/I Date
            "pax": "M20",        # 人數 Pax
            "checkout": "B21",   # 退房日期 C/O Date
            "rooms": "M21",      # 房數 No. of Rooms
            "remark": "L14",     # 特別要求 Special request
        },
        "guest_first_row": 9,
        "guest_cols": {
            "cn_name": "G",   # 中文姓名
            "en_name": "H",   # 英文姓名
            "idno": "I",      # 證件號碼
            "dob": "J",       # 出生日期
            "roomtype": "K",  # 房型
            "smoking": "L",   # 不吸煙（只有名匯有這欄）
        },
        "room_types": [
            ("C23", "RK", "豪華大床房"),
            ("G23", "GK", "尊貴豪華大床房"),
            ("K23", "GSK/GSKP", "尊貴伊莉莎白大床套房"),
            ("N23", "GC2/GC2B/GC2P", "尊貴凱瑟琳套房"),
            ("C24", "R2", "豪華雙床房"),
            ("G24", "G2", "尊貴豪華雙床房"),
            ("K24", "GS2", "尊貴伊莉莎白雙床套房"),
            ("N24", "GHK/GH2B", "尊貴亨利套房"),
            ("C25", "VK", "路氹景觀大床房"),
            ("G25", "LSK/LSKB", "伊莉莎白大床套房"),
            ("K25", "CS2", "凱瑟琳套房"),
            ("N25", "TS", "尊貴陽台套房"),
            ("C26", "V2", "路氹景觀雙床房"),
            ("G26", "LS2/LS2B", "伊莉莎白雙床套房"),
            ("K26", "CS", "亨利套房"),
            ("N26", "JS", "尊貴浴池套房"),
            ("C27", "T2B/T2BP", "都鐸套房"),
            ("G27", "L2B", "蘭開斯特套房"),
            ("K27", "Y2B", "約克套房"),
            ("N27", "N2B/N2BP", "諾曼套房"),
            ("C28", "H2B", "漢諾威套房"),
            ("G28", "H2BG", "漢諾威泳池套房"),
        ],
    },

    "倫敦人": {
        "file": "倫敦人-空白檔.xlsx",
        "main_sheet": "Londoner",
        "guest_sheet": "工作表1",
        "main_cells": {
            "surname": "B16",
            "firstname": "H16",
            "patron": "B17",
            "idno": "H17",
            "dob": "B18",
            "checkin": "B20",
            "pax": "M20",
            "checkout": "B21",
            "rooms": "M21",
            "remark": "L14",
        },
        "guest_first_row": 3,
        "guest_cols": {
            "cn_name": "B",   # 中文姓名
            "en_name": "C",   # 英文姓名
            "idno": "D",      # 證件號碼
            "dob": "E",       # 出生日期
            "roomtype": "F",  # 房型
        },
        "room_types": [
            ("C24", "KC", "路易套房"),
            ("G24", "DBK1", "大衛碧咸套房"),
            ("C25", "TC", "維多利亞套房"),
            ("G25", "DBKD2", "大衛碧咸套房"),
            ("C26", "KS", "溫莎大床套房"),
            ("G26", "DBKQ2", "大衛碧咸套房"),
            ("C27", "TS", "溫莎雙床套房"),
            ("G27", "DBKQD3", "大衛碧咸套房"),
        ],
    },

    "威尼斯": {
        "file": "威尼斯-空白檔.xlsx",
        "main_sheet": "Venetian",
        "guest_sheet": "工作表1",
        "main_cells": {
            "surname": "B16",
            "firstname": "H16",
            "patron": "B17",
            "idno": "H17",
            "dob": "B18",
            "checkin": "B20",
            "pax": "M20",
            "checkout": "B21",
            "rooms": "M21",
            "remark": "L14",
        },
        "guest_first_row": 3,
        "guest_cols": {
            "cn_name": "C",   # 中文姓名
            "en_name": "D",   # 英文姓名
            "idno": "E",      # 證件號碼
            "dob": "F",       # 出生日期
            "roomtype": "G",  # 房型
        },
        "room_types": [
            ("C24", "KC", "皇室套房"),
            ("G24", "KD", "皇室套房金光景"),
            ("K24", "TC", "貝麗套房"),
            ("N24", "TD", "貝麗套房金光景"),
            ("C25", "KP", "皇室套房"),
            ("G25", "TP", "貝麗套房"),
            ("K25", "TS/TSZ", "維雅雙人床套房"),
            ("N25", "KS/KSZ", "維雅套房"),
            ("C26", "TJ", "維羅納套房"),
            ("G26", "KF", "天倫樂套房"),
            ("C27", "SML1", "米蘭套房"),
            ("G27", "SRO1", "羅馬套房"),
            ("K27", "SFL2", "佛羅倫斯套房"),
            ("N27", "SVN2", "威尼斯套房"),
            ("C28", "SVP2", "威尼斯套房有泳池"),
            ("G28", "SCL2", "斯雅萊套房"),
            ("K28", "SEM2", "Empore E5"),
            ("N28", "SNA2", "拿破崙套房"),
            ("C29", "SNP2", "拿破崙套房有泳池"),
            ("G29", "SPA3", "Palazzo P7"),
            ("K29", "SPP3", "Palazzo 有泳池"),
            ("N29", "SPR4", "總統套房"),
        ],
    },

    "巴黎人": {
        "file": "巴黎人-空白檔.xlsx",
        "main_sheet": "Parisian",
        "guest_sheet": "Sheet1",
        "main_cells": {
            "surname": "B16",
            "firstname": "H16",
            "patron": "B17",
            "idno": "H17",
            "dob": "B18",
            "checkin": "B20",
            "pax": "M20",
            "checkout": "B21",
            "rooms": "M21",
            "remark": "L14",
        },
        "guest_first_row": 2,
        "guest_cols": {
            "cn_name": "A",   # 中文姓名
            "en_name": "B",   # 英文姓名
            "dob": "C",       # 出生日期（巴黎人順序：出生在前）
            "idno": "D",      # 證件號碼
            "roomtype": "E",  # 房型
        },
        "room_types": [
            ("C24", "KC", "豪華大床客房"),
            ("G24", "TC", "豪華雙人床客房"),
            ("K24", "KD", "艾菲爾大床客房"),
            ("N24", "TD", "艾菲爾雙人床客房"),
            ("C25", "QF", "天倫大床客房"),
            ("G25", "KSJ", "康城套房"),
            ("K25", "KS", "里昂大床套房"),
            ("N25", "KSV", "里昂大床套房-艾菲爾景觀"),
            ("C26", "TSV", "里昂雙人床套房-艾菲爾景觀"),
            ("G26", "KSL", "典雅里昂大床套房"),
            ("K26", "KSLV", "典雅里昂大床套房-艾菲爾景觀"),
            ("N26", "TSLV", "典雅里昂雙床套房-艾菲爾景觀"),
            ("C27", "TPS", "香檳雙床人套房"),
            ("G27", "TPSV", "香檳雙床人套房-艾菲爾景觀"),
            ("K27", "TSL", "典雅里昂雙床套房"),
            ("N27", "KPSV", "香檳大床套房-艾菲爾景觀"),
            ("C28", "KPS", "香檳大床套房"),
            ("G28", "SPAK2", "巴黎套房"),
            ("K28", "SPAK2V", "巴黎套房-艾菲爾景觀"),
            ("N28", "SMAK1V", "馬賽套房-艾菲爾景觀"),
            ("C29", "SPPAK2", "尊貴巴黎套房"),
            ("G29", "SMAK1", "馬賽雙人床套房"),
            ("K29", "SMAK1", "馬賽套房"),
            ("N29", "SPMK1V+TPV", "尊貴巴黎鐵塔馬賽套房+尊貴巴黎鐵塔雙人床客房"),
            ("C30", "SVRK3", "凡爾賽王室套房"),
            ("G30", "SPPK2V", "尊貴巴黎鐵塔巴黎套房"),
            ("K30", "SPMAK1+TP", "尊貴馬賽套房+尊貴豪華雙人床客房"),
        ],
    },
}


def resolve_hotel(name: str):
    """把使用者輸入的飯店名對到內部 key。"""
    if not name:
        return None
    n = name.strip()
    for key in HOTEL_KEYS:
        if key in n:
            return key
    return None
