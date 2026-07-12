# -*- coding: utf-8 -*-
"""離線測試：用使用者提供的文字格式解析並填表，再讀回驗證。"""
import os
from parse_text import parse_booking_text, looks_like_booking
from fill import fill_booking, output_filename, resolve_hotel
import openpyxl

SAMPLE = """入住：2026/07/20
退房：2026/07/22
飯店：名匯
房型:RK
件數：2間
備注：高樓層
是否吸煙：不吸煙
入住者中文：渠慎重
入住者英文：QU,SHENZHONG
出生年月日：1961/06/11
證件號碼：E12345678
"""

# 1) 判斷函式
print("== looks_like_booking ==")
print("  訂房文字 ->", looks_like_booking(SAMPLE))
print("  隨機hi   ->", looks_like_booking("你好"))
print("  單行     ->", looks_like_booking("飯店：威尼斯"))

# 2) 解析
print("\n== 解析結果 ==")
b = parse_booking_text(SAMPLE)
import json
print(json.dumps(b, ensure_ascii=False, indent=2))

# 3) 填表
print("\n== 產檔 ==")
bio = fill_booking(b)
fn = output_filename(b)
os.makedirs("output", exist_ok=True)
out = os.path.join("output", fn)
with open(out, "wb") as f:
    f.write(bio.getvalue())
print("  輸出檔：", out)

# 4) 讀回驗證（名匯專案）
from config import HOTELS, TEMPLATES_DIR
hk = resolve_hotel(b["飯店"])
cfg = HOTELS[hk]
wb = openpyxl.load_workbook(out)
mws = wb[cfg["main_sheet"]]
mc = cfg["main_cells"]
print("\n== 主表關鍵欄位 ==")
print("  姓/名:", mws[mc["surname"]].value, "/", mws[mc["firstname"]].value)
print("  證件:", mws[mc["idno"]].value)
print("  出生:", mws[mc["dob"]].value)
print("  入住/退房:", mws[mc["checkin"]].value, "/", mws[mc["checkout"]].value)
print("  房數/人數:", mws[mc["rooms"]].value, "/", mws[mc["pax"]].value)
print("  備注:", mws[mc["remark"]].value)

# 找打勾的房型方框
print("\n== 房型方框(含✓) ==")
for cell, code, cn in cfg["room_types"]:
    v = mws[cell].value
    if v and "✓" in str(v):
        print(f"  {cell} -> {v}  ({code}/{cn})")

# 客人清單
gws = wb[cfg["guest_sheet"]]
gc = cfg["guest_cols"]
print("\n== 客人清單 第1列 ==")
for colname, col in gc.items():
    print(f"  {colname}({col}):", gws[f"{col}{cfg['guest_first_row']}"].value)

print("\n[測試完成]")
