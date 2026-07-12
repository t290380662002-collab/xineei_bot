# -*- coding: utf-8 -*-
"""用使用者提供的真實訂房文字測試解析 + 填表。"""
from pathlib import Path
from parse_text import parse_booking_text, looks_like_booking
from fill import fill_booking, output_filename

SAMPLE = """飯店：入住：7月15日
退房：7月17日
飯店：巴黎人酒店
房型:香槟套房/里昂套房（双床）
件數：1间2晚
備注：高楼层
是否吸煙：吸烟房

入住者中文：吴婷婷
入住者英文：WU, TINGTING
出生年月日：1982.07.04
證件號碼：C73960683
"""


def main():
    print("==== looks_like_booking ====")
    print(looks_like_booking(SAMPLE))

    b = parse_booking_text(SAMPLE)
    print("\n==== 解析結果 ====")
    for k in ["飯店", "入住", "退房", "房型", "件數", "備注", "是否吸煙"]:
        print(f"  {k}: {b[k]!r}")
    print("  guests:", b["guests"])

    print("\n==== 填表並讀回驗證 ====")
    bio = fill_booking(b)
    fn = output_filename(b)
    print("  產出檔名:", fn)

    import openpyxl
    from io import BytesIO
    wb = openpyxl.load_workbook(BytesIO(bio.getvalue()))
    # 主表
    mws = wb["Parisian"]
    print("  主表 姓/名:", mws["B16"].value, "/", mws["H16"].value)
    print("  主表 證件:", mws["H17"].value, " 出生:", mws["B18"].value)
    print("  主表 入住/退房:", mws["B20"].value, "/", mws["B21"].value)
    print("  主表 房數/人數:", mws["M21"].value, "/", mws["M20"].value)
    print("  主表 備注:", mws["L14"].value)
    # 房型方框：列出被打勾的
    print("  主表 已打勾房型方框:")
    for cell, code, cn in __import__("config").HOTELS["巴黎人"]["room_types"]:
        v = mws[cell].value
        if v and "(✓)" in str(v):
            print(f"    {cell} {code} {cn} -> {v}")

    # 客人清單
    gws = wb["Sheet1"]
    print("  客人清單 row2:", [gws[f"{c}2"].value for c in "ABCDE"])

    # 存一張給使用者看
    out = Path("output") / fn
    out.parent.mkdir(exist_ok=True)
    with open(out, "wb") as f:
        f.write(bio.getvalue())
    print("\n  已寫出範例:", out)


if __name__ == "__main__":
    main()
