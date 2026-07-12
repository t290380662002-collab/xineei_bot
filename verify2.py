# -*- coding: utf-8 -*-
import openpyxl, glob, os
from config import HOTELS

files = sorted(glob.glob("C:/Users/t2903/WorkBuddy/2026-07-13-01-18-38/output/*.xlsx"))
for f in files:
    hotel = os.path.basename(f).split("_")[1]
    cfg = HOTELS.get(hotel)
    if not cfg:
        continue
    print("=" * 70)
    print("FILE:", os.path.basename(f), "-> 客人清單頁:", cfg["guest_sheet"])
    wb = openpyxl.load_workbook(f)
    gws = wb[cfg["guest_sheet"]]
    for row in gws.iter_rows():
        cells = [(c.coordinate, c.value) for c in row if c.value not in (None, "")]
        if cells:
            print("  ", cells)
