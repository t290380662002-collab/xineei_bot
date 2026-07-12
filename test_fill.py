# -*- coding: utf-8 -*-
"""測試：對四間飯店各產生一份填好的範例 Excel，存到 output/ 供肉眼確認。"""
import os
from fill import fill_booking, output_filename

os.makedirs("output", exist_ok=True)

samples = {
    "名匯": {
        "飯店": "名匯", "入住": "2026/07/20", "退房": "2026/07/22",
        "房型": "RK", "件數": "2", "備注": "高樓層、安靜房", "是否吸煙": "不吸煙",
        "guests": [
            {"cn_name": "渠慎重", "en_name": "QU,SHENZHONG", "dob": "1961/06/11", "idno": "M41646681"},
            {"cn_name": "孫越越", "en_name": "SUN,YUEYUE", "dob": "1986/10/19", "idno": "CH5970813"},
        ],
    },
    "威尼斯": {
        "飯店": "威尼斯", "入住": "2026/08/01", "退房": "2026/08/03",
        "房型": "皇室套房", "件數": "1", "備注": "機場接送", "是否吸煙": "吸煙",
        "guests": [
            {"cn_name": "褚國華", "en_name": "CHU,GUOHUA", "dob": "1977/03/05", "idno": "C97476513"},
        ],
    },
    "巴黎人": {
        "飯店": "巴黎人", "入住": "2026/09/10", "退房": "2026/09/12",
        "房型": "KC", "件數": "3", "備注": "", "是否吸煙": "不吸煙",
        "guests": [
            {"cn_name": "楊雅然", "en_name": "YANG,YA-RAN", "dob": "1998/03/20", "idno": "370384695"},
            {"cn_name": "呂俞靜", "en_name": "LU,YU-JING", "dob": "1996/01/18", "idno": "353021497"},
            {"cn_name": "劉奕妘", "en_name": "LIU,YI-YUN", "dob": "2000/11/14", "idno": "365247574"},
        ],
    },
    "倫敦人": {
        "飯店": "倫敦人", "入住": "2026/10/05", "退房": "2026/10/07",
        "房型": "DBK1", "件數": "1", "備注": "生日布置", "是否吸煙": "不吸煙",
        "guests": [
            {"cn_name": "BANG INHO", "en_name": "BANG,INHO", "dob": "1961/06/11", "idno": "M41646681"},
        ],
    },
}

for hotel, booking in samples.items():
    bio = fill_booking(booking)
    fn = output_filename(booking)
    path = os.path.join("output", fn)
    with open(path, "wb") as f:
        f.write(bio.getvalue())
    print("已產生：", path)
print("完成。")
