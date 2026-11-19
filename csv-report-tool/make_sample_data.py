"""デモ用の「わざと汚い」売上CSVを sample_input/ に作成する。"""
import random
from datetime import date, timedelta
from pathlib import Path

random.seed(42)
OUT = Path(__file__).parent / "sample_input"
OUT.mkdir(exist_ok=True)
PRODUCTS = {"ブレンドコーヒー": 450, "カフェラテ": 520, "チーズケーキ": 480, "サンドイッチ": 650, "クッキー": 280}


def days():
    d = date(2026, 7, 1)
    while d <= date(2026, 9, 30):
        yield d
        d += timedelta(days=1)


def records(store):
    for d in days():
        for p, price in PRODUCTS.items():
            if random.random() < 0.6:
                yield d, store, p, random.randint(1, 25), price


# 渋谷店: UTF-8, 標準的な列名, YYYY/MM/DD
with open(OUT / "渋谷店_売上.csv", "w", encoding="utf-8", newline="") as f:
    f.write("日付,店舗,商品名,数量,単価\n")
    rows = list(records("渋谷店"))
    for d, s, p, q, pr in rows:
        f.write(f"{d:%Y/%m/%d},{s},{p},{q},{pr}\n")
    d, s, p, q, pr = rows[10]
    f.write(f"{d:%Y/%m/%d},{s},{p},{q},{pr}\n")  # 重複行
    f.write("2026/08/32,渋谷店,カフェラテ,3,520\n")  # 存在しない日付

# 新宿店: Shift_JIS(レジ出力想定), 列名が違う, 日付が YYYYMMDD, 金額にカンマと¥
with open(OUT / "shinjuku_pos.csv", "w", encoding="cp932", newline="") as f:
    f.write("売上日,店名,品名,個数,販売単価,備考\n")
    for d, s, p, q, pr in records("新宿店"):
        f.write(f'{d:%Y%m%d},{s},{p},{q},"￥{pr:,}",\n')
    f.write("20260815,新宿店,クッキー,三,280,\n")  # 数量が漢数字

# 横浜店: UTF-8 BOM(Excel保存想定), 日付が和文, 全角数字混在, 列順も違う
with open(OUT / "yokohama.csv", "w", encoding="utf-8-sig", newline="") as f:
    f.write("店舗名,売上日,商品,単価,数量\n")
    for d, s, p, q, pr in records("横浜店"):
        qty = str(q).translate(str.maketrans("0123456789", "０１２３４５６７８９")) if q % 4 == 0 else q
        f.write(f"{s},{d.year}年{d.month}月{d.day}日,{p},{pr},{qty}\n")
    f.write("横浜店,2026年9月10日,,520,2\n")  # 商品が空欄

# 形式が全く違うファイル（読み飛ばされることの確認用）
with open(OUT / "memo.csv", "w", encoding="utf-8", newline="") as f:
    f.write("メモ,担当\n来月から新商品追加,田中\n")

print(f"サンプルCSVを作成しました: {OUT}")
