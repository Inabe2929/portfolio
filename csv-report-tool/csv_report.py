"""売上CSV 統合・月次集計レポート自動作成ツール

フォルダ内の売上CSV（店舗やシステムごとに形式がバラバラでもOK）を読み込み、
整形・集計した Excel レポートを1クリックで作成します。

対応している「バラつき」:
  - 文字コード: UTF-8 / UTF-8(BOM) / Shift_JIS(cp932)
  - 列名の表記ゆれ: 「日付/売上日/date」「数量/個数/qty」など
  - 日付形式: 2026/1/5, 2026-01-05, 20260105, 2026年1月5日
  - 数値: 「1,200」「¥1,200」「１２００」(全角) など
  - 完全に同じ行の重複

使い方:
  py csv_report.py 入力フォルダ [-o 出力ファイル.xlsx]
"""
from __future__ import annotations

import argparse
import csv
import re
import sys
import unicodedata
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------- 設定
COLUMN_ALIASES = {
    "date": ["日付", "売上日", "販売日", "取引日", "date"],
    "store": ["店舗", "店舗名", "店名", "store", "shop"],
    "product": ["商品", "商品名", "品名", "product", "item"],
    "qty": ["数量", "個数", "販売数", "qty", "quantity"],
    "price": ["単価", "価格", "販売単価", "price", "unit_price"],
}
REQUIRED = list(COLUMN_ALIASES)
ENCODINGS = ["utf-8-sig", "cp932"]

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TOTAL_FILL = PatternFill("solid", fgColor="DDEBF7")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
YEN = '#,##0;[Red]-#,##0'


@dataclass(frozen=True)
class Row:
    date: date
    store: str
    product: str
    qty: int
    price: int
    source: str

    @property
    def amount(self) -> int:
        return self.qty * self.price


@dataclass
class Problem:
    source: str
    line: int
    reason: str
    raw: str


# ---------------------------------------------------------------- 読み込み・整形
def norm(text: str) -> str:
    """全角→半角、前後空白除去、小文字化（列名照合用）。"""
    return unicodedata.normalize("NFKC", text or "").strip().lower()


def read_text(path: Path) -> str:
    data = path.read_bytes()
    for enc in ENCODINGS:
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    raise ValueError("文字コードを判別できません（UTF-8 / Shift_JIS 以外）")


def map_columns(header: list[str]) -> dict[str, int]:
    lookup = {norm(a): key for key, aliases in COLUMN_ALIASES.items() for a in aliases}
    mapping: dict[str, int] = {}
    for i, name in enumerate(header):
        key = lookup.get(norm(name))
        if key and key not in mapping:
            mapping[key] = i
    missing = [k for k in REQUIRED if k not in mapping]
    if missing:
        names = "、".join(COLUMN_ALIASES[k][0] for k in missing)
        raise ValueError(f"必要な列が見つかりません: {names}")
    return mapping


def parse_date(text: str) -> date:
    s = unicodedata.normalize("NFKC", text).strip()
    s = re.sub(r"[年月]", "/", s).replace("日", "")
    for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%Y%m%d", "%Y.%m.%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"日付として読めません: {text!r}")


def parse_int(text: str, label: str) -> int:
    s = unicodedata.normalize("NFKC", text).strip()
    s = re.sub(r"[,¥\\円\s]", "", s)
    if not re.fullmatch(r"-?\d+(\.0+)?", s):
        raise ValueError(f"{label}が数値ではありません: {text!r}")
    return int(float(s))


def load_folder(folder: Path) -> tuple[list[Row], list[Problem], list[str]]:
    rows: list[Row] = []
    problems: list[Problem] = []
    log: list[str] = []
    seen: set[tuple] = set()
    files = sorted(folder.glob("*.csv"))
    if not files:
        raise SystemExit(f"CSVファイルが見つかりません: {folder}")

    for path in files:
        try:
            text = read_text(path)
            reader = csv.reader(text.splitlines())
            header = next(reader, [])
            cols = map_columns(header)
        except (ValueError, StopIteration) as e:
            problems.append(Problem(path.name, 1, str(e), ""))
            log.append(f"× {path.name}: 読み込めませんでした（{e}）")
            continue

        ok = dup = 0
        for line_no, rec in enumerate(reader, start=2):
            if not any(c.strip() for c in rec):
                continue  # 空行
            try:
                get = lambda k: rec[cols[k]] if cols[k] < len(rec) else ""
                row = Row(
                    date=parse_date(get("date")),
                    store=unicodedata.normalize("NFKC", get("store")).strip(),
                    product=unicodedata.normalize("NFKC", get("product")).strip(),
                    qty=parse_int(get("qty"), "数量"),
                    price=parse_int(get("price"), "単価"),
                    source=path.name,
                )
                if not row.store or not row.product:
                    raise ValueError("店舗または商品が空欄です")
            except ValueError as e:
                problems.append(Problem(path.name, line_no, str(e), ",".join(rec)))
                continue
            key = (row.date, row.store, row.product, row.qty, row.price)
            if key in seen:
                dup += 1
                problems.append(Problem(path.name, line_no, "重複行のため除外", ",".join(rec)))
                continue
            seen.add(key)
            rows.append(row)
            ok += 1
        log.append(f"○ {path.name}: {ok}件 取り込み" + (f"（重複 {dup}件 除外）" if dup else ""))
    return rows, problems, log


# ---------------------------------------------------------------- Excel 出力
def style_header(ws, row: int, ncols: int) -> None:
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font, cell.border = HEADER_FILL, HEADER_FONT, BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")


def autosize(ws, min_width: int = 8, max_width: int = 50) -> None:
    for col in ws.columns:
        width = max(
            (sum(2 if ord(ch) > 0x7F else 1 for ch in str(c.value)) for c in col if c.value is not None),
            default=0,
        )
        ws.column_dimensions[get_column_letter(col[0].column)].width = max(min_width, min(max_width, width + 2))


def write_table(ws, headers, data, money_cols=(), start_row=1):
    ws.append([]) if start_row > 1 else None
    for c, h in enumerate(headers, 1):
        ws.cell(row=start_row, column=c, value=h)
    style_header(ws, start_row, len(headers))
    for r, values in enumerate(data, start_row + 1):
        for c, v in enumerate(values, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BORDER
            if c in money_cols:
                cell.number_format = YEN
            if isinstance(v, date):
                cell.number_format = "yyyy/mm/dd"
    ws.freeze_panes = ws.cell(row=start_row + 1, column=1)


def build_report(rows: list[Row], problems: list[Problem], log: list[str], out: Path) -> None:
    wb = Workbook()

    # --- サマリー
    ws = wb.active
    ws.title = "サマリー"
    total = sum(r.amount for r in rows)
    months = sorted({r.date.strftime("%Y-%m") for r in rows})
    ws["A1"] = "売上集計レポート"
    ws["A1"].font = Font(size=16, bold=True, color="1F4E78")
    info = [
        ("作成日時", datetime.now().strftime("%Y/%m/%d %H:%M")),
        ("対象期間", f"{months[0]} 〜 {months[-1]}" if months else "-"),
        ("取り込み件数", len(rows)),
        ("要確認の行", len(problems)),
        ("売上合計", total),
    ]
    for i, (k, v) in enumerate(info, 3):
        ws.cell(row=i, column=1, value=k).font = Font(bold=True)
        cell = ws.cell(row=i, column=2, value=v)
        if k == "売上合計":
            cell.number_format = YEN + '"円"'
            cell.font = Font(bold=True, size=13)
    ws.cell(row=9, column=1, value="読み込み結果").font = Font(bold=True)
    for i, line in enumerate(log, 10):
        ws.cell(row=i, column=1, value=line)
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 28

    # --- 月別×店舗
    stores = sorted({r.store for r in rows})
    pivot: dict[str, dict[str, int]] = defaultdict(lambda: defaultdict(int))
    for r in rows:
        pivot[r.date.strftime("%Y-%m")][r.store] += r.amount
    ws2 = wb.create_sheet("月別×店舗")
    data = [[m, *[pivot[m][s] for s in stores], sum(pivot[m].values())] for m in months]
    totals = ["合計", *[sum(pivot[m][s] for m in months) for s in stores], total]
    write_table(ws2, ["月", *stores, "合計"], data + [totals], money_cols=range(2, len(stores) + 3))
    last = len(data) + 2
    for c in range(1, len(stores) + 3):
        ws2.cell(row=last, column=c).fill = TOTAL_FILL
        ws2.cell(row=last, column=c).font = Font(bold=True)
    autosize(ws2, min_width=12)
    if data:
        chart = BarChart()
        chart.type, chart.grouping, chart.overlap = "col", "stacked", 100
        chart.title, chart.y_axis.title = "月別売上（店舗別）", "円"
        chart.add_data(Reference(ws2, min_col=2, max_col=len(stores) + 1, min_row=1, max_row=len(data) + 1), titles_from_data=True)
        chart.set_categories(Reference(ws2, min_col=1, min_row=2, max_row=len(data) + 1))
        chart.height, chart.width = 9, 18
        ws2.add_chart(chart, f"{get_column_letter(len(stores) + 4)}2")

    # --- 商品ランキング
    by_product: dict[str, list[int]] = defaultdict(lambda: [0, 0])
    for r in rows:
        by_product[r.product][0] += r.qty
        by_product[r.product][1] += r.amount
    ranking = sorted(by_product.items(), key=lambda kv: kv[1][1], reverse=True)
    ws3 = wb.create_sheet("商品ランキング")
    write_table(
        ws3,
        ["順位", "商品", "販売数", "売上", "構成比"],
        [[i, p, q, a, a / total if total else 0] for i, (p, (q, a)) in enumerate(ranking, 1)],
        money_cols=(3, 4),
    )
    for r in range(2, len(ranking) + 2):
        ws3.cell(row=r, column=5).number_format = "0.0%"
    autosize(ws3, min_width=10)

    # --- 明細
    ws4 = wb.create_sheet("明細")
    write_table(
        ws4,
        ["日付", "店舗", "商品", "数量", "単価", "金額", "元ファイル"],
        [[r.date, r.store, r.product, r.qty, r.price, r.amount, r.source]
         for r in sorted(rows, key=lambda r: (r.date, r.store, r.product))],
        money_cols=(4, 5, 6),
    )
    ws4.auto_filter.ref = ws4.dimensions
    autosize(ws4, min_width=10)

    # --- 要確認
    ws5 = wb.create_sheet("要確認")
    write_table(ws5, ["ファイル", "行", "理由", "元データ"],
                [[p.source, p.line, p.reason, p.raw] for p in problems])
    autosize(ws5)
    if problems:
        ws5.sheet_properties.tabColor = "C00000"

    wb.save(out)


# ---------------------------------------------------------------- main
def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="売上CSVを統合して集計Excelを作成します")
    ap.add_argument("input_dir", type=Path, help="CSVが入っているフォルダ")
    ap.add_argument("-o", "--output", type=Path, help="出力先 .xlsx（省略時は入力フォルダ内）")
    args = ap.parse_args(argv)

    if not args.input_dir.is_dir():
        print(f"フォルダが見つかりません: {args.input_dir}", file=sys.stderr)
        return 1
    out = args.output or args.input_dir / f"売上レポート_{datetime.now():%Y%m%d_%H%M}.xlsx"

    rows, problems, log = load_folder(args.input_dir)
    print("\n".join(log))
    build_report(rows, problems, log, out)
    print(f"\n完了: {len(rows)}件を集計 / 要確認 {len(problems)}件")
    print(f"出力: {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
