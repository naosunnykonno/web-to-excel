#!/usr/bin/env python3
"""web_to_excel.py — Scrape an HTML table from a web page into Excel.

Fetches a page, extracts a <table> (by default the first table with
class "wikitable", which covers most Wikipedia list pages), cleans the
cells (thousands separators are converted to real numbers), and writes
an Excel workbook:

  Sheet "Data"      the table itself, formatted
  Sheet "取得情報"   source URL, page title, scrape timestamp (JST),
                    and row/column counts — so the file stays auditable

Usage:
    python3 web_to_excel.py [url] [output_xlsx] [--table-index N]
    (default URL: Wikipedia, List of Japanese prefectures by population)
"""
import re
import sys
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

DEFAULT_URL = ("https://en.wikipedia.org/wiki/"
               "List_of_Japanese_prefectures_by_population")
HEADERS = {"User-Agent": "web-to-excel-sample/1.0 (portfolio demo)"}


def clean_cell(text):
    """Trim whitespace; convert plain numbers (with commas) to int/float."""
    text = re.sub(r"\s+", " ", text).strip()
    plain = text.replace(",", "").replace("−", "-")
    if re.fullmatch(r"-?\d+", plain):
        return int(plain)
    if re.fullmatch(r"-?\d+\.\d+", plain):
        return float(plain)
    return text


def scrape_table(url, table_index=0):
    res = requests.get(url, headers=HEADERS, timeout=30)
    res.raise_for_status()
    soup = BeautifulSoup(res.text, "lxml")
    title = soup.title.string.strip() if soup.title else url

    tables = soup.select("table.wikitable") or soup.find_all("table")
    if table_index >= len(tables):
        raise SystemExit(f"Table index {table_index} not found "
                         f"({len(tables)} tables on page)")
    table = tables[table_index]

    rows = []
    for tr in table.find_all("tr"):
        cells = tr.find_all(["th", "td"])
        if not cells:
            continue
        row = []
        for cell in cells:
            # Repeat the value across spanned columns to keep a grid shape.
            span = int(cell.get("colspan", 1))
            row.extend([clean_cell(cell.get_text())] + [""] * (span - 1))
        rows.append(row)
    if not rows:
        raise SystemExit("No rows extracted from the table")

    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    return title, rows


HEADER_FILL = PatternFill("solid", fgColor="305496")
HEADER_FONT = Font(bold=True, color="FFFFFF")
THIN = Side(style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def write_excel(url, title, rows, out_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    for row in rows:
        ws.append(row)
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.border = BORDER
            if isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0.##"
    for col in range(1, ws.max_column + 1):
        longest = max(len(str(c.value or "")) for c in ws.iter_cols(min_col=col, max_col=col).__next__())
        ws.column_dimensions[get_column_letter(col)].width = min(max(10, longest + 2), 45)
    ws.freeze_panes = "A2"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    now_jst = datetime.now(ZoneInfo("Asia/Tokyo"))
    ws2 = wb.create_sheet("取得情報")
    meta = [
        ["項目", "内容"],
        ["取得元URL", url],
        ["ページタイトル", title],
        ["取得日時", now_jst.strftime("%Y-%m-%d %H:%M:%S %Z")],
        ["データ行数（ヘッダー除く）", len(rows) - 1],
        ["列数", len(rows[0])],
        ["ツール", "web_to_excel.py (requests + BeautifulSoup + openpyxl)"],
    ]
    for row in meta:
        ws2.append(row)
    for cell in ws2[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    ws2.column_dimensions["A"].width = 26
    ws2.column_dimensions["B"].width = 95

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return now_jst


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    url = args[0] if len(args) > 0 else DEFAULT_URL
    out_path = Path(args[1]) if len(args) > 1 else Path("output/prefectures_population.xlsx")
    table_index = 0
    if "--table-index" in sys.argv:
        table_index = int(sys.argv[sys.argv.index("--table-index") + 1])

    title, rows = scrape_table(url, table_index)
    ts = write_excel(url, title, rows, out_path)
    print(f"Page: {title}")
    print(f"Rows: {len(rows) - 1} data rows, Cols: {len(rows[0])}")
    print(f"Header: {rows[0]}")
    print(f"First row: {rows[1]}")
    print(f"Scraped at: {ts.strftime('%Y-%m-%d %H:%M:%S %Z')}")
    print(f"Wrote {out_path}")


if __name__ == "__main__":
    main()
