[日本語](README.md) | [English](README.en.md)

# Web Table → Excel Scraper

A sample tool that fetches an HTML table from a web page and exports it
as a formatted Excel workbook — automating the "copy the table from the
browser into Excel by hand" routine.

## Input → Output

- **Input**: the URL of a public web page containing a table
  (default: Wikipedia, "List of Japanese prefectures by population")
- **Output**: `output/prefectures_population.xlsx`
  - Sheet "Data" — the table itself; comma-formatted numbers are
    converted to real Excel numbers
  - Sheet "取得情報" (Metadata) — source URL, page title,
    **scrape timestamp (JST)**, and row/column counts, so the file
    itself records where the data came from and when

![Output: Excel](preview_excel.png)

Sample run: 48 rows (47 prefectures + national total) × 10 columns.

## Usage

```bash
pip install -r requirements.txt
python3 web_to_excel.py
# Custom page / table
python3 web_to_excel.py "https://example.com/page" output/result.xlsx --table-index 0
```

## Example client work this maps to

- Recurring collection of public tables (price lists, directories) for
  competitor / market research
- Turning public statistics and open data into Excel for reporting
- Periodic monitoring of prices or listings (the metadata sheet keeps history auditable)
- Full pipelines: collect → clean → monthly report

Scope: public pages only. Sites that require login, or whose terms
prohibit collection, are out of scope (verified before any real project).

## Technical notes

- Python / requests + BeautifulSoup + openpyxl
- Prefers tables with `class="wikitable"` (Wikipedia style); falls back
  to tables in page order
- Cells with colspan are expanded so the grid shape is preserved
- Number-like cells are converted (`1,234` → `1234`); percentages are
  kept as text
- The workbook is presentation-ready: styled header, borders,
  auto column widths, landscape fit-to-page print setup

## About the sample data

The source data is from Wikipedia (CC BY-SA). The output file is a
sample produced by running this tool.
