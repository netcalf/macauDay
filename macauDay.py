#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outbound statistics generator (academic‑year adaptive)

Given a PDF containing immigration records marked with the keyword **出境 yyyy-mm-dd**,
this script counts outbound events per academic year (Aug 1 – Jul 31), removes
duplicate trips on the same day, optionally excludes Macau public holidays,
and outputs two files in the same directory as the input PDF:

1. <input_basename>.xlsx — Excel workbook with statistics
2. <input_basename>.md   — Markdown table

Dependencies:
    pip install pandas PyPDF2 holidays

Usage:
    python outbound_stats_auto.py <input.pdf>
"""

import sys
import re
from pathlib import Path
import datetime as dt

import pandas as pd
import PyPDF2

# ---------------------------------------------------------------------- 
# Holiday handling: use the 'holidays' library (supports Macau, code='MO').
# Covers 2020‑2030 automatically.
try:
    import holidays  # type: ignore
    MO_HOLIDAYS = holidays.country_holidays("MO", years=range(2020, 2031))
except ImportError:
    print("⚠️  The 'holidays' package is not installed; holiday exclusion will be skipped.")
    MO_HOLIDAYS = None

# ---------------------------------------------------------------------- 
def extract_outbound_dates(pdf_path: Path) -> list[dt.date]:
    """Return list of datetime.date objects for each '出境' record in PDF."""
    with pdf_path.open('rb') as f:
        reader = PyPDF2.PdfReader(f)
        text = "\n".join(page.extract_text() or "" for page in reader.pages)
    pattern = re.compile(r"出境\s*([0-9]{4}-[0-9]{2}-[0-9]{2})")
    return [dt.date.fromisoformat(m.group(1)) for m in pattern.finditer(text)]

def build_academic_years(dates: list[dt.date]) -> list[tuple[dt.date, dt.date, str]]:
    """Generate academic‑year spans (start, end, label) covering the date range."""
    if not dates:
        return []
    first_year = dates[0].year if dates[0].month >= 8 else dates[0].year - 1
    last_year = dates[-1].year if dates[-1].month >= 8 else dates[-1].year - 1
    spans = []
    for year in range(first_year, last_year + 1):
        start = dt.date(year, 8, 1)
        end = dt.date(year + 1, 7, 31)
        label = f"{str(year)[-2:]}-{str(year + 1)[-2:]}学年 ({start}~{end})"
        spans.append((start, end, label))
    return spans

def compute_stats(dates: list[dt.date], start: dt.date, end: dt.date) -> tuple[int, int, int]:
    """Compute (total, unique_days, non_holiday_unique_days) for a given span."""
    within = [d for d in dates if start <= d <= end]
    total = len(within)
    unique_days = set(within)
    if MO_HOLIDAYS is not None:
        non_holiday_days = {d for d in unique_days if d not in MO_HOLIDAYS}
    else:
        non_holiday_days = unique_days
    return total, len(unique_days), len(non_holiday_days)

def generate_markdown(df: pd.DataFrame) -> str:
    lines = [
        "| 学年 | 出境总次数 | 单日去重后次数 | 去除假期后次数 |",
        "|------|-----------|---------------|------------------|"
    ]
    for _, row in df.iterrows():
        lines.append(f"| {row['学年']} | {row['出境总次数']} | {row['单日去重后次数']} | {row['去除假期后次数']} |")
    return "\n".join(lines)

def main() -> None:
    if len(sys.argv) != 2:
        print("Usage: python outbound_stats_auto.py <input.pdf>")
        sys.exit(1)

    pdf_path = Path(sys.argv[1]).expanduser()
    if not pdf_path.is_file():
        print(f"Error: {pdf_path} not found.")
        sys.exit(2)

    dates = sorted(extract_outbound_dates(pdf_path))
    if not dates:
        print("⚠️  No outbound records found in the PDF.")
        sys.exit(0)

    spans = build_academic_years(dates)
    stats_rows = [
        (label, *compute_stats(dates, start, end))
        for start, end, label in spans
    ]
    df = pd.DataFrame(stats_rows, columns=["学年", "出境总次数", "单日去重后次数", "去除假期后次数"])

    base_path = pdf_path.with_suffix("")
    excel_path = base_path.with_suffix(".xlsx")
    md_path = base_path.with_suffix(".md")

    df.to_excel(excel_path, index=False)
    md_path.write_text(generate_markdown(df), encoding="utf-8")

    print(f"✅  Results saved: {excel_path} and {md_path}")

if __name__ == "__main__":
    main()
