#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outbound statistics generator (academic-year adaptive)

[功能增强] 在输出文件中添加总计行
[口径调整] 统计“境外停留天数”，而不是“出境次数”
[新增列] 输出“去除周末和假期后天数”
"""

import sys
import re
from pathlib import Path
import datetime as dt

import pandas as pd
import PyPDF2

# ----------------------------------------------------------------------
# Holiday handling: use the 'holidays' library (supports Macau, code='MO').
# Covers 2020-2030 automatically.
try:
    import holidays  # type: ignore
    MO_HOLIDAYS = holidays.country_holidays("MO", years=range(2020, 2031))
except ImportError:
    print("⚠️  The 'holidays' package is not installed; holiday exclusion will be skipped.")
    MO_HOLIDAYS = None

# ----------------------------------------------------------------------
def extract_inout_events(pdf_path: Path) -> list[tuple[str, dt.date]]:
    """Return ordered list of ('出境'/'入境', date) records from PDF."""
    with pdf_path.open("rb") as f:
        reader = PyPDF2.PdfReader(f)
        text = "\n".join(page.extract_text() or "" for page in reader.pages)

    pattern = re.compile(r"(出境|入境)\s*([0-9]{4}-[0-9]{2}-[0-9]{2})")
    return [(m.group(1), dt.date.fromisoformat(m.group(2))) for m in pattern.finditer(text)]


def normalize_event_order(events: list[tuple[str, dt.date]]) -> list[tuple[str, dt.date]]:
    """
    Normalize event order.
    Many PDFs list records in reverse chronological order; if so, reverse them.
    """
    if len(events) <= 1:
        return events

    asc_count = 0
    desc_count = 0
    for i in range(len(events) - 1):
        if events[i][1] <= events[i + 1][1]:
            asc_count += 1
        if events[i][1] >= events[i + 1][1]:
            desc_count += 1

    if desc_count > asc_count:
        return list(reversed(events))
    return events


def expand_stay_dates(events: list[tuple[str, dt.date]]) -> list[dt.date]:
    """
    Convert ordered 出境/入境 events into a list of stay dates (inclusive).

    Example:
      出境 2025-03-01
      入境 2025-03-02
    => [2025-03-01, 2025-03-02]
    """
    events = normalize_event_order(events)

    stay_dates: list[dt.date] = []
    current_exit = None

    for event_type, event_date in events:
        if event_type == "出境":
            current_exit = event_date
        elif event_type == "入境":
            if current_exit is not None and event_date >= current_exit:
                days = (event_date - current_exit).days + 1
                for i in range(days):
                    stay_dates.append(current_exit + dt.timedelta(days=i))
                current_exit = None
            else:
                # 无匹配出境，或日期异常，直接跳过
                current_exit = None

    return stay_dates


def build_academic_years(dates: list[dt.date]) -> list[tuple[dt.date, dt.date, str]]:
    """Generate academic-year spans (start, end, label) covering the date range."""
    if not dates:
        return []
    first_year = dates[0].year if dates[0].month >= 9 else dates[0].year - 1
    last_year = dates[-1].year if dates[-1].month >= 9 else dates[-1].year - 1
    spans = []
    for year in range(first_year, last_year + 1):
        start = dt.date(year, 9, 1)
        end = dt.date(year + 1, 8, 31)
        label = f"{str(year)[-2:]}-{str(year + 1)[-2:]}学年 ({start}~{end})"
        spans.append((start, end, label))
    return spans


def compute_stats(dates: list[dt.date], start: dt.date, end: dt.date) -> tuple[int, int, int, int]:
    """
    Compute:
    (total, unique_days, non_holiday_unique_days, non_weekend_holiday_unique_days)
    """
    within = [d for d in dates if start <= d <= end]
    total = len(within)
    unique_days = set(within)

    if MO_HOLIDAYS is not None:
        non_holiday_days = {d for d in unique_days if d not in MO_HOLIDAYS}
    else:
        non_holiday_days = unique_days

    # 去除周末和假期
    if MO_HOLIDAYS is not None:
        non_weekend_holiday_days = {
            d for d in unique_days
            if d.weekday() < 5 and d not in MO_HOLIDAYS
        }
    else:
        non_weekend_holiday_days = {
            d for d in unique_days
            if d.weekday() < 5
        }

    return total, len(unique_days), len(non_holiday_days), len(non_weekend_holiday_days)


def generate_markdown(df: pd.DataFrame) -> str:
    lines = [
        "| 学年 | 境外停留总天数 | 单日去重后天数 | 去除假期后天数 | 去除周末和假期后天数 |",
        "|------|----------------|----------------|----------------|----------------------|"
    ]

    for _, row in df.iterrows():
        lines.append(
            f"| {row['学年']} | {row['境外停留总天数']} | {row['单日去重后天数']} | "
            f"{row['去除假期后天数']} | {row['去除周末和假期后天数']} |"
        )

    total_row = df.sum(numeric_only=True)
    lines.append(
        f"| **总计** | **{int(total_row['境外停留总天数'])}** | "
        f"**{int(total_row['单日去重后天数'])}** | "
        f"**{int(total_row['去除假期后天数'])}** | "
        f"**{int(total_row['去除周末和假期后天数'])}** |"
    )

    return "\n".join(lines)


def main() -> None:
    if len(sys.argv) != 2:
        print("Usage: python outbound_stats_auto.py <input.pdf>")
        sys.exit(1)

    pdf_path = Path(sys.argv[1]).expanduser()
    if not pdf_path.is_file():
        print(f"Error: {pdf_path} not found.")
        sys.exit(2)

    events = extract_inout_events(pdf_path)
    if not events:
        print("⚠️  No outbound/inbound records found in the PDF.")
        sys.exit(0)

    dates = sorted(expand_stay_dates(events))
    if not dates:
        print("⚠️  No valid stay-day records found in the PDF.")
        sys.exit(0)

    spans = build_academic_years(dates)
    stats_rows = [
        (label, *compute_stats(dates, start, end))
        for start, end, label in spans
    ]
    df = pd.DataFrame(
        stats_rows,
        columns=[
            "学年",
            "境外停留总天数",
            "单日去重后天数",
            "去除假期后天数",
            "去除周末和假期后天数",
        ],
    )

    # 添加总计行到Excel
    total_row = pd.DataFrame({
        "学年": ["总计"],
        "境外停留总天数": [df["境外停留总天数"].sum()],
        "单日去重后天数": [df["单日去重后天数"].sum()],
        "去除假期后天数": [df["去除假期后天数"].sum()],
        "去除周末和假期后天数": [df["去除周末和假期后天数"].sum()],
    })
    df_with_total = pd.concat([df, total_row], ignore_index=True)

    base_path = pdf_path.with_suffix("")
    excel_path = base_path.with_suffix(".xlsx")
    md_path = base_path.with_suffix(".md")

    # 保存Excel（包含总计行）
    df_with_total.to_excel(excel_path, index=False)

    # 保存Markdown（包含总计行）
    md_path.write_text(generate_markdown(df), encoding="utf-8")

    print(f"✅  Results saved: {excel_path} and {md_path}")
    print(
        f"    Total stay days: {df['境外停留总天数'].sum()} (raw), "
        f"{df['单日去重后天数'].sum()} (deduped), "
        f"{df['去除假期后天数'].sum()} (non-holiday), "
        f"{df['去除周末和假期后天数'].sum()} (non-weekend-holiday)"
    )


if __name__ == "__main__":
    main()
