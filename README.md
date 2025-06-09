# macauDay
Given a PDF containing immigration records marked with the keyword **出境 yyyy-mm-dd**, this script counts outbound events per academic year (Aug 1 – Jul 31), removes duplicate trips on the same day, optionally excludes Macau public holidays, and outputs two files in the same directory as the input PDF
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

适配移民局小程序下载的出入境pdf，自动统计出境次数
