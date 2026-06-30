#!/usr/bin/env python3
"""Rebuild Production_Release_Log.xlsx from release_state.json.

Layout (columns A..K):
  Job # | Project | Contractor | PM | Contact | Phone | Storm | Sanitary | Received | Days Pending | Notes

Row 1 is the header. Days Pending is a real formula =TODAY()-I{row}. recalc.py
re-evaluates that formula and writes the result back as a number so the file has
zero formula errors when opened.
"""
from __future__ import annotations

import json
import os
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STATE_FILE = os.path.join(ROOT, "release_state.json")
OUT_FILE = os.path.join(ROOT, "Production_Release_Log.xlsx")

HEADERS = [
    "Job #",
    "Project",
    "Contractor",
    "PM",
    "Contact",
    "Phone",
    "Storm",
    "Sanitary",
    "Received",
    "Days Pending",
    "Notes",
]


def build() -> None:
    with open(STATE_FILE, "r") as f:
        state = json.load(f)

    releases = sorted(
        state["releases"],
        key=lambda r: (r.get("received") or "9999-99-99"),
    )

    wb = Workbook()
    ws = wb.active
    ws.title = "Pending Releases"

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(border_style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap = Alignment(vertical="top", wrap_text=True)

    for col_idx, name in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = border

    flag_fill = PatternFill("solid", fgColor="FFF2CC")

    for r_idx, rel in enumerate(releases, start=2):
        values = [
            rel.get("jobNumber") or "",
            rel.get("name") or "",
            rel.get("contractor") or "",
            rel.get("pm") or "",
            rel.get("contact") or "",
            rel.get("phone") or "",
            rel.get("storm") or "",
            rel.get("san") or "",
            rel.get("received") or "",
            None,
            rel.get("notes") or "",
        ]
        for c_idx, v in enumerate(values, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=v)
            cell.border = border
            cell.alignment = wrap
        ws.cell(row=r_idx, column=10).value = f"=IFERROR(TODAY()-I{r_idx},\"\")"
        if "FLAG" in (rel.get("notes") or "").upper():
            for c_idx in range(1, len(HEADERS) + 1):
                ws.cell(row=r_idx, column=c_idx).fill = flag_fill

    widths = [10, 38, 28, 18, 22, 16, 8, 10, 12, 10, 50]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 26
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = (
        f"A1:{get_column_letter(len(HEADERS))}{len(releases) + 1}"
    )

    summary = wb.create_sheet("Summary")
    summary["A1"] = "Production Release Log"
    summary["A1"].font = Font(bold=True, size=14)
    summary["A3"] = "Run date"
    summary["B3"] = date.today().isoformat()
    summary["A4"] = "Total pending"
    summary["B4"] = len(releases)
    pm_counts: dict[str, int] = {}
    for r in releases:
        pm_counts[r.get("pm") or "?"] = pm_counts.get(r.get("pm") or "?", 0) + 1
    summary["A6"] = "By PM"
    summary["A6"].font = Font(bold=True)
    for i, (pm, n) in enumerate(sorted(pm_counts.items()), start=7):
        summary.cell(row=i, column=1, value=pm)
        summary.cell(row=i, column=2, value=n)
    summary.column_dimensions["A"].width = 28
    summary.column_dimensions["B"].width = 12

    wb.save(OUT_FILE)
    print(f"Wrote {OUT_FILE} with {len(releases)} releases")


if __name__ == "__main__":
    build()
