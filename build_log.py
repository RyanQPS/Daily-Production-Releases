#!/usr/bin/env python3
"""Rebuild Production_Release_Log.xlsx from release_state.json.

Writes plain values (no formulas) so the workbook opens with zero formula errors.
"""
import json
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).parent
STATE = ROOT / "release_state.json"
OUT = ROOT / "Production_Release_Log.xlsx"

HEADERS = [
    "Job #", "Job Name", "Contractor", "PM", "Contact", "Phone",
    "Storm", "San", "Received", "Notes", "Subject",
]

FIELDS = [
    "jobNumber", "name", "contractor", "pm", "contact", "phone",
    "storm", "san", "received", "notes", "subject",
]


def build():
    data = json.loads(STATE.read_text())
    releases = data.get("releases", [])
    releases_sorted = sorted(releases, key=lambda r: r.get("received", ""))

    wb = Workbook()
    ws = wb.active
    ws.title = "Pending Releases"

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(bold=True, color="FFFFFF")
    thin = Side(border_style="thin", color="B0B0B0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    flag_fill = PatternFill("solid", fgColor="FFF2CC")

    for col_idx, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    for row_idx, item in enumerate(releases_sorted, start=2):
        for col_idx, field in enumerate(FIELDS, start=1):
            value = item.get(field, "") or ""
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = border
            if "FLAGGED" in str(item.get("notes", "")) or "AMBIGUOUS" in str(item.get("notes", "")) or "OLDEST" in str(item.get("notes", "")):
                cell.fill = flag_fill

    widths = [10, 34, 22, 18, 26, 22, 8, 8, 12, 46, 60]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}{len(releases_sorted) + 1}"

    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = "Production Release Log — Summary"
    ws2["A1"].font = Font(bold=True, size=14)
    ws2["A3"] = "Total pending"
    ws2["B3"] = len(releases_sorted)
    ws2["A4"] = "By PM"
    ws2["A4"].font = Font(bold=True)

    by_pm = {}
    for r in releases_sorted:
        by_pm[r.get("pm", "?")] = by_pm.get(r.get("pm", "?"), 0) + 1
    row = 5
    for pm, count in sorted(by_pm.items()):
        ws2.cell(row=row, column=1, value=pm)
        ws2.cell(row=row, column=2, value=count)
        row += 1

    ws2.column_dimensions["A"].width = 26
    ws2.column_dimensions["B"].width = 10

    wb.save(OUT)
    print(f"Wrote {OUT} — {len(releases_sorted)} rows")


if __name__ == "__main__":
    build()
