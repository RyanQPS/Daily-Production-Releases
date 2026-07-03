#!/usr/bin/env python3
"""Rebuild Production_Release_Log.xlsx from release_state.json.

No formulas — Days Pending computed at build time so the file opens with zero
formula errors regardless of Excel's calc mode.
"""
from __future__ import annotations

import json
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
STATE_PATH = ROOT / "release_state.json"
OUT_PATH = ROOT / "Production_Release_Log.xlsx"

HEADER = [
    "Job #", "Job Name", "Contractor", "PM", "Contact", "Phone",
    "Storm Str", "San Str", "Received", "Days Pending", "Notes",
]

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(name="Aptos", size=11, bold=True, color="FFFFFF")
NEW_FILL = PatternFill("solid", fgColor="C6EFCE")   # new-this-run rows
OLDEST_FILL = PatternFill("solid", fgColor="FFD966")  # oldest pending
FLAG_FILL = PatternFill("solid", fgColor="F4B084")    # ambiguous / missing job#
BODY_FONT = Font(name="Aptos", size=10)
THIN = Side(border_style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def parse_date(s: str) -> date:
    return datetime.strptime(s, "%Y-%m-%d").date()


def main() -> None:
    state = json.loads(STATE_PATH.read_text())
    run_date = parse_date(state["runDate"])
    releases = sorted(state["releases"], key=lambda r: r["received"])

    oldest_received = releases[0]["received"] if releases else None

    wb = Workbook()
    ws = wb.active
    ws.title = "Pending Releases"

    for col, name in enumerate(HEADER, start=1):
        cell = ws.cell(row=1, column=col, value=name)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER

    for row_idx, r in enumerate(releases, start=2):
        received = parse_date(r["received"])
        days = (run_date - received).days
        values = [
            r["jobNumber"], r["name"], r["contractor"], r["pm"],
            r["contact"], r["phone"], r["storm"], r["san"],
            received, days, r["notes"],
        ]
        fill = None
        if r["received"] == oldest_received:
            fill = OLDEST_FILL
        elif r["jobNumber"] == "See PDF":
            fill = FLAG_FILL
        elif r.get("newThisRun"):
            fill = NEW_FILL

        for col, val in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col, value=val)
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(col == 11))
            if fill is not None:
                cell.fill = fill
            if col == 9:
                cell.number_format = "yyyy-mm-dd"
            if col == 10:
                cell.alignment = Alignment(horizontal="center", vertical="top")

    widths = [10, 34, 26, 18, 22, 16, 10, 10, 12, 10, 60]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADER))}{ws.max_row}"

    wb.save(OUT_PATH)
    print(f"Wrote {OUT_PATH} — {len(releases)} pending releases")


if __name__ == "__main__":
    main()
