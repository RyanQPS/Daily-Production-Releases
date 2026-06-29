#!/usr/bin/env python3
"""Rebuild Production_Release_Log.xlsx from release_state.json.

Writes computed values (no formulas) so there are zero formula errors.
"""
from __future__ import annotations

import json
import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

REPO = Path(__file__).resolve().parent
STATE = REPO / "release_state.json"
OUT = REPO / "Production_Release_Log.xlsx"

HEADERS = [
    "Job #", "Job Name", "Contractor", "PM",
    "Contact", "Phone", "Storm", "SAN",
    "Received", "Days Pending", "New This Run", "Flag", "Notes",
]

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
NEW_FILL = PatternFill("solid", fgColor="FFF2CC")
FLAG_FILL = PatternFill("solid", fgColor="FCE4D6")
OLD_FILL = PatternFill("solid", fgColor="E2EFDA")


def days_between(received: str, run_date: date) -> int:
    try:
        d = datetime.strptime(received, "%Y-%m-%d").date()
    except ValueError:
        return 0
    return max(0, (run_date - d).days)


def main() -> int:
    state = json.loads(STATE.read_text())
    releases = state["releases"]
    run_date = datetime.strptime(state["lastRunDate"], "%Y-%m-%d").date()
    prev_ids = set(state.get("previousReleaseIds", []))

    releases_sorted = sorted(releases, key=lambda r: r["received"])
    oldest_received = releases_sorted[0]["received"] if releases_sorted else None

    wb = Workbook()
    ws = wb.active
    ws.title = "Pending Releases"

    ws.append(HEADERS)
    for col_idx, _ in enumerate(HEADERS, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER

    for r in releases_sorted:
        is_new = r["messageId"] not in prev_ids
        is_flag = bool(r.get("flagged"))
        days = days_between(r["received"], run_date)
        row = [
            r["jobNumber"],
            r["name"],
            r["contractor"],
            r["pm"],
            r.get("contact", ""),
            r.get("phone", ""),
            "Yes" if r.get("storm") else "",
            "Yes" if r.get("san") else "",
            r["received"],
            days,
            "NEW" if is_new else "",
            r.get("flagReason", "") if is_flag else "",
            r.get("notes", ""),
        ]
        ws.append(row)
        row_idx = ws.max_row
        fill = None
        if r["received"] == oldest_received:
            fill = OLD_FILL
        elif is_flag:
            fill = FLAG_FILL
        elif is_new:
            fill = NEW_FILL
        for col_idx in range(1, len(HEADERS) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if fill:
                cell.fill = fill

    # Summary row (computed, not formula)
    total = len(releases_sorted)
    new_count = sum(1 for r in releases_sorted if r["messageId"] not in prev_ids)
    flag_count = sum(1 for r in releases_sorted if r.get("flagged"))
    storm_count = sum(1 for r in releases_sorted if r.get("storm"))
    san_count = sum(1 for r in releases_sorted if r.get("san"))
    ws.append([])
    summary_row = ws.max_row + 1
    ws.cell(row=summary_row, column=1, value="TOTALS").font = Font(bold=True)
    ws.cell(row=summary_row, column=2, value=f"Pending: {total}")
    ws.cell(row=summary_row, column=3, value=f"New this run: {new_count}")
    ws.cell(row=summary_row, column=4, value=f"Flagged: {flag_count}")
    ws.cell(row=summary_row, column=7, value=f"Storm: {storm_count}")
    ws.cell(row=summary_row, column=8, value=f"SAN: {san_count}")
    if oldest_received:
        ws.cell(row=summary_row, column=9, value=f"Oldest: {oldest_received}")

    # Column widths
    widths = [11, 36, 30, 18, 22, 28, 7, 7, 12, 8, 10, 32, 44]
    for col_idx, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = w

    ws.freeze_panes = "A2"

    # Run metadata sheet
    meta = wb.create_sheet("Run Info")
    meta.append(["Field", "Value"])
    meta.append(["Run date", state["lastRunDate"]])
    meta.append(["Run timestamp", state["lastRunIso"]])
    meta.append(["Previous run", state.get("previousRunDate") or "(first run / bootstrap)"])
    meta.append(["Total pending", total])
    meta.append(["New this run", new_count])
    meta.append(["Flagged", flag_count])
    meta.append(["Oldest received", oldest_received or ""])
    meta.column_dimensions["A"].width = 22
    meta.column_dimensions["B"].width = 48
    for cell in meta[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL

    wb.save(OUT)
    print(f"Wrote {OUT} — {total} rows ({new_count} new, {flag_count} flagged)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
