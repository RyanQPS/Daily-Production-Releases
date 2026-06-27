#!/usr/bin/env python3
"""Build Production_Release_Log.xlsx from release_state.json.

All cell values are pre-computed in Python. No spreadsheet formulas are
written, which is what the operating-instructions "ZERO formula errors"
requirement demands.
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

HEADERS = [
    "Job #",
    "Project Name",
    "Contractor",
    "PM",
    "Contact",
    "Phone",
    "Storm",
    "San",
    "Structures",
    "Received",
    "Days Pending",
    "Notes",
]

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="top")
NEW_FILL = PatternFill("solid", fgColor="FFF2CC")  # highlight for new-since-last-run
OLD_FILL = PatternFill("solid", fgColor="F8CBAD")  # highlight for oldest


def days_pending(received: str, today: date) -> int:
    try:
        return (today - datetime.strptime(received, "%Y-%m-%d").date()).days
    except ValueError:
        return -1


def main() -> None:
    state = json.loads(STATE_PATH.read_text())
    releases = state["releases"]
    last_run = state["metadata"].get("lastRun", "")
    today = (
        datetime.fromisoformat(last_run).date()
        if last_run
        else date.today()
    )

    # Sort: oldest first (so the most-aged jobs are at the top).
    releases_sorted = sorted(releases, key=lambda r: r["received"])
    oldest_job = releases_sorted[0]["jobNumber"] if releases_sorted else None

    # Day-1: everything is "new since last run".
    new_jobs = {r["jobNumber"] for r in releases}

    wb = Workbook()
    ws = wb.active
    ws.title = "Pending Releases"

    # Header
    for col, label in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col, value=label)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = BORDER

    # Rows
    for row_idx, rel in enumerate(releases_sorted, start=2):
        dp = days_pending(rel["received"], today)
        values = [
            rel.get("jobNumber", ""),
            rel.get("name", ""),
            rel.get("contractor", ""),
            rel.get("pm", ""),
            rel.get("contact", ""),
            rel.get("phone", ""),
            rel.get("storm", ""),
            rel.get("san", ""),
            rel.get("structures", ""),
            rel.get("received", ""),
            dp if dp >= 0 else "",
            rel.get("notes", ""),
        ]
        for col, val in enumerate(values, start=1):
            c = ws.cell(row=row_idx, column=col, value=val)
            c.border = BORDER
            c.alignment = WRAP if col in (2, 12) else CENTER

        # Highlight new-since-last-run (day 1: all rows).
        if rel["jobNumber"] in new_jobs:
            for col in range(1, len(HEADERS) + 1):
                ws.cell(row=row_idx, column=col).fill = NEW_FILL

        # Highlight oldest pending separately.
        if rel["jobNumber"] == oldest_job:
            for col in range(1, len(HEADERS) + 1):
                ws.cell(row=row_idx, column=col).fill = OLD_FILL

    # Column widths
    widths = [12, 38, 28, 18, 30, 22, 6, 6, 14, 12, 8, 60]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"

    # Summary sheet
    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = "Production Release Log — Summary"
    ws2["A1"].font = Font(bold=True, size=14)
    ws2["A3"] = "Generated"
    ws2["B3"] = today.isoformat()
    ws2["A4"] = "Total pending"
    ws2["B4"] = len(releases)
    ws2["A5"] = "New since last run"
    ws2["B5"] = len(new_jobs)
    ws2["A6"] = "Oldest job"
    ws2["B6"] = oldest_job or ""
    ws2["A7"] = "Oldest received"
    ws2["B7"] = releases_sorted[0]["received"] if releases_sorted else ""

    # Per-PM tally
    ws2["A9"] = "By PM"
    ws2["A9"].font = Font(bold=True)
    pm_counts: dict[str, int] = {}
    for r in releases:
        pm_counts[r.get("pm", "?")] = pm_counts.get(r.get("pm", "?"), 0) + 1
    row = 10
    for pm, n in sorted(pm_counts.items()):
        ws2.cell(row=row, column=1, value=pm)
        ws2.cell(row=row, column=2, value=n)
        row += 1

    ws2.column_dimensions["A"].width = 28
    ws2.column_dimensions["B"].width = 18

    wb.save(OUT_PATH)
    print(f"Wrote {OUT_PATH} ({len(releases)} rows).")


if __name__ == "__main__":
    main()
