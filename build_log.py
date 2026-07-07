#!/usr/bin/env python3
"""Build Production_Release_Log.xlsx from release_state.json.

Rebuilds the log every run — no formulas, no leftover error cells.
"""
import json
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
STATE = ROOT / "release_state.json"
OUT = ROOT / "Production_Release_Log.xlsx"

HEADERS = [
    "Job #",
    "Job Name",
    "Contractor",
    "PM",
    "Contact",
    "Phone / Email",
    "Storm",
    "Sanitary",
    "Received",
    "Days Pending",
    "Notes",
]

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
FLAG_FILL   = PatternFill("solid", fgColor="FFF2CC")
OLDEST_FILL = PatternFill("solid", fgColor="FCE4D6")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def parse_iso(d):
    try:
        return datetime.strptime(d, "%Y-%m-%d").date()
    except (TypeError, ValueError):
        return None


def main() -> None:
    data = json.loads(STATE.read_text())
    releases = data.get("releases", [])

    releases.sort(key=lambda r: r.get("received", ""))

    today = date.today()
    oldest_received = None
    for r in releases:
        d = parse_iso(r.get("received"))
        if d and (oldest_received is None or d < oldest_received):
            oldest_received = d

    wb = Workbook()
    ws = wb.active
    ws.title = "Pending Releases"

    ws.append(HEADERS)
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER

    for r in releases:
        received = parse_iso(r.get("received"))
        days_pending = (today - received).days if received else ""
        row = [
            r.get("jobNumber", ""),
            r.get("name", ""),
            r.get("contractor", ""),
            r.get("pm", ""),
            r.get("contact", ""),
            r.get("phone", ""),
            r.get("storm", ""),
            r.get("san", ""),
            r.get("received", ""),
            days_pending,
            r.get("notes", ""),
        ]
        ws.append(row)
        row_idx = ws.max_row

        is_flagged = "FLAGGED" in (r.get("notes") or "")
        is_oldest = received is not None and received == oldest_received

        for cell in ws[row_idx]:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = BORDER
            if is_oldest:
                cell.fill = OLDEST_FILL
            elif is_flagged:
                cell.fill = FLAG_FILL

    widths = [10, 34, 26, 18, 22, 32, 12, 12, 12, 8, 60]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"

    summary_row = ws.max_row + 2
    ws.cell(row=summary_row, column=1, value="Total pending:").font = Font(bold=True)
    ws.cell(row=summary_row, column=2, value=len(releases)).font = Font(bold=True)
    ws.cell(row=summary_row + 1, column=1, value="Generated:").font = Font(bold=True)
    ws.cell(row=summary_row + 1, column=2, value=data.get("generatedAt", today.isoformat()))

    wb.save(OUT)
    print(f"Wrote {OUT} — {len(releases)} rows")


if __name__ == "__main__":
    main()
