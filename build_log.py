#!/usr/bin/env python3
"""Build Production_Release_Log.xlsx from release_state.json."""
import json
import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).parent
STATE = ROOT / "release_state.json"
OUT = ROOT / "Production_Release_Log.xlsx"

HEADERS = [
    "Job #",
    "Job Name",
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

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
NEW_FILL = PatternFill("solid", fgColor="FFF2CC")
OLDEST_FILL = PatternFill("solid", fgColor="F8CBAD")
FLAG_FILL = PatternFill("solid", fgColor="FCE4D6")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def parse_date(s):
    return datetime.strptime(s, "%Y-%m-%d").date()


def build(prev_state_path=None):
    state = json.loads(STATE.read_text())
    releases = state["releases"]

    prev_ids = set()
    if prev_state_path and Path(prev_state_path).exists():
        prev = json.loads(Path(prev_state_path).read_text())
        prev_ids = {r["messageId"] for r in prev.get("releases", [])}

    today = date.today()
    if not releases:
        oldest_id = None
    else:
        oldest = min(releases, key=lambda r: r["received"])
        oldest_id = oldest["messageId"]

    wb = Workbook()
    ws = wb.active
    ws.title = "Pending Releases"

    for col_idx, h in enumerate(HEADERS, 1):
        c = ws.cell(row=1, column=col_idx, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[1].height = 26
    ws.freeze_panes = "A2"

    releases_sorted = sorted(releases, key=lambda r: r["received"])

    for i, r in enumerate(releases_sorted, start=2):
        recv = parse_date(r["received"])
        days = (today - recv).days
        row_vals = [
            r["jobNumber"],
            r["name"],
            r["contractor"],
            r["pm"],
            r["contact"],
            r["phone"],
            r["storm"],
            r["san"],
            recv,
            days,
            r["notes"],
        ]
        for col_idx, v in enumerate(row_vals, 1):
            c = ws.cell(row=i, column=col_idx, value=v)
            c.alignment = Alignment(vertical="top", wrap_text=True)
            c.border = BORDER
            if col_idx == 9:
                c.number_format = "yyyy-mm-dd"
            if col_idx == 10:
                c.alignment = Alignment(horizontal="center", vertical="top")

        is_new = r["messageId"] not in prev_ids and bool(prev_ids)
        is_flag = str(r["jobNumber"]).startswith("TBD") or "FLAG" in r.get("notes", "")

        fill = None
        if r["messageId"] == oldest_id:
            fill = OLDEST_FILL
        elif is_new:
            fill = NEW_FILL
        elif is_flag:
            fill = FLAG_FILL

        if fill is not None:
            for col_idx in range(1, len(HEADERS) + 1):
                ws.cell(row=i, column=col_idx).fill = fill

    widths = [12, 34, 28, 18, 26, 26, 10, 10, 12, 10, 60]
    for idx, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = w

    legend_row = len(releases_sorted) + 4
    ws.cell(row=legend_row, column=1, value="Legend:").font = Font(bold=True)
    ws.cell(row=legend_row, column=2, value="Oldest pending").fill = OLDEST_FILL
    ws.cell(row=legend_row, column=3, value="New since last run").fill = NEW_FILL
    ws.cell(row=legend_row, column=4, value="Flagged / missing job#").fill = FLAG_FILL

    meta_row = legend_row + 2
    ws.cell(row=meta_row, column=1, value="Run:").font = Font(bold=True)
    ws.cell(row=meta_row, column=2, value=state.get("lastRun", ""))
    ws.cell(row=meta_row + 1, column=1, value="Total pending:").font = Font(bold=True)
    ws.cell(row=meta_row + 1, column=2, value=len(releases_sorted))

    wb.save(OUT)
    print(f"Wrote {OUT} ({len(releases_sorted)} rows).")


if __name__ == "__main__":
    prev = sys.argv[1] if len(sys.argv) > 1 else None
    build(prev)
