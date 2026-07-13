#!/usr/bin/env python3
"""Rebuild Production_Release_Log.xlsx from release_state.json.

Zero-formula-error output: all values are literals; no formulas used.
Highlights new-since-last-run rows by shading the Received column.
"""
from __future__ import annotations

import json
import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


REPO = Path(__file__).resolve().parent
STATE_PATH = REPO / "release_state.json"
OUT_PATH = REPO / "Production_Release_Log.xlsx"

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
NEW_FILL = PatternFill("solid", fgColor="FFF2CC")
FLAG_FILL = PatternFill("solid", fgColor="FCE4D6")
BORDER = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="top", wrap_text=True)

COLUMNS = [
    ("Job #", "jobNumber", 14),
    ("Project Name", "name", 34),
    ("Contractor", "contractor", 22),
    ("PM", "pm", 20),
    ("Contact", "contact", 26),
    ("Phone", "phone", 22),
    ("Storm", "storm", 10),
    ("San", "san", 10),
    ("Lift", "lift", 12),
    ("Received", "received", 14),
    ("Age (days)", "age", 10),
    ("Notes", "notes", 60),
]


def age_days(received: str, today: date) -> int:
    try:
        d = datetime.strptime(received, "%Y-%m-%d").date()
    except ValueError:
        return -1
    return (today - d).days


def build(today: date | None = None) -> Path:
    state = json.loads(STATE_PATH.read_text())
    entries = state["entries"]
    new_ids = {e["messageId"] for e in state.get("newSinceLastRun", []) or []}
    flagged_jobs = {f["job"] for f in state.get("flags", []) or []}
    today = today or date.today()

    wb = Workbook()
    ws = wb.active
    ws.title = "Pending RTPs"

    for i, (title, _, width) in enumerate(COLUMNS, 1):
        cell = ws.cell(row=1, column=i, value=title)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"

    entries_sorted = sorted(entries, key=lambda e: (e["received"], e["jobNumber"]))
    for r_idx, e in enumerate(entries_sorted, start=2):
        e_view = dict(e)
        e_view["age"] = age_days(e["received"], today)
        is_new = e["messageId"] in new_ids
        is_flagged = e["jobNumber"] in flagged_jobs
        for c_idx, (_, key, _) in enumerate(COLUMNS, 1):
            val = e_view.get(key, "")
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.alignment = LEFT
            cell.border = BORDER
            if is_new and key == "received":
                cell.fill = NEW_FILL
            elif is_flagged and key == "jobNumber":
                cell.fill = FLAG_FILL

    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = "Production Release Log — Summary"
    ws2["A1"].font = Font(bold=True, size=14)
    ws2["A3"] = "Generated at"
    ws2["B3"] = state.get("generatedAt", "")
    ws2["A4"] = "Prior run"
    ws2["B4"] = state.get("priorRunEmail", "")
    ws2["A5"] = "Log rows"
    ws2["B5"] = state["totals"]["logRows"]
    ws2["A6"] = "Unique jobs"
    ws2["B6"] = state["totals"]["uniqueJobs"]

    ws2["A8"] = "Per-PM totals"
    ws2["A8"].font = Font(bold=True)
    r = 9
    for pm, n in state["perPmTotals"].items():
        ws2.cell(row=r, column=1, value=pm)
        ws2.cell(row=r, column=2, value=n)
        r += 1

    ws2.column_dimensions["A"].width = 26
    ws2.column_dimensions["B"].width = 42

    r += 1
    ws2.cell(row=r, column=1, value="Flags / notes").font = Font(bold=True)
    r += 1
    for f in state.get("flags", []) or []:
        ws2.cell(row=r, column=1, value=f["job"])
        ws2.cell(row=r, column=2, value=f"{f['type']}: {f['detail']}")
        ws2.cell(row=r, column=2).alignment = LEFT
        r += 1

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT_PATH)
    return OUT_PATH


if __name__ == "__main__":
    out = build()
    print(f"wrote {out}", file=sys.stderr)
