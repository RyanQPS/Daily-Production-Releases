"""
build_log.py — Rebuild Production_Release_Log.xlsx from release_state.json.

The log is a flat, formula-free workbook so no #REF!/#NAME? errors can appear
after recalc.py runs. Ordering: oldest received date first (oldest to fresh),
because Ryan wants the oldest job flagged at the top.
"""
from __future__ import annotations

import json
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).parent
STATE = ROOT / "release_state.json"
OUT = ROOT / "Production_Release_Log.xlsx"

COLUMNS = [
    ("jobNumber",  "Job #",       12),
    ("name",       "Job Name",    38),
    ("contractor", "Contractor",  26),
    ("pm",         "PM",          16),
    ("contact",    "Contact",     28),
    ("phone",      "Phone",       26),
    ("storm",      "Storm",       11),
    ("san",        "San",         11),
    ("received",   "Received",    12),
    ("notes",      "Notes",       60),
]

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
OLDEST_FILL = PatternFill("solid", fgColor="FFD966")   # amber — oldest
FLAG_FILL   = PatternFill("solid", fgColor="F4B084")   # peach — ambiguous
WRAP = Alignment(wrap_text=True, vertical="top")


def parse_received(row: dict) -> date:
    s = row.get("received", "") or ""
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except ValueError:
        return date.max


def main() -> None:
    state = json.loads(STATE.read_text())
    releases = list(state.get("releases", []))
    releases.sort(key=parse_received)

    oldest_idx = 0 if releases else -1

    wb = Workbook()
    ws = wb.active
    ws.title = "Pending Releases"

    for col, (_key, label, width) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col, value=label)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"

    for r, row in enumerate(releases, start=2):
        is_flagged = "FLAG" in (row.get("notes") or "").upper()
        is_oldest = (r - 2) == oldest_idx
        fill = OLDEST_FILL if is_oldest else (FLAG_FILL if is_flagged else None)
        for c, (key, _label, _w) in enumerate(COLUMNS, start=1):
            cell = ws.cell(row=r, column=c, value=row.get(key, ""))
            cell.alignment = WRAP
            if fill:
                cell.fill = fill

    meta = wb.create_sheet("Run Info")
    meta["A1"] = "Field"; meta["B1"] = "Value"
    meta["A1"].font = Font(bold=True); meta["B1"].font = Font(bold=True)
    info = [
        ("Last run (UTC)",   state.get("lastRunISO", "")),
        ("Last run (local)", state.get("lastRunLocal", "")),
        ("Pending count",    len(releases)),
        ("Notes",            state.get("notes", "")),
    ]
    for i, (k, v) in enumerate(info, start=2):
        meta.cell(row=i, column=1, value=k).font = Font(bold=True)
        meta.cell(row=i, column=2, value=v).alignment = WRAP
    meta.column_dimensions["A"].width = 22
    meta.column_dimensions["B"].width = 60

    wb.save(OUT)
    print(f"Wrote {OUT} — {len(releases)} pending rows.")


if __name__ == "__main__":
    main()
