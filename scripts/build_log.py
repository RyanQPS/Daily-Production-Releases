#!/usr/bin/env python3
"""Rebuild Production_Release_Log.xlsx from state/release_state.json.

Zero formulas — all values written directly so the workbook opens with no
#VALUE!/#REF!/#NAME? errors regardless of Excel version.

Usage:  python scripts/build_log.py
Reads:  state/release_state.json
Writes: output/Production_Release_Log.xlsx
"""
from __future__ import annotations

import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
STATE = ROOT / "state" / "release_state.json"
OUT = ROOT / "output" / "Production_Release_Log.xlsx"

NEW_CUTOFF = "2026-06-24"  # rows received on/after this date are highlighted NEW
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
NEW_FILL = PatternFill("solid", fgColor="E2EFDA")
STALE_FILL = PatternFill("solid", fgColor="FFE0E0")
ALT_FILL = PatternFill("solid", fgColor="DCE6F1")
THIN = Side(border_style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

HEADERS = ["#", "Job #", "Job Name", "Contractor", "PM", "Contact",
           "Phone", "Storm", "San", "Total", "Received", "Notes"]
WIDTHS = [4, 10, 32, 24, 16, 18, 16, 6, 6, 10, 12, 30]


def _int_sum(rows, key):
    return sum(r[key] for r in rows if isinstance(r[key], int))


def build(state: dict) -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "Pending Releases"

    for c, h in enumerate(HEADERS, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for i, row in enumerate(state["pending"], start=2):
        values = [row["n"], row["jobNumber"], row["name"], row["contractor"],
                  row["pm"], row["contact"], row["phone"], row["storm"],
                  row["san"], row["tot"], row["received"], row.get("notes", "")]
        is_stale = row.get("stale", False)
        is_new = row["received"] >= NEW_CUTOFF and not is_stale
        for c, v in enumerate(values, 1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if is_stale:
                cell.fill = STALE_FILL
                cell.font = Font(bold=True)
            elif is_new:
                cell.fill = NEW_FILL
                cell.font = Font(bold=True)
            elif i % 2 == 0:
                cell.fill = ALT_FILL

    tot_row = len(state["pending"]) + 2
    ws.cell(row=tot_row, column=1, value="TOTAL").font = Font(bold=True)
    ws.cell(row=tot_row, column=8, value=_int_sum(state["pending"], "storm")).font = Font(bold=True)
    ws.cell(row=tot_row, column=9, value=_int_sum(state["pending"], "san")).font = Font(bold=True)
    ws.cell(row=tot_row, column=10, value=f"{len(state['pending'])} pending").font = Font(bold=True)

    for i, w in enumerate(WIDTHS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"

    notes = wb.create_sheet("Notes & Log")
    lines = [
        [f"Production Release Log — generated {state['generated']}"],
        [],
        ["Total pending", len(state["pending"])],
        ["Storm total", _int_sum(state["pending"], "storm")],
        ["San total", _int_sum(state["pending"], "san")],
    ]
    for r, line in enumerate(lines, start=1):
        for c, v in enumerate(line, start=1):
            cell = notes.cell(row=r, column=c, value=v)
            if r == 1:
                cell.font = Font(bold=True, size=14)
    notes.column_dimensions["A"].width = 55
    notes.column_dimensions["B"].width = 40
    return wb


def main() -> None:
    state = json.loads(STATE.read_text())
    OUT.parent.mkdir(exist_ok=True)
    build(state).save(OUT)
    print(f"Wrote {OUT} ({OUT.stat().st_size} bytes, {len(state['pending'])} pending)")


if __name__ == "__main__":
    main()
