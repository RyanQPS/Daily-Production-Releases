#!/usr/bin/env python3
"""Build Production_Release_Log.xlsx from release_state.json.

Usage: python3 build_log.py [state_path] [output_path]

Reads release_state.json (per the routine contract), writes a workbook whose
rows mirror the running list. New-since-last-run rows are highlighted; the
oldest job is flagged; ambiguous / missing-job-number rows are called out.
recalc.py performs a follow-up sweep to confirm zero formula errors.
"""
from __future__ import annotations

import json
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


COLUMNS = [
    ("jobNumber", "Job #", 14),
    ("name", "Job Name", 34),
    ("contractor", "Contractor", 24),
    ("pm", "PM", 18),
    ("contact", "Contact", 34),
    ("phone", "Phone", 24),
    ("storm", "Storm", 18),
    ("san", "San", 22),
    ("received", "Received", 12),
    ("notes", "Notes", 60),
]

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
NEW_FILL = PatternFill("solid", fgColor="FFF2CC")
FLAG_FILL = PatternFill("solid", fgColor="FCE4D6")
OLDEST_FILL = PatternFill("solid", fgColor="F8CBAD")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)


def load_state(path: Path) -> dict:
    return json.loads(path.read_text())


def build(state: dict, out: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Pending Releases"

    for col, (_, header, width) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"

    releases = state.get("releases", [])
    oldest = min(releases, key=lambda r: r.get("received", "9999-12-31"), default=None)
    oldest_id = oldest.get("messageId") if oldest else None

    for row_idx, entry in enumerate(releases, start=2):
        for col, (key, _, _) in enumerate(COLUMNS, start=1):
            cell = ws.cell(row=row_idx, column=col, value=entry.get(key, ""))
            cell.alignment = WRAP
            cell.border = BORDER
        if entry.get("flag"):
            for col in range(1, len(COLUMNS) + 1):
                ws.cell(row=row_idx, column=col).fill = FLAG_FILL
        elif entry.get("isNew"):
            for col in range(1, len(COLUMNS) + 1):
                ws.cell(row=row_idx, column=col).fill = NEW_FILL
        if entry.get("messageId") == oldest_id:
            ws.cell(row=row_idx, column=1).fill = OLDEST_FILL

    summary_row = len(releases) + 3
    ws.cell(row=summary_row, column=1, value="TOTAL PENDING").font = Font(bold=True)
    ws.cell(row=summary_row, column=2, value=len(releases)).font = Font(bold=True)
    ws.cell(row=summary_row + 1, column=1, value="NEW THIS RUN").font = Font(bold=True)
    ws.cell(
        row=summary_row + 1,
        column=2,
        value=sum(1 for r in releases if r.get("isNew")),
    ).font = Font(bold=True)
    ws.cell(row=summary_row + 2, column=1, value="FLAGGED").font = Font(bold=True)
    ws.cell(
        row=summary_row + 2,
        column=2,
        value=sum(1 for r in releases if r.get("flag")),
    ).font = Font(bold=True)
    ws.cell(row=summary_row + 3, column=1, value="RUN DATE").font = Font(bold=True)
    ws.cell(row=summary_row + 3, column=2, value=state.get("meta", {}).get("last_run_iso", str(date.today())))

    wb.save(out)


def main() -> None:
    root = Path(__file__).resolve().parent
    state_path = Path(sys.argv[1]) if len(sys.argv) > 1 else root / "release_state.json"
    out_path = Path(sys.argv[2]) if len(sys.argv) > 2 else root / "Production_Release_Log.xlsx"
    build(load_state(state_path), out_path)
    print(f"wrote {out_path}")


if __name__ == "__main__":
    main()
