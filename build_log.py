#!/usr/bin/env python3
"""Rebuild Production_Release_Log.xlsx from release_state.json.

Output columns: #, Job #, Job Name, Contractor, PM, Contact, Phone,
Storm, San, Total, Received, Days Pending, Notes, Message-Id.

Formulas:
  Total       = Storm + San (integers only; else literal from state)
  Days Pending = TODAY() - Received

Formatting: header row bold with fill, freeze top row, autosize columns,
new rows (firstSeen == today) shaded yellow.
"""
from __future__ import annotations

import json
import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
STATE_PATH = ROOT / "release_state.json"
XLSX_PATH = ROOT / "Production_Release_Log.xlsx"

HEADERS = [
    "#", "Job #", "Job Name", "Contractor", "PM", "Contact", "Phone",
    "Storm", "San", "Total", "Received", "Days Pending", "Notes", "Message-Id",
]

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
NEW_FILL = PatternFill("solid", fgColor="FFF2CC")
OLD_FILL = PatternFill("solid", fgColor="F8CBAD")


def _int_or_str(val):
    if isinstance(val, int):
        return val
    try:
        return int(str(val).strip())
    except (TypeError, ValueError):
        return val


def build(state_path: Path = STATE_PATH, out_path: Path = XLSX_PATH) -> Path:
    state = json.loads(state_path.read_text())
    pending = state.get("pending", [])
    today = date.today()

    wb = Workbook()
    ws = wb.active
    ws.title = "Pending"

    for col, header in enumerate(HEADERS, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.freeze_panes = "A2"

    for idx, row in enumerate(pending, 1):
        excel_row = idx + 1
        received = row.get("received", "")
        try:
            recv_date = datetime.strptime(received, "%Y-%m-%d").date()
        except ValueError:
            recv_date = None

        storm = _int_or_str(row.get("storm"))
        san = _int_or_str(row.get("san"))

        values = [
            idx,
            row.get("jobNumber", ""),
            row.get("name", ""),
            row.get("contractor", ""),
            row.get("pm", ""),
            row.get("contact", ""),
            row.get("phone", ""),
            storm,
            san,
            None,
            recv_date if recv_date else received,
            None,
            row.get("notes", ""),
            row.get("messageId", ""),
        ]

        for col, val in enumerate(values, 1):
            ws.cell(row=excel_row, column=col, value=val)

        # Total = Storm + San when both numeric; else literal
        if isinstance(storm, int) and isinstance(san, int):
            ws.cell(row=excel_row, column=10, value=f"=H{excel_row}+I{excel_row}")
        else:
            ws.cell(row=excel_row, column=10, value="See Notes")

        # Days Pending
        if recv_date:
            ws.cell(row=excel_row, column=11).number_format = "yyyy-mm-dd"
            ws.cell(row=excel_row, column=12, value=f"=TODAY()-K{excel_row}")
            ws.cell(row=excel_row, column=12).number_format = "0"
        else:
            ws.cell(row=excel_row, column=12, value="")

        # Highlight new (firstSeen == today) or stale (> 30 days)
        first_seen = row.get("firstSeen")
        if first_seen == today.isoformat():
            for col in range(1, len(HEADERS) + 1):
                ws.cell(row=excel_row, column=col).fill = NEW_FILL
        elif recv_date and (today - recv_date).days > 30:
            for col in range(1, len(HEADERS) + 1):
                ws.cell(row=excel_row, column=col).fill = OLD_FILL

    # Column widths
    widths = [4, 14, 34, 24, 18, 22, 26, 8, 8, 9, 12, 10, 60, 60]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    for r in range(2, len(pending) + 2):
        ws.row_dimensions[r].height = 30
        for c in range(1, len(HEADERS) + 1):
            ws.cell(row=r, column=c).alignment = Alignment(vertical="top", wrap_text=True)

    # Flagged sheet
    flagged = state.get("flagged", [])
    if flagged:
        ws2 = wb.create_sheet("Flagged")
        flag_headers = ["Reason", "Job #", "Job Name", "Contractor", "PM", "Detail", "Received (RE:)"]
        for col, header in enumerate(flag_headers, 1):
            cell = ws2.cell(row=1, column=col, value=header)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
        ws2.freeze_panes = "A2"
        for i, row in enumerate(flagged, 2):
            ws2.cell(row=i, column=1, value=row.get("reason", ""))
            ws2.cell(row=i, column=2, value=row.get("jobNumber", ""))
            ws2.cell(row=i, column=3, value=row.get("name", ""))
            ws2.cell(row=i, column=4, value=row.get("contractor", ""))
            ws2.cell(row=i, column=5, value=row.get("pm", ""))
            ws2.cell(row=i, column=6, value=row.get("detail", ""))
            ws2.cell(row=i, column=7, value=row.get("receivedRE", ""))
        for i, w in enumerate([40, 14, 24, 22, 18, 80, 14], 1):
            ws2.column_dimensions[get_column_letter(i)].width = w
        for r in range(2, len(flagged) + 2):
            ws2.row_dimensions[r].height = 45
            for c in range(1, len(flag_headers) + 1):
                ws2.cell(row=r, column=c).alignment = Alignment(vertical="top", wrap_text=True)

    wb.save(out_path)
    return out_path


if __name__ == "__main__":
    path = build()
    print(f"Wrote {path}")
