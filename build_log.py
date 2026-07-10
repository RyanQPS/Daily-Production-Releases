#!/usr/bin/env python3
"""Build Production_Release_Log.xlsx from release_state.json."""
import json
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

REPO = Path(__file__).resolve().parent
STATE_PATH = REPO / "release_state.json"
OUT_PATH = REPO / "Production_Release_Log.xlsx"

COLUMNS = [
    ("jobNumber", "Job #", 12),
    ("name", "Project Name", 40),
    ("contractor", "Contractor", 28),
    ("pm", "PM", 20),
    ("contact", "Contact", 28),
    ("phone", "Phone", 22),
    ("storm", "Storm", 12),
    ("san", "Sanitary", 14),
    ("received", "Received", 14),
    ("age_days", "Age (days)", 12),
    ("notes", "Notes", 60),
    ("messageId", "Message-ID", 60),
]

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
NEW_FILL = PatternFill("solid", fgColor="FFF2CC")
OLDEST_FILL = PatternFill("solid", fgColor="F8CBAD")
FLAG_FILL = PatternFill("solid", fgColor="E2EFDA")
BORDER = Border(*(Side(style="thin", color="BFBFBF") for _ in range(4)))


def load_state():
    with STATE_PATH.open() as fh:
        return json.load(fh)


def parse_date(text):
    if not text:
        return None
    try:
        return datetime.strptime(text[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def build_workbook(state):
    today = datetime.strptime(state["generatedAt"][:10], "%Y-%m-%d").date()
    releases = state.get("releases", [])
    flagged = state.get("flagged", [])
    new_ids = {r["messageId"] for r in releases}  # first run: everything is new

    # Determine oldest by received date
    with_dates = [(parse_date(r.get("received")), r) for r in releases]
    with_dates = [(d, r) for d, r in with_dates if d]
    oldest_id = None
    if with_dates:
        with_dates.sort(key=lambda x: x[0])
        oldest_id = with_dates[0][1]["messageId"]

    wb = Workbook()
    ws = wb.active
    ws.title = "Pending Releases"

    # Title row
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLUMNS))
    title_cell = ws.cell(row=1, column=1, value=f"Production Release Log — as of {state['generatedAt'][:10]}")
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal="left")

    # Header
    for idx, (_, header, width) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=3, column=idx, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.row_dimensions[3].height = 22

    # Rows
    row_num = 4
    for release in releases:
        received = parse_date(release.get("received"))
        age = (today - received).days if received else ""
        row_fill = None
        if release["messageId"] == oldest_id:
            row_fill = OLDEST_FILL
        elif release["messageId"] in new_ids:
            row_fill = NEW_FILL

        for idx, (key, _, _) in enumerate(COLUMNS, start=1):
            if key == "age_days":
                value = age
            else:
                value = release.get(key, "")
            cell = ws.cell(row=row_num, column=idx, value=value)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = BORDER
            if row_fill:
                cell.fill = row_fill
        row_num += 1

    # Freeze header
    ws.freeze_panes = "A4"

    # Auto row height guess
    for r in range(4, row_num):
        ws.row_dimensions[r].height = 42

    # Flagged sheet
    if flagged:
        fw = wb.create_sheet("Flagged")
        headers = ["Received", "PM", "Subject", "Reason", "Message-ID"]
        widths = [14, 22, 60, 60, 60]
        for idx, (header, width) in enumerate(zip(headers, widths), start=1):
            cell = fw.cell(row=1, column=idx, value=header)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = BORDER
            fw.column_dimensions[get_column_letter(idx)].width = width
        for i, item in enumerate(flagged, start=2):
            row = [item.get("received", ""), item.get("pm", ""), item.get("subject", ""),
                   item.get("reason", ""), item.get("messageId", "")]
            for idx, value in enumerate(row, start=1):
                cell = fw.cell(row=i, column=idx, value=value)
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.border = BORDER
                cell.fill = FLAG_FILL
            fw.row_dimensions[i].height = 48

    # Summary sheet
    sw = wb.create_sheet("Summary", 0)
    sw.column_dimensions["A"].width = 32
    sw.column_dimensions["B"].width = 20
    sw["A1"] = "Metric"
    sw["B1"] = "Value"
    for c in (sw["A1"], sw["B1"]):
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center")
        c.border = BORDER

    pm_counts = {}
    for r in releases:
        pm_counts[r.get("pm", "Unknown")] = pm_counts.get(r.get("pm", "Unknown"), 0) + 1

    rows = [
        ("Generated", state["generatedAt"]),
        ("Total pending releases", len(releases)),
        ("New this run", len(new_ids)),
        ("Flagged / ambiguous", len(flagged)),
        ("Oldest job (by received)", f"{with_dates[0][1].get('jobNumber','')} {with_dates[0][1].get('name','')} — {with_dates[0][0].isoformat()}" if with_dates else "n/a"),
    ]
    for pm, count in sorted(pm_counts.items()):
        rows.append((f"PM: {pm}", count))
    for i, (label, value) in enumerate(rows, start=2):
        sw.cell(row=i, column=1, value=label).border = BORDER
        cell = sw.cell(row=i, column=2, value=value)
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="left")

    wb.save(OUT_PATH)
    return OUT_PATH


def main():
    state = load_state()
    path = build_workbook(state)
    print(f"Wrote {path}")


if __name__ == "__main__":
    main()
