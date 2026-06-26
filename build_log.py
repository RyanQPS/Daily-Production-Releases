#!/usr/bin/env python3
"""Build Production_Release_Log.xlsx from release_state.json."""
import json
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).parent
STATE_FILE = ROOT / "release_state.json"
OUTPUT_FILE = ROOT / "Production_Release_Log.xlsx"

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
FLAG_FILL = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
NEW_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
OLDEST_FILL = PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid")
THIN = Side(border_style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

COLS = [
    ("Job #", "jobNumber", 14),
    ("Project / Site", "name", 36),
    ("Contractor", "contractor", 24),
    ("PM", "pm", 18),
    ("Contact", "contact", 22),
    ("Phone", "phone", 16),
    ("Storm", "storm", 8),
    ("San", "san", 8),
    ("Received", "received", 12),
    ("Age (days)", None, 11),
    ("Notes", "notes", 60),
]


def parse_date(s):
    if not s:
        return None
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except ValueError:
        return None


def build():
    state = json.loads(STATE_FILE.read_text())
    releases = state.get("releases", [])
    today = date.fromisoformat(state.get("lastRun")) if state.get("lastRun") else date.today()

    received_dates = [parse_date(r.get("received")) for r in releases]
    valid_dates = [d for d in received_dates if d]
    oldest = min(valid_dates) if valid_dates else None

    prev_message_ids = set(state.get("previousMessageIds", []))

    wb = Workbook()
    ws = wb.active
    ws.title = "Pending Releases"

    for col_idx, (header, _key, _width) in enumerate(COLS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER

    ws.row_dimensions[1].height = 28
    for col_idx, (_, _, width) in enumerate(COLS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    sorted_releases = sorted(
        releases,
        key=lambda r: (parse_date(r.get("received")) or date.max),
    )

    for row_idx, r in enumerate(sorted_releases, start=2):
        received = parse_date(r.get("received"))
        age = (today - received).days if received else None
        is_flag = str(r.get("jobNumber", "")).startswith("TBD") or "FLAG" in (r.get("notes") or "")
        is_new = r.get("messageId") not in prev_message_ids
        is_oldest = received is not None and oldest is not None and received == oldest

        values = [
            r.get("jobNumber", ""),
            r.get("name", ""),
            r.get("contractor", ""),
            r.get("pm", ""),
            r.get("contact", ""),
            r.get("phone", ""),
            r.get("storm", 0) or "",
            r.get("san", 0) or "",
            received.isoformat() if received else "",
            age if age is not None else "",
            r.get("notes", ""),
        ]

        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = BORDER
            if is_flag:
                cell.fill = FLAG_FILL
            elif is_oldest:
                cell.fill = OLDEST_FILL
            elif is_new:
                cell.fill = NEW_FILL

    last_row = len(sorted_releases) + 1
    summary_row = last_row + 2
    ws.cell(row=summary_row, column=1, value="Total Pending:").font = Font(bold=True)
    ws.cell(row=summary_row, column=2, value=f"=COUNTA(A2:A{last_row})")
    ws.cell(row=summary_row + 1, column=1, value="Storm Strs:").font = Font(bold=True)
    ws.cell(row=summary_row + 1, column=2, value=f"=SUM(G2:G{last_row})")
    ws.cell(row=summary_row + 2, column=1, value="San Strs:").font = Font(bold=True)
    ws.cell(row=summary_row + 2, column=2, value=f"=SUM(H2:H{last_row})")
    ws.cell(row=summary_row + 3, column=1, value="Oldest (days):").font = Font(bold=True)
    ws.cell(row=summary_row + 3, column=2, value=f"=MAX(J2:J{last_row})")
    ws.cell(row=summary_row + 4, column=1, value="Run Date:").font = Font(bold=True)
    ws.cell(row=summary_row + 4, column=2, value=today.isoformat())

    ws.freeze_panes = "A2"

    legend_row = summary_row + 6
    ws.cell(row=legend_row, column=1, value="Legend").font = Font(bold=True, underline="single")
    legend_items = [
        ("New since last run", NEW_FILL),
        ("Oldest pending", OLDEST_FILL),
        ("Flagged (ambiguous/missing job#)", FLAG_FILL),
    ]
    for i, (label, fill) in enumerate(legend_items, start=1):
        c = ws.cell(row=legend_row + i, column=1, value=label)
        c.fill = fill
        c.border = BORDER

    wb.save(OUTPUT_FILE)
    print(f"Wrote {OUTPUT_FILE} ({len(releases)} rows)")


if __name__ == "__main__":
    build()
