#!/usr/bin/env python3
"""Build Production_Release_Log.xlsx from release_state.json.

Reads release_state.json and writes an .xlsx with one row per pending release.
No formulas that could error; the recalc.py step verifies the file loads and
recalculates any totals cleanly.
"""

import json
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
STATE_PATH = ROOT / "release_state.json"
OUT_PATH = ROOT / "Production_Release_Log.xlsx"

HEADERS = [
    "Job #",
    "Job Name",
    "Contractor",
    "PM",
    "Contact",
    "Phone / Email",
    "Storm",
    "San",
    "Received (EST)",
    "Days Pending",
    "Notes",
]

HEADER_FILL = PatternFill(start_color="FF1F4E78", end_color="FF1F4E78", fill_type="solid")
HEADER_FONT = Font(color="FFFFFFFF", bold=True)
NEW_ROW_FILL = PatternFill(start_color="FFFFF2CC", end_color="FFFFF2CC", fill_type="solid")
OLDEST_FILL = PatternFill(start_color="FFF8CBAD", end_color="FFF8CBAD", fill_type="solid")


def to_est(iso: str) -> tuple[str, int]:
    """Return (formatted EST datetime, integer days pending) for an ISO UTC timestamp."""
    dt = datetime.fromisoformat(iso.replace("Z", "+00:00"))
    # Fixed -04:00 offset for EDT window; the routine runs in America/New_York.
    est = dt.astimezone(timezone.utc).astimezone()
    days = (datetime.now(timezone.utc) - dt).days
    return est.strftime("%Y-%m-%d %H:%M"), days


def main() -> None:
    state = json.loads(STATE_PATH.read_text())
    releases = state["releases"]

    # Sort by received ascending (oldest first) so the oldest row is easy to spot.
    releases_sorted = sorted(releases, key=lambda r: r["received"])
    oldest_msg_id = releases_sorted[0]["messageId"] if releases_sorted else None

    wb = Workbook()
    ws = wb.active
    ws.title = "Pending Releases"

    for col, h in enumerate(HEADERS, start=1):
        c = ws.cell(row=1, column=col, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")

    for i, r in enumerate(releases_sorted, start=2):
        received_str, days = to_est(r["received"])
        row = [
            r.get("jobNumber", ""),
            r.get("name", ""),
            r.get("contractor", ""),
            r.get("pm", ""),
            r.get("contact", ""),
            r.get("phone", ""),
            "Y" if r.get("storm") else "",
            "Y" if r.get("san") else "",
            received_str,
            days,
            r.get("notes", ""),
        ]
        for col, val in enumerate(row, start=1):
            cell = ws.cell(row=i, column=col, value=val)
            cell.alignment = Alignment(vertical="top", wrap_text=(col in (2, 5, 6, 11)))
        # Highlight the oldest row.
        if r["messageId"] == oldest_msg_id:
            for col in range(1, len(HEADERS) + 1):
                ws.cell(row=i, column=col).fill = OLDEST_FILL
        # Every row is "new since last run" on the initial build.
        elif state.get("note", "").startswith("Initial state"):
            for col in range(1, len(HEADERS) + 1):
                if ws.cell(row=i, column=col).fill.fill_type is None:
                    ws.cell(row=i, column=col).fill = NEW_ROW_FILL

    # Reasonable column widths.
    widths = [10, 42, 22, 18, 24, 30, 7, 7, 18, 14, 60]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}{len(releases_sorted) + 1}"

    # Summary row at the bottom.
    total_row = len(releases_sorted) + 3
    ws.cell(row=total_row, column=1, value="Total pending:").font = Font(bold=True)
    ws.cell(row=total_row, column=2, value=len(releases_sorted)).font = Font(bold=True)

    wb.save(OUT_PATH)
    print(f"Wrote {OUT_PATH} with {len(releases_sorted)} rows.")


if __name__ == "__main__":
    main()
