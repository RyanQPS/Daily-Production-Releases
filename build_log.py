#!/usr/bin/env python3
"""Build Production_Release_Log.xlsx from release_state.json.

Deliberately writes only literal values (no formulas) so the workbook opens with
zero formula errors. recalc.py is a headless LibreOffice pass that resaves the
file to normalize any dependent metadata Excel would otherwise regenerate on
first open.
"""
from __future__ import annotations

import json
import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
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
    "Phone",
    "Storm",
    "Sanitary",
    "Received",
    "Age (days)",
    "New?",
    "Notes",
]

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
NEW_FILL = PatternFill("solid", fgColor="FFF2CC")
OLDEST_FILL = PatternFill("solid", fgColor="F8CBAD")
FLAG_FILL = PatternFill("solid", fgColor="FFC7CE")
HEADER_FONT = Font(bold=True, color="FFFFFF")
THIN = Side(border_style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _parse_date(value: str) -> date | None:
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError:
        return None


def build() -> None:
    state = json.loads(STATE_PATH.read_text())
    releases = state["releases"]

    today = date.today()
    for r in releases:
        d = _parse_date(r.get("received", ""))
        r["_age"] = (today - d).days if d else None

    ages = [r["_age"] for r in releases if r["_age"] is not None]
    oldest_age = max(ages) if ages else None

    wb = Workbook()
    ws = wb.active
    ws.title = "Pending Releases"

    ws.append(HEADERS)
    for col_idx, _ in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER

    for r in releases:
        row = [
            r.get("jobNumber", ""),
            r.get("name", ""),
            r.get("contractor", ""),
            r.get("pm", ""),
            r.get("contact", ""),
            r.get("phone", ""),
            "Y" if r.get("storm") else "",
            "Y" if r.get("san") else "",
            r.get("received", ""),
            r["_age"] if r["_age"] is not None else "",
            "NEW" if r.get("new") else "",
            r.get("notes", ""),
        ]
        ws.append(row)

        last = ws.max_row
        is_flag = str(r.get("jobNumber", "")).startswith("FLAG")
        is_oldest = oldest_age is not None and r["_age"] == oldest_age
        is_new = bool(r.get("new"))
        fill = None
        if is_flag:
            fill = FLAG_FILL
        elif is_oldest:
            fill = OLDEST_FILL
        elif is_new:
            fill = NEW_FILL
        for col_idx in range(1, len(HEADERS) + 1):
            cell = ws.cell(row=last, column=col_idx)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = BORDER
            if fill is not None:
                cell.fill = fill

    widths = [12, 34, 22, 18, 22, 26, 8, 10, 12, 10, 8, 60]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    summary = wb.create_sheet("Summary")
    total = len(releases)
    new_count = sum(1 for r in releases if r.get("new"))
    flagged = [r for r in releases if str(r.get("jobNumber", "")).startswith("FLAG")]
    oldest_row = None
    if oldest_age is not None:
        for r in releases:
            if r["_age"] == oldest_age:
                oldest_row = r
                break

    lines: list[list[object]] = [
        ["Production Release Log — Summary"],
        ["Generated (state)", state.get("generatedAt", "")],
        ["Total pending", total],
        ["New this run", new_count],
        ["Flagged (ambiguous)", len(flagged)],
        ["Oldest job",
         f"{oldest_row['jobNumber']} — {oldest_row['name']}" if oldest_row else "",
         f"{oldest_age} days" if oldest_age is not None else ""],
        [],
        ["Notes from this run"],
    ]
    for n in state.get("runNotes", []):
        lines.append([n])

    for line in lines:
        summary.append(line)

    summary.column_dimensions["A"].width = 28
    summary.column_dimensions["B"].width = 60
    summary.column_dimensions["C"].width = 16
    summary["A1"].font = Font(bold=True, size=14)

    wb.save(OUT_PATH)
    print(f"Wrote {OUT_PATH}")


if __name__ == "__main__":
    try:
        build()
    except Exception as exc:
        print(f"build_log.py failed: {exc}", file=sys.stderr)
        raise
