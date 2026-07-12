#!/usr/bin/env python3
"""Rebuild Production_Release_Log.xlsx from release_state.json.

Values-only workbook — no formulas, so no formula errors.
"""
import json
import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
STATE = ROOT / "release_state.json"
XLSX = ROOT / "Production_Release_Log.xlsx"

COLUMNS = [
    ("Job #", "jobNumber", 14),
    ("Job Name", "name", 34),
    ("Contractor", "contractor", 22),
    ("PM", "pm", 16),
    ("Contact", "contact", 22),
    ("Phone", "phone", 20),
    ("Storm", "storm", 8),
    ("San", "san", 8),
    ("Received", "received", 12),
    ("Age (d)", None, 8),
    ("Notes", "notes", 44),
]


def as_of() -> date:
    with STATE.open() as fh:
        blob = json.load(fh)
    ts = blob.get("as_of", "")
    if ts:
        return datetime.fromisoformat(ts.replace("Z", "+00:00")).date()
    return date(2026, 7, 12)


def age_days(received: str, today: date) -> int:
    try:
        return (today - date.fromisoformat(received)).days
    except (TypeError, ValueError):
        return 0


def build() -> Path:
    with STATE.open() as fh:
        blob = json.load(fh)
    releases = blob["releases"]
    today = as_of()

    wb = Workbook()
    ws = wb.active
    ws.title = "Pending RTPs"

    header_fill = PatternFill("solid", fgColor="1F3864")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    body_align = Alignment(vertical="top", wrap_text=True)
    oldest_fill = PatternFill("solid", fgColor="F8CBAD")
    stale_fill = PatternFill("solid", fgColor="FFE699")
    border = Border(
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF"),
    )

    for idx, (title, _, width) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=idx, value=title)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_align
        cell.border = border
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.freeze_panes = "A2"

    oldest_age = max(age_days(r.get("received", ""), today) for r in releases)

    for row_idx, rel in enumerate(sorted(releases, key=lambda r: r.get("received", "")), start=2):
        age = age_days(rel.get("received", ""), today)
        for col_idx, (_, key, _) in enumerate(COLUMNS, start=1):
            if key is None:
                value = age
            else:
                value = rel.get(key, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.alignment = body_align
            cell.border = border
            if age == oldest_age:
                cell.fill = oldest_fill
            elif age >= 30:
                cell.fill = stale_fill

    # Summary sheet
    summary = wb.create_sheet("Summary")
    per_pm = {}
    for rel in releases:
        pm = rel.get("pm", "?")
        per_pm[pm] = per_pm.get(pm, 0) + 1
    summary.append(["Metric", "Value"])
    summary.append(["As-of date", today.isoformat()])
    summary.append(["Total log rows", len(releases)])
    unique_jobs = len({r.get("jobNumber", "") for r in releases if r.get("jobNumber")})
    summary.append(["Unique job numbers", unique_jobs])
    summary.append(["Oldest age (d)", oldest_age])
    summary.append(["", ""])
    summary.append(["Per-PM rows", ""])
    for pm, count in sorted(per_pm.items(), key=lambda kv: (-kv[1], kv[0])):
        summary.append([pm, count])

    summary.column_dimensions["A"].width = 22
    summary.column_dimensions["B"].width = 14
    for row in summary.iter_rows(min_row=1, max_row=1):
        for cell in row:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_align

    wb.save(XLSX)
    return XLSX


if __name__ == "__main__":
    path = build()
    print(f"Wrote {path}")
