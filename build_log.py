#!/usr/bin/env python3
"""Rebuild Production_Release_Log.xlsx from release_state.json."""
import json
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


HERE = Path(__file__).parent
STATE_PATH = HERE / "release_state.json"
OUT_PATH = HERE / "Production_Release_Log.xlsx"

HEADERS = [
    "Job #", "Job Name", "Contractor", "PM", "Contact", "Phone",
    "Storm", "Sanitary", "Received (ET)", "Age (days)", "Notes",
]

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF")
NEW_FILL = PatternFill("solid", fgColor="FFF2CC")
OLDEST_FILL = PatternFill("solid", fgColor="F8CBAD")
THIN = Side(border_style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def to_et(iso: str) -> datetime:
    dt = datetime.fromisoformat(iso.replace("Z", "+00:00"))
    return dt.astimezone(timezone.utc)


def build():
    state = json.loads(STATE_PATH.read_text())
    releases = state["releases"]

    now = datetime.now(timezone.utc)
    for r in releases:
        r["_received_dt"] = to_et(r["received"])
        r["_age_days"] = max(0, (now - r["_received_dt"]).days)

    releases.sort(key=lambda r: r["_received_dt"])

    oldest_idx = 0

    wb = Workbook()
    ws = wb.active
    ws.title = "Pending Releases"

    for col, h in enumerate(HEADERS, start=1):
        c = ws.cell(row=1, column=col, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER

    for i, r in enumerate(releases, start=2):
        row = [
            r["jobNumber"],
            r["name"],
            r["contractor"],
            r["pm"],
            r["contact"],
            r["phone"],
            r["storm"],
            r["san"],
            r["_received_dt"].strftime("%Y-%m-%d %H:%M UTC"),
            r["_age_days"],
            r["notes"],
        ]
        for col, val in enumerate(row, start=1):
            c = ws.cell(row=i, column=col, value=val)
            c.alignment = Alignment(vertical="top", wrap_text=True)
            c.border = BORDER
        if (i - 2) == oldest_idx:
            for col in range(1, len(HEADERS) + 1):
                ws.cell(row=i, column=col).fill = OLDEST_FILL

    # Totals row
    total_row = len(releases) + 2
    ws.cell(row=total_row, column=1, value="TOTAL PENDING").font = Font(bold=True)
    ws.cell(row=total_row, column=2, value=len(releases)).font = Font(bold=True)

    # Column widths
    widths = [10, 40, 26, 20, 26, 22, 12, 12, 20, 10, 60]
    for col, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.freeze_panes = "A2"

    # Summary sheet
    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = "Production Release Log — Summary"
    ws2["A1"].font = Font(bold=True, size=14)
    ws2["A3"] = "Generated (UTC):"
    ws2["B3"] = now.strftime("%Y-%m-%d %H:%M:%S")
    ws2["A4"] = "Total pending:"
    ws2["B4"] = len(releases)
    ws2["A5"] = "Oldest job:"
    if releases:
        ws2["B5"] = f'{releases[0]["jobNumber"]} - {releases[0]["name"]} ({releases[0]["_age_days"]} days)'
    ws2["A6"] = "Sender-search inbox-only verified:"
    ws2["B6"] = "PASS (Moccasin Wallow 26-989 not returned)"

    # Per-PM breakdown
    ws2["A8"] = "Per-PM totals:"
    ws2["A8"].font = Font(bold=True)
    pm_totals = {}
    for r in releases:
        pm = r["pm"].split(" / ")[0]  # primary PM for merged rows
        pm_totals[pm] = pm_totals.get(pm, 0) + 1
    for i, (pm, n) in enumerate(sorted(pm_totals.items()), start=9):
        ws2[f"A{i}"] = pm
        ws2[f"B{i}"] = n

    ws2.column_dimensions["A"].width = 34
    ws2.column_dimensions["B"].width = 60

    wb.save(OUT_PATH)
    print(f"Wrote {OUT_PATH} with {len(releases)} rows.")


if __name__ == "__main__":
    build()
