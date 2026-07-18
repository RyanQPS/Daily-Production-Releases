#!/usr/bin/env python3
"""Rebuild Production_Release_Log.xlsx from release_state.json."""
import json
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
STATE_PATH = ROOT / "release_state.json"
OUT_PATH = ROOT / "Production_Release_Log.xlsx"

HEADERS = [
    "#", "Job Number", "Job Name", "Contractor", "PM(s)", "Contact",
    "Phone", "Storm", "San", "Structures",
    "Received (UTC)", "Days Pending", "Notes", "Flagged",
]

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
FLAG_FILL = PatternFill("solid", fgColor="FFF2CC")
OLD_FILL = PatternFill("solid", fgColor="F8CBAD")
THIN = Side(border_style="thin", color="B4B4B4")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def parse_iso(value: str) -> datetime:
    return datetime.fromisoformat(value.replace("Z", "+00:00"))


def build() -> None:
    with STATE_PATH.open() as fh:
        state = json.load(fh)

    releases = sorted(
        state["releases"],
        key=lambda r: parse_iso(r["received"]),
    )

    wb = Workbook()
    ws = wb.active
    ws.title = "Pending Releases"
    ws.freeze_panes = "A2"

    for col, name in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col, value=name)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER

    now = datetime.now(timezone.utc)
    oldest = min(parse_iso(r["received"]) for r in releases) if releases else now

    for idx, r in enumerate(releases, start=1):
        received = parse_iso(r["received"])
        days = (now - received).days
        row_values = [
            idx,
            r.get("jobNumber", ""),
            r.get("name", ""),
            r.get("contractor", ""),
            " / ".join(r.get("pm", [])) if isinstance(r.get("pm"), list) else r.get("pm", ""),
            r.get("contact", ""),
            r.get("phone", ""),
            r.get("storm", ""),
            r.get("san", ""),
            r.get("structures", ""),
            received.strftime("%Y-%m-%d %H:%M"),
            days,
            r.get("notes", ""),
            "YES" if r.get("flagged") else "",
        ]
        row_no = idx + 1
        for col, val in enumerate(row_values, start=1):
            cell = ws.cell(row=row_no, column=col, value=val)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = BORDER
        if received == oldest:
            for col in range(1, len(HEADERS) + 1):
                ws.cell(row=row_no, column=col).fill = OLD_FILL
        elif r.get("flagged"):
            for col in range(1, len(HEADERS) + 1):
                ws.cell(row=row_no, column=col).fill = FLAG_FILL

    widths = [4, 22, 40, 22, 22, 22, 18, 12, 12, 14, 18, 8, 50, 8]
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width

    summary = wb.create_sheet("Summary")
    summary["A1"] = "Field"
    summary["B1"] = "Value"
    summary["A1"].font = HEADER_FONT
    summary["B1"].font = HEADER_FONT
    summary["A1"].fill = HEADER_FILL
    summary["B1"].fill = HEADER_FILL

    generated = state.get("generatedAt", now.strftime("%Y-%m-%dT%H:%M:%SZ"))
    rows = [
        ("Generated (UTC)", generated),
        ("Total pending", len(releases)),
        ("Oldest release (UTC)", oldest.strftime("%Y-%m-%d %H:%M")),
        ("Flagged rows", sum(1 for r in releases if r.get("flagged"))),
    ]
    for i, (k, v) in enumerate(rows, start=2):
        summary.cell(row=i, column=1, value=k)
        summary.cell(row=i, column=2, value=v)
    summary.column_dimensions["A"].width = 22
    summary.column_dimensions["B"].width = 40

    wb.save(OUT_PATH)
    print(f"Wrote {OUT_PATH} with {len(releases)} pending releases")


if __name__ == "__main__":
    build()
