#!/usr/bin/env python3
"""Build Production_Release_Log.xlsx from release_state.json.

Aggregates per-message entries into one row per jobNumber so the log
matches the rule: "the same job released by two PMs is ONE row."
All totals are computed in Python and written as static values to
guarantee zero formula errors when the file is opened in Excel.
"""

import json
from pathlib import Path
from datetime import datetime, date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
STATE_FILE = ROOT / "release_state.json"
OUT_FILE = ROOT / "Production_Release_Log.xlsx"

HEADER = [
    "Job #",
    "Project Name",
    "Contractor",
    "PM",
    "Contact",
    "Phone",
    "Storm",
    "San",
    "Received",
    "Days Pending",
    "Notes",
]


def aggregate(releases):
    """Group by jobNumber. Same job from multiple messages -> one row.

    storm/san fields are summed across messages (Rodrigo's 26-552 sends
    one storm message and one san message). Other text fields take the
    first non-empty value.
    """
    by_job = {}
    for r in releases:
        key = r["jobNumber"]
        if key not in by_job:
            by_job[key] = {
                "jobNumber": r["jobNumber"],
                "name": r.get("name", ""),
                "contractor": r.get("contractor", ""),
                "pm": r.get("pm", ""),
                "contact": r.get("contact", ""),
                "phone": r.get("phone", ""),
                "storm": 0,
                "san": 0,
                "received": r.get("received", ""),
                "notes": r.get("notes", ""),
                "messageIds": [r["messageId"]],
            }
            for extra in r.get("additionalMessageIds", []):
                by_job[key]["messageIds"].append(extra)
        else:
            entry = by_job[key]
            entry["messageIds"].append(r["messageId"])
            for extra in r.get("additionalMessageIds", []):
                entry["messageIds"].append(extra)
            for field in ("name", "contractor", "pm", "contact", "phone"):
                if not entry[field] and r.get(field):
                    entry[field] = r[field]
            if r.get("notes"):
                entry["notes"] = (entry["notes"] + " | " + r["notes"]).strip(" |")
            received = r.get("received", "")
            if received and (not entry["received"] or received < entry["received"]):
                entry["received"] = received
        by_job[key]["storm"] += int(r.get("storm") or 0)
        by_job[key]["san"] += int(r.get("san") or 0)
    return list(by_job.values())


def days_pending(received_iso, today):
    try:
        rcvd = datetime.fromisoformat(received_iso).date()
    except (ValueError, TypeError):
        return ""
    return (today - rcvd).days


def main():
    with STATE_FILE.open() as fh:
        state = json.load(fh)
    rows = aggregate(state["releases"])
    rows.sort(key=lambda r: r["received"])

    today = date.today()

    wb = Workbook()
    ws = wb.active
    ws.title = "Pending Releases"

    title_font = Font(bold=True, size=14)
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    border = Border(
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF"),
    )
    flag_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    stale_fill = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")

    ws["A1"] = "Production Release Log"
    ws["A1"].font = title_font
    ws.merge_cells("A1:K1")
    ws["A2"] = f"Generated {today.isoformat()} | Total pending: {len(rows)}"
    ws.merge_cells("A2:K2")

    header_row = 4
    for col, name in enumerate(HEADER, start=1):
        cell = ws.cell(row=header_row, column=col, value=name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    for i, r in enumerate(rows, start=header_row + 1):
        dp = days_pending(r["received"], today)
        values = [
            r["jobNumber"],
            r["name"],
            r["contractor"],
            r["pm"],
            r["contact"],
            r["phone"],
            r["storm"] if r["storm"] else "",
            r["san"] if r["san"] else "",
            r["received"],
            dp,
            r["notes"],
        ]
        for col, val in enumerate(values, start=1):
            cell = ws.cell(row=i, column=col, value=val)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = border
        notes_lower = (r["notes"] or "").lower()
        if "flag" in notes_lower:
            for col in range(1, len(HEADER) + 1):
                ws.cell(row=i, column=col).fill = flag_fill
        elif isinstance(dp, int) and dp >= 30:
            for col in range(1, len(HEADER) + 1):
                ws.cell(row=i, column=col).fill = stale_fill

    widths = [12, 32, 22, 18, 18, 15, 7, 7, 12, 8, 60]
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width

    summary_row = header_row + 1 + len(rows) + 1
    total_storm = sum(int(r["storm"] or 0) for r in rows)
    total_san = sum(int(r["san"] or 0) for r in rows)
    ws.cell(row=summary_row, column=1, value="TOTAL").font = Font(bold=True)
    ws.cell(row=summary_row, column=7, value=total_storm).font = Font(bold=True)
    ws.cell(row=summary_row, column=8, value=total_san).font = Font(bold=True)

    ws.freeze_panes = "A5"

    wb.save(OUT_FILE)
    print(f"Wrote {OUT_FILE} with {len(rows)} job rows (storm total {total_storm}, san total {total_san}).")


if __name__ == "__main__":
    main()
