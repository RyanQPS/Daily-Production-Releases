"""Build Production_Release_Log.xlsx from release_state.json."""
import json
from datetime import datetime
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).parent
STATE = json.loads((ROOT / "release_state.json").read_text())
RELEASES = STATE["releases"]
LAST_RUN = STATE.get("lastRun", "")

HEADERS = ["#", "Received", "Job #", "Job Name", "Contractor", "PM",
           "Contact", "Phone", "Storm", "San", "Notes", "Age (days)"]

today = datetime.fromisoformat("2026-06-28").date()

def age_days(received_iso: str) -> int:
    try:
        d = datetime.fromisoformat(received_iso).date()
        return (today - d).days
    except Exception:
        return 0

# Sort: oldest first (highest age)
rows_sorted = sorted(RELEASES, key=lambda r: r.get("received", ""))

wb = Workbook()
ws = wb.active
ws.title = "Pending Releases"

# Title row
ws["A1"] = "Production Release Log — Pending RTPs"
ws["A1"].font = Font(bold=True, size=14)
ws.merge_cells("A1:L1")
ws["A2"] = f"Generated: 2026-06-28  |  Total pending: {len(rows_sorted)}  |  Last run: {LAST_RUN}"
ws["A2"].font = Font(italic=True, color="555555")
ws.merge_cells("A2:L2")

# Header row at row 4
header_fill = PatternFill("solid", fgColor="305496")
header_font = Font(bold=True, color="FFFFFF")
thin = Side(border_style="thin", color="999999")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

for col_idx, h in enumerate(HEADERS, 1):
    c = ws.cell(row=4, column=col_idx, value=h)
    c.fill = header_fill
    c.font = header_font
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = border

# Flag fill (ambiguous / no job#)
flag_fill = PatternFill("solid", fgColor="FFF2CC")
old_fill = PatternFill("solid", fgColor="FCE4D6")  # >30d
oldest_fill = PatternFill("solid", fgColor="F8CBAD")  # the single oldest

# Compute oldest received
oldest_received = min((r.get("received", "") for r in rows_sorted if r.get("received")), default="")

for idx, r in enumerate(rows_sorted, 1):
    row_num = 4 + idx
    received = r.get("received", "")
    age = age_days(received)
    is_ambiguous = "FLAG" in (r.get("notes", "") or "") or "NO-JOBNUM" in r.get("jobNumber", "")
    is_oldest = received == oldest_received and received != ""
    values = [
        idx,
        received,
        r.get("jobNumber", ""),
        r.get("name", ""),
        r.get("contractor", ""),
        r.get("pm", ""),
        r.get("contact", ""),
        r.get("phone", ""),
        "Y" if r.get("storm") else "",
        "Y" if r.get("san") else "",
        r.get("notes", ""),
        age,
    ]
    for col_idx, v in enumerate(values, 1):
        c = ws.cell(row=row_num, column=col_idx, value=v)
        c.border = border
        c.alignment = Alignment(vertical="top", wrap_text=True)
        if is_oldest:
            c.fill = oldest_fill
        elif is_ambiguous:
            c.fill = flag_fill
        elif age > 30:
            c.fill = old_fill

# Column widths
widths = [4, 12, 16, 36, 32, 18, 22, 18, 6, 6, 50, 8]
for col_idx, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(col_idx)].width = w

# Footer summary block (no formulas with errors — use plain counts)
footer_row = 4 + len(rows_sorted) + 2
ws.cell(row=footer_row, column=1, value="Summary").font = Font(bold=True)
ws.cell(row=footer_row + 1, column=1, value="Total pending:")
ws.cell(row=footer_row + 1, column=2, value=len(rows_sorted))
ws.cell(row=footer_row + 2, column=1, value="Storm jobs:")
ws.cell(row=footer_row + 2, column=2, value=sum(1 for r in rows_sorted if r.get("storm")))
ws.cell(row=footer_row + 3, column=1, value="Sanitary jobs:")
ws.cell(row=footer_row + 3, column=2, value=sum(1 for r in rows_sorted if r.get("san")))
ws.cell(row=footer_row + 4, column=1, value="Flagged / ambiguous:")
ws.cell(row=footer_row + 4, column=2, value=sum(1 for r in rows_sorted if "FLAG" in (r.get("notes") or "") or "NO-JOBNUM" in r.get("jobNumber", "")))
ws.cell(row=footer_row + 5, column=1, value="Oldest job #:")
oldest = next((r for r in rows_sorted if r.get("received") == oldest_received), None)
ws.cell(row=footer_row + 5, column=2, value=oldest.get("jobNumber", "") if oldest else "")
ws.cell(row=footer_row + 5, column=3, value=oldest.get("name", "") if oldest else "")
ws.cell(row=footer_row + 5, column=4, value=f"({oldest_received})")

# Freeze panes below header
ws.freeze_panes = "A5"

# PM breakdown
pm_counts = {}
for r in rows_sorted:
    pm_counts[r.get("pm", "?")] = pm_counts.get(r.get("pm", "?"), 0) + 1

pm_row = footer_row + 7
ws.cell(row=pm_row, column=1, value="By PM").font = Font(bold=True)
for i, (pm, n) in enumerate(sorted(pm_counts.items(), key=lambda x: -x[1])):
    ws.cell(row=pm_row + 1 + i, column=1, value=pm)
    ws.cell(row=pm_row + 1 + i, column=2, value=n)

out = ROOT / "Production_Release_Log.xlsx"
wb.save(out)
print(f"Wrote {out} with {len(rows_sorted)} rows.")
