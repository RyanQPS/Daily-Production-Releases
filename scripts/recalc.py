#!/usr/bin/env python3
"""Re-evaluate formulas in Production_Release_Log.xlsx and write the results back
as values so Excel never opens a stale `#VALUE!` or `0` placeholder.

We only have one formula column (Days Pending = TODAY() - Received) so we
compute it directly in Python and overwrite the cell. The file then has zero
unevaluated formulas.
"""
from __future__ import annotations

import os
from datetime import date, datetime

from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_FILE = os.path.join(ROOT, "Production_Release_Log.xlsx")


def recalc() -> None:
    wb = load_workbook(OUT_FILE)
    ws = wb["Pending Releases"]
    today = date.today()
    fixed = 0
    for row in ws.iter_rows(min_row=2, values_only=False):
        received_cell = row[8]  # column I
        days_cell = row[9]      # column J
        received_value = received_cell.value
        if not received_value:
            days_cell.value = ""
            continue
        if isinstance(received_value, datetime):
            received_date = received_value.date()
        elif isinstance(received_value, date):
            received_date = received_value
        else:
            try:
                received_date = datetime.strptime(
                    str(received_value), "%Y-%m-%d"
                ).date()
            except ValueError:
                days_cell.value = ""
                continue
        days_cell.value = (today - received_date).days
        fixed += 1
    wb.save(OUT_FILE)
    print(f"Recalculated {fixed} Days Pending values")


if __name__ == "__main__":
    recalc()
