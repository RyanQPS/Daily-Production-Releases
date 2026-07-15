#!/usr/bin/env python3
"""Evaluate every formula in Production_Release_Log.xlsx and write the
result back into the cell so consumers see values instead of #NAME?/#REF!
in viewers without a recalculation engine.

Only handles the two formulas produced by build_log.py:
    =Hn+In        -> Storm + San
    =TODAY()-Kn   -> days since received (Kn is a date cell)
"""
from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
XLSX = ROOT / "Production_Release_Log.xlsx"


def _to_date(val):
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    if isinstance(val, str):
        try:
            return datetime.strptime(val, "%Y-%m-%d").date()
        except ValueError:
            return None
    return None


def recalc(path: Path = XLSX) -> int:
    wb = load_workbook(path)
    ws = wb["Pending"]
    errors = 0
    today = date.today()
    sum_re = re.compile(r"^=([A-Z]+)(\d+)\+([A-Z]+)(\d+)$")
    day_re = re.compile(r"^=TODAY\(\)-([A-Z]+)(\d+)$", re.IGNORECASE)

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            v = cell.value
            if not isinstance(v, str) or not v.startswith("="):
                continue
            m = sum_re.match(v)
            if m:
                a = ws[f"{m.group(1)}{m.group(2)}"].value
                b = ws[f"{m.group(3)}{m.group(4)}"].value
                if isinstance(a, int) and isinstance(b, int):
                    cell.value = a + b
                else:
                    cell.value = "See Notes"
                continue
            m = day_re.match(v)
            if m:
                ref = ws[f"{m.group(1)}{m.group(2)}"].value
                d = _to_date(ref)
                cell.value = (today - d).days if d else ""
                continue
            errors += 1
            print(f"Unrecognised formula at {cell.coordinate}: {v}")

    wb.save(path)
    return errors


if __name__ == "__main__":
    n = recalc()
    if n:
        raise SystemExit(f"{n} unrecognised formulas — resolve before shipping.")
    print("recalc: 0 formula errors")
