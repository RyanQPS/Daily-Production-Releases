#!/usr/bin/env python3
"""Validate Production_Release_Log.xlsx has zero formula errors.

build_log.py writes only static values (no formulas), so this script is a
belt-and-suspenders check that the workbook opens cleanly and every cell is a
value, not #REF!/#N/A/#DIV/0/etc.
"""
from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
XLSX = ROOT / "Production_Release_Log.xlsx"

ERR_MARKERS = {"#REF!", "#N/A", "#DIV/0!", "#VALUE!", "#NAME?", "#NULL!", "#NUM!"}


def main() -> int:
    wb = load_workbook(XLSX)
    errors = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v in ERR_MARKERS:
                    errors += 1
                    print(f"{ws.title}!{cell.coordinate}: {v}")
    wb.save(XLSX)
    print(f"Formula-error count: {errors}")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
