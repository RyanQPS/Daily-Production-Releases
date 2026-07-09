#!/usr/bin/env python3
"""Open Production_Release_Log.xlsx and verify it loads with zero formula errors.

build_log.py writes plain values (no formulas), so there is nothing to recalc.
This script simply loads the workbook, scans every cell for #REF!/#NAME?/#VALUE!
error strings, and exits non-zero if any are found.
"""
import sys
from pathlib import Path
from openpyxl import load_workbook

ROOT = Path(__file__).parent
XLSX = ROOT / "Production_Release_Log.xlsx"

ERROR_MARKERS = {"#REF!", "#NAME?", "#VALUE!", "#DIV/0!", "#N/A", "#NULL!", "#NUM!"}


def main():
    if not XLSX.exists():
        print(f"ERROR: {XLSX} does not exist. Run build_log.py first.")
        sys.exit(1)

    wb = load_workbook(XLSX, data_only=False)
    bad = []
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=False):
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v.strip() in ERROR_MARKERS:
                    bad.append(f"{ws.title}!{cell.coordinate} = {v!r}")

    if bad:
        print("Formula errors found:")
        for b in bad:
            print(" ", b)
        sys.exit(2)

    print(f"OK: {XLSX} loaded, zero formula errors.")


if __name__ == "__main__":
    main()
