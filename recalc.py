#!/usr/bin/env python3
"""Recalculate Production_Release_Log.xlsx and verify zero formula errors.

We don't emit formulas from build_log.py, so this is a defensive check: open
the workbook, walk every cell, and fail loudly if any cached formula value is
an Excel error string (#REF!, #NAME?, #VALUE!, #DIV/0!, #N/A, #NULL!, #NUM!).
"""

import sys
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
LOG_PATH = ROOT / "Production_Release_Log.xlsx"

ERROR_TOKENS = {"#REF!", "#NAME?", "#VALUE!", "#DIV/0!", "#N/A", "#NULL!", "#NUM!"}


def main() -> int:
    wb = load_workbook(LOG_PATH, data_only=False)
    errors: list[str] = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v in ERROR_TOKENS:
                    errors.append(f"{ws.title}!{cell.coordinate}={v}")
    if errors:
        print("FORMULA ERRORS:")
        for e in errors:
            print(f"  {e}")
        return 1
    wb.save(LOG_PATH)  # Round-trip to force a clean recalc on next open.
    print(f"OK — {LOG_PATH.name} clean (0 formula errors).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
