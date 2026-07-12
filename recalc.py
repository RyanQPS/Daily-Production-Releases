#!/usr/bin/env python3
"""Verify Production_Release_Log.xlsx contains no formula-error cells.

Values-only workbook (no formulas), but this pass re-opens and re-saves it
so the routine's spec — build_log.py then recalc.py — is honored, and it
raises loudly if any cell contains a #REF!/#VALUE!/#NAME?/#N/A/#DIV/0!/#NULL!
error string.
"""
import sys
from pathlib import Path

from openpyxl import load_workbook

XLSX = Path(__file__).resolve().parent / "Production_Release_Log.xlsx"
ERR_TOKENS = ("#REF!", "#VALUE!", "#NAME?", "#N/A", "#DIV/0!", "#NULL!", "#NUM!")


def main() -> int:
    if not XLSX.exists():
        print(f"missing: {XLSX}", file=sys.stderr)
        return 1
    wb = load_workbook(XLSX)
    errors = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v in ERR_TOKENS:
                    errors.append(f"{ws.title}!{cell.coordinate}={v}")
    if errors:
        print("formula errors:", *errors, sep="\n  ", file=sys.stderr)
        return 2
    wb.save(XLSX)
    print(f"recalc OK: {XLSX}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
