#!/usr/bin/env python3
"""Open Production_Release_Log.xlsx, force a formula recalc on next open,
and verify no formula-error cells remain. On this workbook there are no
formulas — this is a safety net that will surface issues if any are added.
"""
import sys
from pathlib import Path

from openpyxl import load_workbook

REPO = Path(__file__).resolve().parent
XLSX = REPO / "Production_Release_Log.xlsx"

ERROR_TOKENS = {"#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!"}


def main():
    wb = load_workbook(XLSX)
    # Tell Excel to fully recalc on next open.
    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True

    errors = []
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value in ERROR_TOKENS:
                    errors.append(f"{sheet.title}!{cell.coordinate} = {cell.value}")

    if errors:
        print("Formula errors detected:", file=sys.stderr)
        for e in errors:
            print(f"  {e}", file=sys.stderr)
        sys.exit(1)

    wb.save(XLSX)
    print(f"Recalc-on-open enabled; 0 formula errors in {XLSX.name}")


if __name__ == "__main__":
    main()
