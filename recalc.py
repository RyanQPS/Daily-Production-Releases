#!/usr/bin/env python3
"""Force Excel to recalculate on next open and verify no #REF/#VALUE errors.

The Excel produced by build_log.py contains no formulas (all cells are literal
values), so the only 'recalc' work is flipping the workbook's calc_on_load flag
and scanning every cell for error strings before we ship the file.
"""
import sys
from pathlib import Path
from openpyxl import load_workbook

HERE = Path(__file__).parent
OUT_PATH = HERE / "Production_Release_Log.xlsx"

ERROR_TOKENS = ("#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#NULL!", "#N/A", "#NUM!")


def main():
    wb = load_workbook(OUT_PATH)
    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True

    errors = []
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=False):
            for cell in row:
                v = cell.value
                if isinstance(v, str) and any(t in v for t in ERROR_TOKENS):
                    errors.append(f"{ws.title}!{cell.coordinate}: {v!r}")

    if errors:
        print("FORMULA ERRORS FOUND:")
        for e in errors:
            print(f"  {e}")
        sys.exit(1)

    wb.save(OUT_PATH)
    print(f"Recalc OK: {OUT_PATH} has zero formula errors.")


if __name__ == "__main__":
    main()
