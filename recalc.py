#!/usr/bin/env python3
"""Force a recalc pass on Production_Release_Log.xlsx.

The workbook is formula-free (all values are literals), so this only sanity-checks
that every cell parses as a value and no ``#`` error strings are present.
"""
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
XLSX = ROOT / "Production_Release_Log.xlsx"

ERROR_TOKENS = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!")


def recalc() -> None:
    wb = load_workbook(XLSX)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                if isinstance(cell.value, str) and cell.value in ERROR_TOKENS:
                    raise SystemExit(f"Formula error found at {ws.title}!{cell.coordinate}")
    wb.save(XLSX)
    print(f"Recalc OK: {XLSX}")


if __name__ == "__main__":
    recalc()
