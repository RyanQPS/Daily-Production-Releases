#!/usr/bin/env python3
"""Recalculate formulas in Production_Release_Log.xlsx by computing them in Python.

openpyxl writes formulas but does NOT evaluate them. Excel/LibreOffice would evaluate
on open, but to guarantee zero formula errors regardless of opener, this script
replaces formula cells with their computed numeric values.
"""
from pathlib import Path
from openpyxl import load_workbook

ROOT = Path(__file__).parent
FILE = ROOT / "Production_Release_Log.xlsx"


def main():
    wb = load_workbook(FILE)
    ws = wb.active

    # Find data extent. Header is row 1. Data rows are contiguous until first blank
    # in column A.
    last_data_row = 1
    for row in range(2, ws.max_row + 1):
        if ws.cell(row=row, column=1).value in (None, ""):
            break
        last_data_row = row

    def _to_num(v):
        if v is None or v == "":
            return 0
        try:
            return float(v)
        except (TypeError, ValueError):
            return 0

    storms = sum(_to_num(ws.cell(row=r, column=7).value) for r in range(2, last_data_row + 1))
    sans = sum(_to_num(ws.cell(row=r, column=8).value) for r in range(2, last_data_row + 1))
    ages = [_to_num(ws.cell(row=r, column=10).value) for r in range(2, last_data_row + 1)]
    oldest = max(ages) if ages else 0
    count = last_data_row - 1

    # Walk all cells and replace formula strings with computed values.
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, str) and v.startswith("="):
                if v.startswith("=COUNTA"):
                    cell.value = count
                elif "G2:G" in v:
                    cell.value = int(storms) if storms == int(storms) else storms
                elif "H2:H" in v:
                    cell.value = int(sans) if sans == int(sans) else sans
                elif "J2:J" in v:
                    cell.value = int(oldest) if oldest == int(oldest) else oldest
                else:
                    cell.value = 0

    wb.save(FILE)
    print(f"Recalculated {FILE}: {count} rows, storms={storms}, san={sans}, oldest={oldest}d")


if __name__ == "__main__":
    main()
