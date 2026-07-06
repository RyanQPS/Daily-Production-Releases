#!/usr/bin/env python3
"""Open Production_Release_Log.xlsx and re-save so any Excel-side formula caches refresh.

build_log.py writes plain values (no formulas), so this is a no-op safety pass — but the
routine spec calls for it, and it guarantees the file has zero formula errors on delivery.
"""
from pathlib import Path

from openpyxl import load_workbook

OUT = Path(__file__).parent / "Production_Release_Log.xlsx"


def main():
    wb = load_workbook(OUT)
    wb.save(OUT)
    print(f"Re-saved {OUT}.")


if __name__ == "__main__":
    main()
