#!/usr/bin/env python3
"""Re-open Production_Release_Log.xlsx and verify it round-trips cleanly.

build_log.py writes only static values (no formulas), so there is nothing
to recalculate — but the operating instructions explicitly call recalc.py
after build_log.py. This script:

  1. Loads the workbook with openpyxl (which fails loudly if the file is
     corrupt).
  2. Scans every cell for the "#REF!", "#NAME?", "#VALUE!", "#DIV/0!" or
     "#N/A" error markers and exits non-zero if any are found.
  3. Re-saves the file so the workbook reflects a clean round-trip.
"""

from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parent
PATH = ROOT / "Production_Release_Log.xlsx"

ERROR_MARKERS = ("#REF!", "#NAME?", "#VALUE!", "#DIV/0!", "#N/A", "#NULL!", "#NUM!")


def main() -> int:
    if not PATH.exists():
        print(f"ERROR: {PATH} not found — run build_log.py first.", file=sys.stderr)
        return 2
    wb = load_workbook(PATH)
    errors: list[str] = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                if isinstance(cell.value, str) and any(
                    m in cell.value for m in ERROR_MARKERS
                ):
                    errors.append(f"{ws.title}!{cell.coordinate}: {cell.value!r}")
    if errors:
        print("FORMULA ERRORS FOUND:", file=sys.stderr)
        for e in errors:
            print(f"  {e}", file=sys.stderr)
        return 1
    wb.save(PATH)
    print(f"Round-trip OK: {PATH} — 0 formula errors.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
