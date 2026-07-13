#!/usr/bin/env python3
"""Open Production_Release_Log.xlsx and verify no formula errors.

The build script emits literals only, so this check should always pass;
if it ever fails, we've regressed to formulas and need to fix that.
"""
from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import load_workbook

REPO = Path(__file__).resolve().parent
XLSX = REPO / "Production_Release_Log.xlsx"

ERROR_MARKERS = ("#REF!", "#NAME?", "#DIV/0!", "#VALUE!", "#N/A", "#NULL!", "#NUM!")


def main() -> int:
    wb = load_workbook(XLSX, data_only=False)
    errors = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if v is None:
                    continue
                if isinstance(v, str) and v.startswith("=") and cell.data_type == "f":
                    errors.append((ws.title, cell.coordinate, "formula present", v))
                if isinstance(v, str) and any(m in v for m in ERROR_MARKERS):
                    errors.append((ws.title, cell.coordinate, "error marker", v))
    if errors:
        for sheet, coord, kind, val in errors:
            print(f"{sheet}!{coord}\t{kind}\t{val!r}", file=sys.stderr)
        return 1
    rows = wb["Pending RTPs"].max_row - 1
    print(f"ok: no formula errors; {rows} data rows in Pending RTPs")
    return 0


if __name__ == "__main__":
    sys.exit(main())
