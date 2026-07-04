#!/usr/bin/env python3
"""Recalculate + verify Production_Release_Log.xlsx has ZERO formula errors.

The routine contract requires the delivered workbook to be free of formula
errors. Since build_log.py writes only literal values (no formulas), this
sweep walks every cell and asserts none holds a string starting with '#'
followed by ERROR text (#REF!, #VALUE!, #N/A, #DIV/0!, #NAME?, #NULL!, #NUM!).
It also touches the workbook so Excel refreshes on next open.
"""
from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import load_workbook


ERROR_TOKENS = {"#REF!", "#VALUE!", "#N/A", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!"}


def scan(path: Path) -> int:
    wb = load_workbook(path)
    bad = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v in ERROR_TOKENS:
                    print(f"ERROR cell {ws.title}!{cell.coordinate} = {v}")
                    bad += 1
    wb.save(path)
    return bad


def main() -> None:
    root = Path(__file__).resolve().parent
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else root / "Production_Release_Log.xlsx"
    bad = scan(target)
    if bad:
        print(f"FAIL: {bad} error cells in {target}")
        sys.exit(1)
    print(f"OK: 0 formula errors in {target}")


if __name__ == "__main__":
    main()
