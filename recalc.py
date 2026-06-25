#!/usr/bin/env python3
"""Verify Production_Release_Log.xlsx has zero formula errors.

build_log.py writes all totals as static Python-computed values rather
than Excel formulas, so the recalc step is a check: walk every cell and
fail if any string cell looks like a formula or an Excel error token.
"""

import sys
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
OUT_FILE = ROOT / "Production_Release_Log.xlsx"

ERROR_TOKENS = {"#REF!", "#NAME?", "#VALUE!", "#DIV/0!", "#NULL!", "#N/A", "#NUM!"}


def main():
    if not OUT_FILE.exists():
        print(f"ERROR: {OUT_FILE} not found - run build_log.py first.", file=sys.stderr)
        sys.exit(1)
    wb = load_workbook(OUT_FILE, data_only=False)
    issues = []
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str):
                    if v.startswith("="):
                        issues.append(f"{sheet.title}!{cell.coordinate}: unexpected formula {v!r}")
                    if v in ERROR_TOKENS:
                        issues.append(f"{sheet.title}!{cell.coordinate}: error token {v}")
    if issues:
        print("FORMULA / ERROR ISSUES:", file=sys.stderr)
        for issue in issues:
            print(f"  - {issue}", file=sys.stderr)
        sys.exit(2)
    print(f"OK: {OUT_FILE} - zero formulas, zero error tokens.")


if __name__ == "__main__":
    main()
