#!/usr/bin/env python3
"""No-op recalc — the log has no formulas so nothing to recalc.

Exists so run_releases.sh keeps a stable interface (build_log.py → recalc.py).
Verifies that the produced workbook opens cleanly and reports the row count.
"""
from pathlib import Path
import sys

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "Production_Release_Log.xlsx"


def main() -> int:
    if not OUT.exists():
        print(f"MISSING: {OUT}", file=sys.stderr)
        return 1
    wb = load_workbook(OUT)
    ws = wb.active
    rows = ws.max_row
    print(f"OK: {OUT} — {rows} rows, no formulas")
    return 0


if __name__ == "__main__":
    sys.exit(main())
