"""
recalc.py — Open the xlsx, verify every cell parses, resave.

The rebuild uses only literal values (no formulas) so this pass is a
sanity check: openpyxl reads/writes each cell and confirms zero formula
errors before the file lands in Ryan's inbox.
"""
from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).parent
OUT = ROOT / "Production_Release_Log.xlsx"


def main() -> None:
    wb = load_workbook(OUT)
    errors = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=False):
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v.startswith("#") and v.endswith(
                    ("!", "?")
                ) and v in {"#REF!", "#NAME?", "#VALUE!", "#DIV/0!", "#N/A", "#NULL!", "#NUM!"}:
                    errors += 1
                    print(f"  ERR at {ws.title}!{cell.coordinate}: {v}")
    wb.save(OUT)
    print(f"Recalc complete — {errors} formula errors.")


if __name__ == "__main__":
    main()
