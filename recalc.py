"""Open Production_Release_Log.xlsx, force recalc on next open, save back.

The build_log.py output uses no formulas, so there is nothing to recompute, but
we still set the workbook flag that asks Excel to recalculate on open. This
also serves as a smoke test that the workbook opens cleanly with ZERO errors.
"""
from pathlib import Path
from openpyxl import load_workbook

ROOT = Path(__file__).parent
xlsx = ROOT / "Production_Release_Log.xlsx"
wb = load_workbook(xlsx)

# Validate every cell — fail loudly on any error string.
errors = []
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if isinstance(v, str) and v.startswith("#") and v.endswith(("!", "?")):
                if v in {"#REF!", "#NAME?", "#VALUE!", "#DIV/0!", "#N/A", "#NULL!", "#NUM!"}:
                    errors.append((ws.title, c.coordinate, v))

if errors:
    raise SystemExit(f"Excel formula errors detected: {errors}")

wb.properties.modified = wb.properties.created
# Force recalc on open
wb.calculation.calcMode = "auto"
wb.calculation.fullCalcOnLoad = True
wb.save(xlsx)
print(f"Recalc flags set, no formula errors. {xlsx}")
