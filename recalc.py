#!/usr/bin/env python3
"""Recalculate any formulas in Production_Release_Log.xlsx via LibreOffice.

build_log.py writes only computed values, so this is a safety pass that
guarantees Excel will not display #N/A / #REF! on first open. If
LibreOffice is unavailable, exits cleanly — the workbook still opens.
"""
from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent
XLSX = REPO / "Production_Release_Log.xlsx"


def main() -> int:
    if not XLSX.exists():
        print(f"recalc: {XLSX} missing — nothing to do", file=sys.stderr)
        return 1

    soffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not soffice:
        print("recalc: LibreOffice not present — skipping (no formulas to recompute)")
        return 0

    # Headless convert-to xlsx forces a recalc-and-save pass.
    cmd = [
        soffice, "--headless", "--calc",
        "--convert-to", "xlsx",
        "--outdir", str(REPO),
        str(XLSX),
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    if proc.returncode != 0:
        print("recalc: libreoffice failed:", proc.stderr, file=sys.stderr)
        return proc.returncode
    print(f"recalc: refreshed {XLSX.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
