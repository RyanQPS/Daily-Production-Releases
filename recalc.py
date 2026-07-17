#!/usr/bin/env python3
"""Recalculate Production_Release_Log.xlsx via headless LibreOffice.

build_log.py writes only literal values, but this pass makes the file byte-clean
and refreshes cached formula metadata so Excel opens with zero errors. It's a
no-op if LibreOffice is unavailable; the file remains valid either way.
"""
from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
XLSX = ROOT / "Production_Release_Log.xlsx"


def main() -> int:
    if not XLSX.exists():
        print(f"missing {XLSX}; run build_log.py first", file=sys.stderr)
        return 1

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    ran_libreoffice = False
    if soffice:
        outdir = ROOT / ".recalc_tmp"
        outdir.mkdir(exist_ok=True)
        cmd = [
            soffice,
            "--headless",
            "--calc",
            "--convert-to",
            "xlsx",
            "--outdir",
            str(outdir),
            str(XLSX),
        ]
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
            produced = outdir / XLSX.name
            if result.returncode == 0 and produced.exists():
                produced.replace(XLSX)
                print(f"Recalculated {XLSX} via LibreOffice")
                ran_libreoffice = True
            else:
                print(
                    "LibreOffice recalc failed (returncode="
                    f"{result.returncode}); falling back to openpyxl resave.",
                    file=sys.stderr,
                )
                if result.stderr:
                    print(result.stderr, file=sys.stderr)
        except subprocess.TimeoutExpired:
            print("LibreOffice recalc timed out; falling back to openpyxl resave.", file=sys.stderr)
        finally:
            shutil.rmtree(outdir, ignore_errors=True)

    if not ran_libreoffice:
        try:
            from openpyxl import load_workbook
        except ImportError:
            print("openpyxl missing; leaving file untouched.", file=sys.stderr)
            return 0
        wb = load_workbook(XLSX)
        wb.save(XLSX)
        print(f"Recalculated {XLSX} via openpyxl resave (no formulas to compute).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
