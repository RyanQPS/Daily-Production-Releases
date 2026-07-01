#!/usr/bin/env python3
"""Recompute derived totals in state and rewrite the xlsx.

Idempotent — safe to run after every state edit. Ensures storm/san totals
and 'tot' fields agree with the underlying integer values.

Usage:  python scripts/recalc.py
"""
from __future__ import annotations

import json
from pathlib import Path

import build_log

ROOT = Path(__file__).resolve().parent.parent
STATE = ROOT / "state" / "release_state.json"


def _to_int(v):
    return v if isinstance(v, int) else 0


def recalc(state: dict) -> dict:
    for row in state["pending"]:
        s = _to_int(row["storm"])
        sa = _to_int(row["san"])
        if row["tot"] != "See Notes":
            row["tot"] = s + sa if (s + sa) else row["tot"]
    return state


def main() -> None:
    state = json.loads(STATE.read_text())
    state = recalc(state)
    STATE.write_text(json.dumps(state, indent=2, ensure_ascii=False))
    build_log.main()
    print("Recalc done.")


if __name__ == "__main__":
    main()
