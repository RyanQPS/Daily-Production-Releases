#!/usr/bin/env bash
# Cron entry point for "run releases".
# Delegates to Claude Code non-interactively so the routine executes the
# operating instructions in `RELEASE_ROUTINE.md` end-to-end.
set -euo pipefail
cd "$(dirname "$0")"
exec claude --print "run releases"
