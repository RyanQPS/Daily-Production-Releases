#!/usr/bin/env bash
# Cron entrypoint. Executed Mon-Fri 5:55 AM America/New_York.
# Invokes Claude Code with "run releases" — CLAUDE.md governs the routine.
set -euo pipefail
cd "$(dirname "$0")"
exec claude --print "run releases"
