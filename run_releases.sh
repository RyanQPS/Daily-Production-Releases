#!/usr/bin/env bash
# Cron entrypoint for the daily production-release routine.
# Invoked Mon-Fri 5:55 AM EST. System TZ = America/New_York.
set -euo pipefail
cd "$(dirname "$0")"
exec claude --print "run releases"
