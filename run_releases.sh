#!/usr/bin/env bash
# Cron entrypoint for the Daily Production Releases routine.
# System TZ is expected to be America/New_York; scheduled Mon-Fri 5:55 AM ET.
# The actual work is driven by Claude Code via `claude --print "run releases"`;
# this shell wrapper only exists so the crontab has a stable command.
set -euo pipefail
cd "$(dirname "$0")"
exec claude --print "run releases"
