#!/usr/bin/env bash
# Cron entry point — Mon-Fri 5:55 AM ET.
# This is a stub; the actual routine is driven by Claude Code per the
# operating instructions at the top of the repo README / task prompt.
set -euo pipefail
cd "$(dirname "$0")"
exec claude --print "run releases"
