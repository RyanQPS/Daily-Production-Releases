#!/usr/bin/env bash
# Entry point for the "run releases" cron job.
# Invokes Claude Code with the operating-instructions prompt.
set -euo pipefail
cd "$(dirname "$0")"
exec claude --print "run releases"
