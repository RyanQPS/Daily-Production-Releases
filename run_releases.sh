#!/usr/bin/env bash
# Cron entry point — invoked by:  claude --print "run releases"
# Runs Mon-Fri at 5:55 AM EST (America/New_York).
set -euo pipefail
cd "$(dirname "$0")"
claude --print "run releases"
