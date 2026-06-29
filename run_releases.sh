#!/usr/bin/env bash
# Wrapper used by cron Mon-Fri 5:55 AM EST (America/New_York).
# Invokes Claude Code with the "run releases" prompt; CLAUDE.md governs behavior.
set -euo pipefail
cd "$(dirname "$0")"
claude --print "run releases"
