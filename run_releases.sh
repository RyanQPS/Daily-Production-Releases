#!/usr/bin/env bash
# Cron-invoked entrypoint for the daily production release routine.
# Schedule: Mon-Fri 05:55 EST (America/New_York).
# The actual work — querying inbox, updating release_state.json, rebuilding
# the xlsx, and emailing the summary — is governed by CLAUDE.md and runs
# inside `claude --print "run releases"`. This script keeps the cron line
# simple and serves as the single hook into the routine.
set -euo pipefail
cd "$(dirname "$0")"
exec claude --print "run releases"
