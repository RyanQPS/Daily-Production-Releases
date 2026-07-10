#!/usr/bin/env bash
# Cron entrypoint: `claude --print "run releases"`.
# This script exists as a placeholder that documents what the cron job runs
# and can be extended for local execution or logging in the future.
set -euo pipefail
cd "$(dirname "$0")"
claude --print "run releases"
