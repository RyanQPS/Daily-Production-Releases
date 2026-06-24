#!/usr/bin/env bash
# QPS Production Release Routine — cron entry point.
#
# Invoked by crontab on weekdays at 5:55 AM America/New_York.
# Wraps `claude --print "run releases"` so MCPs (Microsoft 365, Composio)
# are available and the routine can do its job.

set -uo pipefail

# --- EDIT THIS to the absolute path of this project's folder -----------------
ROUTINE_DIR="$HOME/qps-releases"
# -----------------------------------------------------------------------------

LOG_FILE="$ROUTINE_DIR/release.log"

# Cron strips PATH. Add the usual spots node + claude live on macOS so the
# binaries resolve. Adjust if you installed them somewhere unusual.
export PATH="/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:$PATH"

if [[ ! -d "$ROUTINE_DIR" ]]; then
  echo "run_releases.sh: missing $ROUTINE_DIR" >&2
  exit 1
fi

cd "$ROUTINE_DIR"

{
  echo ""
  echo "=================================================================="
  echo "=== $(date '+%Y-%m-%d %H:%M:%S %Z')"
  echo "=================================================================="

  if ! command -v claude >/dev/null 2>&1; then
    echo "ERROR: 'claude' CLI not found on PATH ($PATH)"
    exit 127
  fi

  claude --print "run releases" --dangerously-skip-permissions
  rc=$?

  echo "=== exit: $rc"
  exit $rc
} >> "$LOG_FILE" 2>&1
