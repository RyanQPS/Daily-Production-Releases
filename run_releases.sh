#!/usr/bin/env bash
# Cron entrypoint for the Daily Production Release routine.
# Schedule: Mon-Fri 5:55 AM EST (America/New_York).
#
# Invokes: claude --print "run releases"
# That run is what executes the operating instructions in RELEASES.md:
#   1. Load release_state.json
#   2. Pass 1-2 (DROP): for each of the 5 PMs, pull inbox-scoped mail BY SENDER,
#      apply Rule 8 subject filter; drop any state row whose messageId is no
#      longer in the inbox.
#   3. Pass 3-4 (ADD): from the same filtered results, append any new releases
#      whose messageId is not yet in state. De-dupe by jobNumber.
#   4. Rebuild Production_Release_Log.xlsx (build_log.py then recalc.py).
#   5. Email summary to Ryan, CC Vic, with the xlsx attached.

set -euo pipefail

cd "$(dirname "$0")"

claude --print "run releases"
