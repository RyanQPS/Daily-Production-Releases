# Daily Production Releases

Scheduled Claude Code routine — surfaces PM-released jobs (RTPs) that are still
sitting in Ryan's inbox and hands the summary + Excel log to Ryan (CC Vic) each
weekday morning.

- Cron: **Mon–Fri 5:55 AM EST** (`America/New_York`) via `run_releases.sh`.
- Entry point: `claude --print "run releases"`.
- Operating rules and full procedure: see the initial prompt attached to the
  routine (Rules 1–8, four-pass, drop-before-add, PM sender verify).

## Layout

```
state/release_state.json          # running list (append-only, dropped when filed)
scripts/build_log.py              # generates output/Production_Release_Log.xlsx from state
scripts/recalc.py                 # recomputes totals + rebuilds xlsx (formula-free)
output/Production_Release_Log.xlsx# emailed attachment
run_releases.sh                   # cron entry point
```

## What each run does

1. Load `state/release_state.json`.
2. Pull each of the 5 PMs' inbox-visible sent mail (sender-filtered, order:newest).
3. Apply Rule 8 subject filter (release pattern, not RE:/FW:/Out-of-office).
4. **DROP** state rows whose messageId is no longer visible (job filed).
5. **ADD** new releases whose messageId is not yet in state.
6. Regenerate `output/Production_Release_Log.xlsx` (`build_log.py` then
   `recalc.py`, no formulas so zero `#VALUE!` errors).
7. Email `ryan.levesque@qualityprecastsolutions.com` CC
   `victor.adame@qualityprecastsolutions.com` with the Excel attached.

## Known caveats

- **Maryanne Siurano** — `sender + order:newest` returns 0 in her mailbox scope.
  Fall back to an Inbox-scoped query for her.
- **Inbox-scoped sender search** verified by the Moccasin Wallow 26-989 test
  (filed job does NOT appear in Ian's sender results). Drop logic is safe.
