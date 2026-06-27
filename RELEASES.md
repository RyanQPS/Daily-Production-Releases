# Daily Production Release Routine

Operating instructions for the **"run releases"** routine that Claude Code runs
on a cron (Mon–Fri 5:55 AM EST, America/New_York) and on demand.

A manual `run releases` is the authoritative run.

## What a "release" is

An RTP / production-release email sent by one of the five Project Managers to
Ryan. A job sitting in Ryan's inbox = **PENDING**. Once Ryan files it to a PM
folder, it must drop from the list.

## The five PMs (the ONLY people who release jobs) — Rule 7

Search each one's mail **BY SENDER**, scoped to Ryan's Inbox. Never use
per-job keyword search — keyword search ranks our own "Production Release Log"
report emails above the real RTPs and produces false drops.

| PM | Email |
| --- | --- |
| Ian Zanetta | `ian.zanetta@qualityprecastsolutions.com` |
| Keaton Elliott | `keaton@qualityprecastsolutions.com` |
| Lee Goodwin | `lee.goodwin@qualityprecastsolutions.com` |
| Rodrigo Gracia | `rodrigo.gracia@qualityprecastsolutions.com` |
| Maryanne Siurano | `maryanne.siurano@qualityprecastsolutions.com` |

## Release-subject filter — Rule 8

From each PM's results, KEEP an email only if BOTH are true:

1. Subject matches the release pattern: a job number like `NN-NNN`
   (e.g. `26-701`, `25-1671`) OR contains `RTP`, `RTP Package`, or
   `RELEASED to Prod`.
2. Subject does **NOT** start with `RE:`, `FW:`, or `Out of office`.

If a subject is ambiguous (e.g. a bare `RTP - <job>` that might be a thread
reply), OPEN it and decide from the body, then **FLAG** it in the summary —
never silently judge.

## Run procedure — four passes, drop before add (Rule 6)

1. Load `release_state.json` (the running list).
2. **Pass 1–2 (DROP):** for each of the 5 PMs, pull sent mail (newest first),
   apply the Rule 8 filter. Collect the set of release `internetMessageId`s
   currently in the inbox. For every job in `release_state.json` whose
   `messageId` (or any of its `messageIds`) is NOT in that set, DROP it. Match
   on `messageId + jobNumber`, never on name (Rule 2).
3. **Pass 3–4 (ADD):** from the same filtered results, any release whose
   `messageId` is not already in state is NEW — append it. De-dupe by
   `jobNumber` (e.g. `25-952` released by both Ian and Keaton is ONE row).
4. Save `release_state.json`.

## Output

1. Rebuild `Production_Release_Log.xlsx` (`build_log.py` then `recalc.py`).
2. Email the summary to Ryan, **CC `victor.adame@qualityprecastsolutions.com`
   (Rule 4)**, Excel **attached** (Rule 5). Use Composio `OUTLOOK_SEND_EMAIL`
   (or `CREATE_DRAFT` → `ADD_MAIL_ATTACHMENT` → `SEND_DRAFT` for attachments).
3. The email body must include: total pending, new-since-last-run rows
   highlighted, oldest job flagged, and any ambiguous/flagged items called out.

## Confirmed open item

The PM-sender search has been **inbox-scoped** to date — none of the example
filed jobs (e.g. Moccasin Wallow `26-989`) appeared in any PM's results, so
drops worked. If that ever changes, the drop logic breaks; re-verify with a
known-filed job and adjust.

## Full rule set

1. Running list — append only.
2. No duplicates — match `jobNumber + messageId`, not name.
3. Auto-drop — release leaves the list when its RTP email leaves the inbox.
4. Always CC Vic.
5. Always attach the Excel.
6. Four-pass — drop existing first, then add new.
7. Verify by PM sender, never keyword.
8. Subject must match release pattern AND not be a `RE:`/`FW:`/Out-of-office reply.

## Files

- `release_state.json` — running list (running list, append-only with drops).
- `build_log.py` — renders `Production_Release_Log.xlsx` from
  `release_state.json` (no formulas — all values pre-computed for zero
  formula errors).
- `recalc.py` — opens the xlsx, scans for `#REF!`/`#NAME?`/`#VALUE!`/etc.
  error markers, and re-saves. Exits non-zero on any error.
- `run_releases.sh` — cron entrypoint (`claude --print "run releases"`).

## Schedule

Mon–Fri 5:55 AM EST via cron (`run_releases.sh`). System TZ =
`America/New_York`.
