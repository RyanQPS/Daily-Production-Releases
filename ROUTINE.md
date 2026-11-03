# Production Release Routine — Operating Instructions

This file tells Claude Code exactly how to execute "run releases". The scheduled
cron job runs `claude --print "run releases"`, so THIS FILE is what governs the
automated run. Follow it step for step. A manual "run releases" is the authoritative run.

## What a "release" is

An RTP / production-release email sent by one of the five Project Managers to Ryan.
A job sitting in Ryan's inbox = PENDING. Once Ryan files it to a PM folder, it must drop.

## The five PMs (the ONLY people who release jobs) — Rule 7

Search each one's sent mail BY SENDER. Never use per-job keyword search — keyword
search ranks our own "Production Release Log" report emails above the real RTPs and
produces false drops. Senders:

- ian.zanetta@qualityprecastsolutions.com
- keaton@qualityprecastsolutions.com
- lee.goodwin@qualityprecastsolutions.com
- rodrigo.gracia@qualityprecastsolutions.com
- maryanne.siurano@qualityprecastsolutions.com

## Release-subject filter — Rule 8

From each PM's results, KEEP an email only if BOTH are true:

- Subject matches the release pattern: a job number like `NN-NNN` (e.g. `26-701`,
  `25-1671`) OR contains "RTP", "RTP Package", or "RELEASED to Prod".
- Subject does NOT start with `RE:`, `FW:`, or `Out of office`.

If a subject is ambiguous (e.g. a bare "RTP - <job>" that might be a thread reply),
OPEN it and decide from the body, then FLAG it in the summary — never silently judge.

## Run procedure — four passes, drop before add (Rule 6)

1. Load `release_state.json` (the running list: each entry has `jobNumber`, `name`,
   `contractor`, `pm`, `contact`, `phone`, `storm`, `san`, `received`, `notes`, `messageId`).
2. Pass 1-2 (DROP): For each of the 5 PMs, pull sent mail (newest first), apply the
   Rule 8 filter. Collect the set of release `internetMessageId`s currently in the inbox.
   For every job in `release_state.json` whose `messageId` is NOT in that set, DROP it
   (it was filed). Match on `messageId + jobNumber`, never on name (Rule 2).
3. Pass 3-4 (ADD): From the same filtered results, any release whose `messageId` is not
   already in the state is NEW — append it (parse `jobNumber`, `name`, `contractor`, `contact`,
   `phone`, structure counts from the subject; counts only in the PDF -> "See Notes").
   De-dupe: the same job released by two PMs (e.g. 25-952 from Ian and Keaton) is ONE row.
4. Save `release_state.json`.

## Output

- Rebuild `Production_Release_Log.xlsx` from current state (run `build_log.py`, then
  `recalc.py` — deliver with ZERO formula errors).
- Email the summary to `ryan.levesque@qualityprecastsolutions.com`, CC
  `victor.adame@qualityprecastsolutions.com` (Rule 4), Excel ATTACHED (Rule 5).
  Use Composio `OUTLOOK_SEND_EMAIL` for the send.
- In the email: total pending, new-since-last-run rows highlighted, oldest job flagged,
  and any ambiguous/flagged items called out.

## Open item to verify once

Confirm whether the sender search is inbox-only or searches ALL folders. Today it
behaved as inbox-scoped (filed jobs did NOT appear in PM results, so drops worked).
If it ever returns filed mail, the drop logic breaks — test once with a known-filed
job (e.g. Moccasin Wallow 26-989) and confirm it does NOT appear.

## Full rule set

1. Running list — append only.
2. No duplicates — match `jobNumber + messageId`, not name.
3. Auto-drop — release leaves the list when its RTP email leaves the inbox.
4. Always CC Vic.
5. Always attach the Excel.
6. Four-pass — drop existing first, then add new.
7. Verify by PM sender, never keyword.
8. Subject must match release pattern AND not be a `RE:`/`FW:`/Out-of-office reply.

## Schedule

Mon-Fri 5:55 AM EST via cron (`run_releases.sh`). System TZ = `America/New_York`.
