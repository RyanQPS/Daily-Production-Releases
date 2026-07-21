# Production Release Routine — Operating Instructions

This file tells Claude Code how to execute "run releases". The scheduled cron
job runs `claude --print "run releases"`, so this file governs the automated
run. Follow it step for step. A manual "run releases" is the authoritative run.

## What a "release" is
An RTP / production-release email sent by one of the five Project Managers to
Ryan. A job sitting in Ryan's inbox = **PENDING**. Once Ryan files it to a PM
folder it must drop.

## The five PMs (search BY SENDER — Rule 7)
- ian.zanetta@qualityprecastsolutions.com
- keaton@qualityprecastsolutions.com
- lee.goodwin@qualityprecastsolutions.com
- rodrigo.gracia@qualityprecastsolutions.com
- maryanne.siurano@qualityprecastsolutions.com

## Release-subject filter (Rule 8)
Keep a message only if BOTH are true:
1. Subject matches the release pattern: a job number `NN-NNN`
   OR contains `RTP`, `RTP Package`, or `RELEASED to Prod`.
2. Subject does NOT start with `RE:`, `FW:`, or `Out of office`.

Ambiguous subjects (e.g. bare "RTP - <job>" that might be a thread reply) are
opened, decided from the body, and **flagged** in the summary — never silently
judged.

## Run procedure — four passes, drop before add (Rule 6)
1. Load `release_state.json` (running list).
2. **DROP (passes 1-2):** for each PM, pull sent mail (newest first) via
   Outlook sender search, apply Rule 8. Any entry in state whose `messageId`
   is not in the current inbox set is DROPPED (it was filed). Match on
   `messageId` + `jobNumber`, never on name (Rule 2).
3. **ADD (passes 3-4):** for the same filtered results, append any release
   whose `messageId` is not already in state. Parse `jobNumber`, `name`,
   `contractor`, `contact`, `phone`, PM(s), structure counts from subject.
   Counts only in the PDF → `See Notes`.
4. De-dupe: the same job released by two PMs is ONE row.
5. Save `release_state.json`.

## Output
- Rebuild `Production_Release_Log.xlsx` (`python3 build_log.py && python3 recalc.py`).
- Email summary to `ryan.levesque@qualityprecastsolutions.com`, CC
  `victor.adame@qualityprecastsolutions.com` (Rule 4), Excel attached (Rule 5).
  Use Composio Outlook via `OUTLOOK_CREATE_DRAFT` + `OUTLOOK_ADD_MAIL_ATTACHMENT`
  + `OUTLOOK_SEND_DRAFT` (or `OUTLOOK_SEND_EMAIL` when the attachment fits the
  single-shot payload).
- Include: total pending, new-since-last-run rows highlighted, oldest job
  flagged, any ambiguous/flagged items called out.

## Open item to verify once
Confirm the sender search is inbox-only. Suggested test: search Ian's sent
mail for the known-filed job Moccasin Wallow 26-989; it must NOT appear. If
it does, the drop logic breaks.

## Full rule set
1. Running list — append only.
2. No duplicates — match `jobNumber` + `messageId`, not name.
3. Auto-drop — release leaves the list when its RTP email leaves the inbox.
4. Always CC Vic.
5. Always attach the Excel.
6. Four-pass — drop existing first, then add new.
7. Verify by PM sender, never keyword.
8. Subject must match release pattern AND not be a RE:/FW:/Out-of-office reply.

## Schedule
Mon-Fri 5:55 AM EST via cron (`run_releases.sh`). System TZ =
`America/New_York`.
