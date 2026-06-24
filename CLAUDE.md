# QPS Production Release Routine

This project tracks Quality Precast Solutions production-release emails
(RTP / NTO / "release" notices) that land in Ryan's Outlook inbox, keeps
a rolling Excel log of the currently-open releases, and emails a daily
summary to Ryan and Vic.

It is driven by Claude Code — `run_releases.sh` calls
`claude --print "run releases"` from cron, and Claude follows the steps
below.

---

## When the user (or cron) says "run releases"

Run these steps in order. Log progress to stdout — everything is captured
to `release.log` by the wrapper.

### 1. Make sure the spreadsheet exists

```
node releases.mjs init
```

This is a no-op if `releases.xlsx` already exists.

### 2. Read the prior state

```
node releases.mjs read
```

Returns the current rows as JSON. Keep this in mind so the summary email
can call out *what changed* (new this run vs. dropped this run).

### 3. Pull the current inbox snapshot from Outlook

Use the **Microsoft-365** MCP (`outlook_email_search`). What counts as a
release email:

- The message is currently in the **Inbox** folder (not archived,
  not filed into a sub-folder). If a message has been moved out of the
  Inbox, it MUST NOT appear in the new log — that is how Ryan "closes"
  a release.
- The subject or body contains one of:
  `RTP`, `NTO`, `release`, `ready to publish`, `notice to occupy`,
  `notice of occupancy`, `production release`.
- Look back ~60 days. Older messages still in the inbox should still
  be included.

For each matching message, capture:

| field      | source                                                   |
|------------|----------------------------------------------------------|
| Date       | the email's received date (YYYY-MM-DD, ET)               |
| Type       | RTP / NTO / Release (pick whichever the subject implies) |
| Project    | the job / project name, parsed from the subject or body  |
| Sender     | sender display name                                      |
| Subject    | the subject line, trimmed                                |
| MessageId  | the Outlook message id (used for de-dup across runs)     |

If the same project appears in multiple inbox emails, keep ONE row —
the most recent — and note older entries in the body of the summary
email rather than the spreadsheet.

### 4. Write the new state back

Write the row list to `/tmp/releases.json` then:

```
node releases.mjs replace /tmp/releases.json
```

This **replaces** the sheet. Anything that was in the prior log but is no
longer in the inbox is dropped automatically — that is the intended
behaviour.

### 5. Send the summary email

Use the **Composio** MCP (`OUTLOOK_SEND_EMAIL` or equivalent). Look up
the right tool with `mcp__Composio__COMPOSIO_SEARCH_TOOLS` if unsure.

- **To:** `ryan.levesque@qualityprecastsolutions.com`
- **Cc:** Vic — look up his address from a recent inbox email
  (search "Vic" in the From field). If you cannot find it with
  reasonable confidence, send to Ryan only and call that out in the body.
  Do NOT guess an address.
- **Subject:** `QPS Daily Production Releases — YYYY-MM-DD`
- **Body:** plaintext / lightweight markdown. Include:
  - Total open releases by type (e.g. "4 RTP, 2 NTO, 1 other")
  - A bulleted list of project names with date received
  - Any releases NEW this run (compared to step 2's snapshot)
  - Any releases DROPPED this run (in the prior log but no longer in
    the inbox — Ryan filed them)
  - A note if Vic's address could not be confirmed
- **Attachment:** `releases.xlsx`

### 6. Send a routine summary push

End with a `PushNotification` (status: `proactive`) wrapped in
`<routine_summary>` tags. Lead sentence is what lands on Ryan's phone:

- If anything is new or anything dropped: lead with that count.
- If the snapshot is identical to yesterday's: a one-liner is fine
  ("No changes — 6 releases still open."), but the email still goes out.
- If the routine could not run (MCP unreachable, send failed, etc.):
  notify with the failure, **don't** swallow it.

---

## Helper CLI — `releases.mjs`

Small Node script around `xlsx`. Always run from this directory.

| command                              | what it does                                    |
|--------------------------------------|-------------------------------------------------|
| `node releases.mjs init`             | Create empty `releases.xlsx` if missing         |
| `node releases.mjs read`             | Print current rows as JSON                      |
| `node releases.mjs replace <file>`   | Replace the sheet with rows from a JSON file    |

Column order is fixed: `Date, Type, Project, Sender, Subject, MessageId`.

---

## Files

```
releases.mjs        spreadsheet helper (Node + xlsx)
run_releases.sh     cron wrapper that calls `claude --print "run releases"`
package.json        declares the xlsx dependency
CLAUDE.md           this file
CRON_SETUP.txt      one-time cron / timezone setup notes
releases.xlsx       generated; the rolling log (gitignored)
release.log         generated; cron stdout/stderr (gitignored)
```

## Conventions

- Never email Vic without confirming his address from the inbox.
- Never invent project names — if a subject is opaque, use the subject
  itself as the Project value and note it.
- All times are America/New_York.
- Don't push the spreadsheet to git — `.gitignore` already excludes it.
