# v0.2 Safe Intake — operator runbook

Safe Intake adds a **review gate** in front of the Gmail → Sheets → Drive workflow.
An operator reviews each detected invoice / receipt in the `Review` sheet and
explicitly sets **APPROVE**, **HOLD** or **SKIP**. Only a valid APPROVE can save PDFs
to Drive, append a `Records` row or apply the Gmail processed label.

> Scope: synthetic-data, locally tested demo. This has **not** been run against a real
> Google account by this change, is not deployed, and makes no accounting, tax or
> compliance claim. Detected totals are best-effort text matches, not verified figures.

## Entry points

| Function | Menu item | Writes |
|---|---|---|
| `setupSafeIntake()` | Safe Intake: 1. Set up review sheets | Creates `Review` (+ legacy sheets), adds `Records` columns K:M headers, sets Records B:C to plain text |
| `refreshReviewQueue()` | Safe Intake: 2. Refresh review queue | `Review` sheet only — never Drive, Records or Gmail |
| `processApprovedReviewRows()` | Safe Intake: 3. Process APPROVE rows | Drive PDFs, `Records`, Gmail label — **only for valid APPROVE rows** |
| `auditSafeIntake()` | Safe Intake: Audit / reconciliation report | `Audit` sheet only (reads Records, Review, Drive folder). The `Audit` sheet belongs to the script: it is cleared on every run, and the audit refuses to run if a sheet with that name holds other content |

**Backward compatibility.** `previewInvoiceEmails()` and `processInvoiceEmails()` still
exist and behave as before. The legacy `processInvoiceEmails()` is **not review-gated
and does not take the script lock**: it processes every matching message, including
ones you set to HOLD or SKIP. It does skip messages already in `Records` (including
ones written by Safe Intake), so it will not duplicate them. For a review-gated setup,
do not run the legacy process entry point; consider removing its menu item in a client
copy. The v0.2 functions never call it.

The menu (`onOpen`) only appears when the script is **bound to the spreadsheet**
(Extensions → Apps Script). In a standalone project, run the functions from the editor.

## Review sheet

| Column | Who edits | Meaning |
|---|---|---|
| Message ID, Thread ID, Sender, Email Date, Subject | system | Source identity (stored as plain text) |
| Detected Total, Detected Currency, PDF Count | system | Best-effort extraction; **do not edit** — use the correction columns |
| Revision | system | SHA-256 fingerprint of the message content, PDFs and extraction result |
| **Decision** | operator | `APPROVE`, `HOLD`, `SKIP` or blank (dropdown) |
| **Corrected Amount** | operator | Optional. Positive number, max 2 decimals (0 for JPY). `1,234.50` is accepted |
| **Corrected Currency** | operator | Optional. `USD`, `EUR`, `GBP`, `JPY` (also `$ € £ ¥ ￥ 円`) |
| Status, Status Reason, Updated At | system | Lifecycle and the reason for the last change |
| Saved PDF URLs, Artifact File IDs | system | Drive files written or reused for this row |

### Status lifecycle

```
NEEDS_REVIEW --(APPROVE + Process)--> IN_PROGRESS --> DONE
      |                                   |   \--> LABEL_PENDING --(next Process)--> DONE
      |                                   \--> ERROR (transient: Decision stays APPROVE, retried next run)
      |                                   \--> ERROR + Decision=HOLD (conflict: manual reconciliation)
      +--(HOLD)--> HELD      +--(SKIP)--> SKIPPED
      +--(invalid / missing amount / changed source)--> Decision=HOLD, HELD (with reason)
```

### What APPROVE requires (checked again at processing time)

1. Exactly one Review row has this Message ID.
2. The Gmail message still exists and is not in Trash.
3. The message's current Revision equals the Revision you reviewed. A change (content,
   attachments, or a parser rule change via `SAFE_INTAKE.revisionVersion`) → HOLD.
4. The system columns were not hand-edited (e.g. typing into Detected Total) → HOLD.
5. An amount and a supported currency are known: either detected, or from valid
   corrections. Unknown total, total without currency, non-positive, malformed, too many
   decimals (JPY must be whole) → HOLD with the reason in `Status Reason`.

When a row is held by the system, its Decision is changed to HOLD. Fix the cause and
set APPROVE again; a held row never processes by itself. If an earlier attempt on the row
may have left PDFs (status was IN_PROGRESS / ERROR / LABEL_PENDING, or `Artifact File IDs`
is filled), the row shows `ERROR` instead of `HELD`, and HOLD / SKIP will not relabel it
as `HELD` / `SKIPPED` until you have reconciled it (see the runbook).

A message that is already in `Records` gets only the Gmail label on APPROVE. That label
is applied only if the row still matches the current message.

### Refresh rules

`refreshReviewQueue()` appends newly matching messages (blank Decision, status
`NEEDS_REVIEW`), keeps Decision and corrections for rows whose Revision and system
columns still match, and for any changed row rewrites the system columns, sets
Decision = HOLD and explains why. Corrections stay filled in, and the reason warns that
they will be used if you APPROVE again. Refresh stops taking new messages after
~4 minutes; any left over are reported as `deferred` and picked up by the next refresh. Rows that
are `DONE` are never reopened. Rows are never deleted, including messages that no
longer match the search window.

## Processing safety model

Google Sheets, Drive and Gmail are separate services with no shared transaction, so
**exactly-once processing across them is impossible to guarantee**. Safe Intake instead
makes every step re-runnable and fail-visible:

- **Serialization.** v0.2 functions run under `LockService.getScriptLock()`
  (`tryLock`, 10 s wait). If the lock is busy they return `LOCKED` and change nothing.
- **Bounded runs.** At most 25 APPROVE rows or ~4 minutes per run; the rest are reported
  as `deferred` and stay APPROVE for the next run.
- **Claim before side effects.** A row is marked `IN_PROGRESS` (and flushed) before the
  first Drive write, so an interrupted run is visible.
- **Records check first.** If `Records` already contains the Message ID (earlier run,
  interrupted run, or legacy path) no PDF or Records write happens; only the label is
  (re)applied.
- **Deterministic PDF names.** `<messageId>_SI<n>_<attachment name>`, created in a single
  call with the final name. Created file IDs are written to `Artifact File IDs`
  immediately. On a retry, an existing file is reused only when its size and SHA-256
  match the attachment.
- **Conflicts are not "fixed" automatically.** A same-named file with different content,
  more than one same-named file, or a legacy-named file (`<messageId>-<name>`) without a
  Records row → `ERROR`, Decision = HOLD, an `Errors` row, and nothing overwritten or
  created. Trashed files are ignored.
- **Read-back.** After appending to Records, the Message ID is read back; if Sheets
  altered it the row is set to ERROR/HOLD so a retry cannot silently duplicate.
- **Identity-checked writes.** Status writes to the Review row go through a lookup by Message ID
  and are checked again afterwards. If rows were sorted or inserted during the write, the
  run records an error, and the Audit report flags any row left `DONE` without a Records
  row (`DONE_WITHOUT_RECORD`).
- **Formula-safe text.** Email-controlled text starting with `= + - @` is written with a
  leading `'` in Review and in Safe Intake Records rows.

### Known limits (residual risk)

- The legacy `processInvoiceEmails()` bypasses review and the lock. Running it at the
  same time as Safe Intake can race.
- Operator edits in the Sheets UI are not covered by the script lock. Each row is
  re-read right before its side effects, which narrows but does not close that window.
  Sorting or inserting Review rows during a run can still put a status on the wrong row.
  That is detected and logged (see above), but has to be fixed by hand. Avoid editing
  the Review sheet while Process is running.
- Message IDs that Sheets had already turned into numbers in **legacy** Records rows
  (before v0.2 set column B to plain text) will not match. A message with no PDFs could
  then be recorded a second time; for one with PDFs, the legacy-file check stops it.
  Inspect legacy Records rows before switching.
- New Records rows rely on Sheets carrying the plain-text format of column B into
  appended rows. The read-back check catches a failure, but this has not been verified
  on a real spreadsheet.
- Drive name search can lag behind a just-created file. If a run dies after `createFile`
  but before the file ID is saved in the row, an immediate retry may create a second
  copy. The Audit report flags it as `DUPLICATE_ARTIFACT`.
- Files moved out of the folder are found only through `Artifact File IDs`, and only while their
  name still starts with `<messageId>_SI`.
- The lock only serializes executions of this Apps Script project.
- Anyone with edit access to the spreadsheet can approve. Sheets version history is the
  only approver audit trail.
- The parser takes the first matching amount (it may be a subtotal or tax), reads `¥`
  as JPY and `$` as USD. Japanese totals written as `5,500円` are not detected and need a correction.
- Legacy `Preview` / legacy `Records` writes are unchanged (no formula hardening there).
- Local tests use an in-memory mock of Apps Script services. Real-service behaviour
  (type coercion, quotas, search consistency) is only approximated.

## Synthetic walkthrough (future, voluntary test)

Use a **throwaway Google account** with synthetic data only. This change did not perform
these steps.

1. Create a new spreadsheet and Drive folder, then follow [SETUP.md](SETUP.md) for Script
   Properties. Prefer a script bound to the spreadsheet so the menu appears.
2. Run **Safe Intake: 1. Set up review sheets**. Expect the `Review` sheet with the
   headers above.
3. Send the account three synthetic emails:
   - `Invoice - Demo Vendor #1001`, body `Grand Total: USD 123.45`, with a harmless PDF.
   - `Your receipt from Demo Shop`, body without an amount, with a harmless PDF.
   - `【請求書】Invoice No. 2026-001 デモ商事`, body `ご請求金額: ¥12,000`, no attachment.
4. Run **2. Refresh review queue**. Expect 3 rows with status `NEEDS_REVIEW`, and no new
   Drive files, Records rows or labels.
5. Set #1001 to `APPROVE`. Set the receipt to `APPROVE` without a correction. Set the
   Japanese row to `SKIP`.
6. Run **3. Process APPROVE rows**. Expect:
   - #1001: one Records row (`Source` = `safe-intake-v0.2`), one Drive file
     `<id>_SI1_<name>.pdf`, the Gmail label, and status `DONE`.
   - Receipt: Decision changed to `HOLD`, reason asks for Corrected Amount; nothing
     written.
   - Japanese row: `SKIPPED`, nothing written.
7. Enter `Corrected Amount` = `45.60` and `Corrected Currency` = `USD` on the receipt, set
   `APPROVE`, and process again. Expect a Records row with `Amount Source` = `corrected`.
8. Run Process and Refresh again. Expect no new Records rows or Drive files.
9. Run **Audit**. Expect a single `OK` row.

### Permissions this would request

Detected automatically from the code (no `oauthScopes` set in `appsscript.json`):
Gmail via `https://mail.google.com/`. That is **full mailbox access** (read, send,
delete), because `GmailApp` cannot request less; this code only searches, reads and
labels. Also Google Sheets
(`.../auth/spreadsheets`), Google Drive (`.../auth/drive`, needed for
`getFolderById` / `getFileById` / `createFile`), and, for a bound script, the container
UI for the menu and toasts. Script Properties and LockService need no extra scope. No
triggers are installed by this code.

## Reconciliation runbook

Run **Audit** first. It writes one row per finding to the `Audit` sheet and does not
modify Gmail, Drive or Records.

| Finding / status | What happened | What to do |
|---|---|---|
| `ERROR`, Decision still APPROVE | Transient failure (quota, timeout, service error) | Fix the cause if known, then run Process again. The retry re-checks Records and reuses matching PDFs. |
| `ERROR`, "belongs to another message" | `Artifact File IDs` holds another message's file (e.g. after rows moved) | Clear that cell only. Do **not** trash the file; it belongs to the other message. Then APPROVE again. |
| `ERROR`, Decision = HOLD, reason mentions a file | Conflicting, duplicate or legacy-named Drive file | Open the folder. Keep at most one file whose content is the attachment, and trash the others. Then set APPROVE and run Process. Trashed files are ignored, so if all are trashed a fresh copy is created. |
| `ERROR`, "did not read back verbatim" | Sheets changed the Message ID in Records | Set Records column B to Plain text and retype the exact ID from the Review row. Do this before any retry. |
| `INTERRUPTED` / `IN_PROGRESS` | A run stopped mid-row (timeout, crash) | If the row is still APPROVE, run Process: it reconciles. If you no longer want it, trash the PDFs listed under `ORPHAN_ARTIFACT`, clear `Status`, `Saved PDF URLs` and `Artifact File IDs`, then set HOLD/SKIP. |
| `LABEL_PENDING` | Recorded, but the Gmail label failed | Run Process again. If the message was deleted, set Status to `DONE` by hand. |
| `DONE_WITHOUT_RECORD` | A Review row says DONE / LABEL_PENDING, but Records has no row | Usually a status write landed on the wrong row while the sheet was being sorted. Check `Errors`, clear that row's `Status`, then APPROVE it again if it should be processed. |
| `ERROR`, "Review rows moved while writing" | Rows were sorted or inserted during a run | Run Audit, fix any `DONE_WITHOUT_RECORD` row, then run Process again. |
| `ORPHAN_ARTIFACT` | A `_SI<n>_` PDF has no Records row | If its row is APPROVE, run Process (the PDF is reused). Otherwise trash the file. |
| `DUPLICATE_ARTIFACT` | Two live files share a deterministic name | Keep the one whose URL is in the Records row, and trash the other. |
| `DUPLICATE_RECORDS` / `DUPLICATE_REVIEW_ROWS` | Same Message ID twice | Delete the extra row, keeping the one that matches `Saved PDF URLs`. |

## Rollback

- **Undo one processed message:** delete its Records row, trash its `<id>_SI<n>_` file(s),
  and remove the Gmail label. Then, in Review, clear `Status`, `Saved PDF URLs` and
  `Artifact File IDs`, and set Decision to HOLD or SKIP.
- **Return to the pre-v0.2 script:** restore the previous `Code.gs` (for example
  `git show f194ed9:Code.gs`). The extra `Review` / `Audit` sheets and the Records K:M
  columns are ignored by the legacy code and can be deleted by hand. Records B:C plain
  text formatting is harmless.
- Neither the code nor this runbook deletes Drive files, Records rows or Gmail data on
  its own.

## Local verification

No Google account is needed. Requires Node.js 20+ (no npm packages):

```sh
node --test tests/*.test.js
```

The tests load the real `Code.gs` into a `node:vm` context with an in-memory mock of
GmailApp, SpreadsheetApp, DriveApp, PropertiesService, LockService and Utilities
([tests/harness/gas-env.js](tests/harness/gas-env.js)), using synthetic English and
Japanese fixtures ([tests/fixtures.js](tests/fixtures.js)).
