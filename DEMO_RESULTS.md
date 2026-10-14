# Demo Results

Date: 2026-09-30

Test data: synthetic only.

| Check | Result |
|---|---|
| Sheet initialization | PASS |
| Gmail message detection | PASS |
| Preview | PASS |
| Record write | PASS |
| PDF attachment save to Drive | PASS |
| Gmail processed label | PASS |
| Duplicate-safe second run | PASS |

No real company, customer, or financial data was used.

This test verifies the bounded portfolio workflow only; it is not a production-scale reliability or security certification.

## v0.2 Safe Intake — local automated verification

Date: 2026-10-14. Environment: local Node.js (`node:test`), real `Code.gs` executed in a
`node:vm` context against an in-memory Apps Script mock. Synthetic English and Japanese
fixtures only. **No Google account, Gmail, Sheets or Drive was used. Nothing was
deployed.**

Command: `node --test tests/*.test.js`. Result: 47 tests, 47 passed (7 legacy regression,
40 Safe Intake acceptance, including 10 added after independent review) on Node v20.18.0
and v24.21.0.

Covered: explicit approve, HOLD / SKIP / blank decisions, missing amount or currency,
malformed corrections, stale approval and refresh invalidation, refresh preserving
decisions, edited system columns, repeat runs, lock unavailable and nested concurrent
runs, failure after PDF creation before the Records append, ambiguous Drive failure,
conflicting / duplicate / legacy orphan artifacts, trashed artifacts, label retry, bad
config and sheet layouts, duplicate Review rows, missing / trashed messages, numeric
Message ID coercion (digits and e-notation), formula-like subjects, bounded runs
(row and time budget), the audit report (orphan, interrupted, duplicate, DONE without
Records, refusing to overwrite a foreign Audit sheet), Review rows moved during a write,
Records read-back failure, and legacy compatibility.

Real-account behaviour of v0.2 has **not** been verified. See the walkthrough in
[SAFE_INTAKE.md](SAFE_INTAKE.md) for a future voluntary synthetic test.
