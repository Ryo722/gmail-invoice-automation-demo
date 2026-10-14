/**
 * Gmail → Google Sheets → Google Drive automation demo.
 *
 * Portfolio goals:
 * - find Invoice / Receipt emails
 * - extract sender, date, subject, and a best-effort total amount
 * - save PDF attachments to Drive
 * - prevent duplicate processing using Gmail message ID
 * - keep an error log
 * - support preview/dry-run before writing production rows
 *
 * v0.2 Safe Intake (opt-in): a review queue in the `Review` sheet where an
 * operator sets APPROVE / HOLD / SKIP (plus optional amount/currency
 * corrections). Only valid APPROVE rows create PDFs, append Records or apply
 * the Gmail label. Entry points: setupSafeIntake(), refreshReviewQueue(),
 * processApprovedReviewRows(), auditSafeIntake(). The legacy
 * processInvoiceEmails() entry point is kept for compatibility and is NOT
 * review-gated.
 *
 * Required Script Properties:
 *   SHEET_ID
 *   FOLDER_ID
 *
 * Optional Script Properties:
 *   PROCESSED_LABEL (default: portfolio/processed)
 *   MAX_THREADS     (default: 50)
 */

const CONFIG = {
  recordsSheet: 'Records',
  previewSheet: 'Preview',
  errorSheet: 'Errors',
  defaultProcessedLabel: 'portfolio/processed',
  searchQuery: '(subject:Invoice OR subject:Receipt) newer_than:30d',
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Invoice Automation')
    .addItem('Safe Intake: 1. Set up review sheets', 'setupSafeIntake')
    .addItem('Safe Intake: 2. Refresh review queue', 'refreshReviewQueue')
    .addItem('Safe Intake: 3. Process APPROVE rows', 'processApprovedReviewRows')
    .addItem('Safe Intake: Audit / reconciliation report', 'auditSafeIntake')
    .addSeparator()
    .addItem('Legacy: Setup sheets', 'setupSheets')
    .addItem('Legacy: Preview matching emails', 'previewInvoiceEmails')
    .addItem('Legacy: Process emails (no review gate)', 'processInvoiceEmails')
    .addToUi();
}

function setupSheets() {
  ensureBaseSheets_(getSpreadsheet_());
}

function ensureBaseSheets_(ss) {
  ensureSheet_(ss, CONFIG.recordsSheet, [
    'Processed At',
    'Message ID',
    'Thread ID',
    'Sender Name',
    'Sender Email',
    'Email Date',
    'Subject',
    'Detected Total',
    'Currency',
    'Saved PDF URLs',
  ]);

  ensureSheet_(ss, CONFIG.previewSheet, [
    'Message ID',
    'Sender',
    'Email Date',
    'Subject',
    'Detected Total',
    'Currency',
    'PDF Count',
  ]);

  ensureSheet_(ss, CONFIG.errorSheet, [
    'Occurred At',
    'Message ID',
    'Subject',
    'Error',
  ]);
}

function previewInvoiceEmails() {
  setupSheets();

  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(CONFIG.previewSheet);
  clearDataRows_(sheet);

  const messages = findCandidateMessages_();
  const rows = messages.map(message => {
    const parsed = parseMessage_(message);
    const pdfCount = message.getAttachments({
      includeInlineImages: false,
      includeAttachments: true,
    }).filter(a => isPdf_(a)).length;

    return [
      message.getId(),
      message.getFrom(),
      message.getDate(),
      message.getSubject(),
      parsed.total ?? '',
      parsed.currency ?? '',
      pdfCount,
    ];
  });

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }

  return {matchedMessages: rows.length};
}

function processInvoiceEmails() {
  setupSheets();

  const ss = getSpreadsheet_();
  const recordsSheet = ss.getSheetByName(CONFIG.recordsSheet);
  const processedIds = loadProcessedMessageIds_(recordsSheet);
  const folder = getDriveFolder_();
  const processedLabel = getOrCreateLabel_(
    getScriptProperty_('PROCESSED_LABEL', CONFIG.defaultProcessedLabel)
  );

  let processed = 0;
  let skipped = 0;
  let failed = 0;

  const messages = findCandidateMessages_();

  messages.forEach(message => {
    const messageId = message.getId();

    if (processedIds.has(messageId)) {
      skipped += 1;
      return;
    }

    try {
      const parsed = parseMessage_(message);
      const sender = parseSender_(message.getFrom());
      const pdfUrls = savePdfAttachments_(message, folder);

      recordsSheet.appendRow([
        new Date(),
        messageId,
        message.getThread().getId(),
        sender.name,
        sender.email,
        message.getDate(),
        message.getSubject(),
        parsed.total ?? '',
        parsed.currency ?? '',
        pdfUrls.join('\n'),
      ]);

      processedIds.add(messageId);
      message.getThread().addLabel(processedLabel);
      processed += 1;
    } catch (err) {
      failed += 1;
      logError_(ss, message, err);
    }
  });

  return {
    matchedMessages: messages.length,
    processed,
    skipped,
    failed,
  };
}

function findCandidateMessages_(validatedMaxThreads) {
  const maxThreads = validatedMaxThreads || Number(getScriptProperty_('MAX_THREADS', '50'));
  const threads = GmailApp.search(CONFIG.searchQuery, 0, maxThreads);
  const out = [];

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      if (subjectMatches_(message.getSubject())) {
        out.push(message);
      }
    });
  });

  return out;
}

function subjectMatches_(subject) {
  return /\b(invoice|receipt)\b/i.test(subject || '');
}

function parseMessage_(message) {
  const text = [
    message.getPlainBody() || '',
    message.getSubject() || '',
  ].join('\n');

  return parseTotal_(text);
}

function parseTotal_(text) {
  const normalized = String(text || '').replace(/\u00A0/g, ' ');

  const patterns = [
    {
      regex: /(?:grand\s+total|total\s+due|amount\s+due|total)\s*[:\-]?\s*(USD|EUR|GBP|JPY|\$|€|£|¥)?\s*([0-9][0-9,]*(?:\.[0-9]{1,2})?)/i,
      currencyGroup: 1,
      amountGroup: 2,
    },
    {
      regex: /(USD|EUR|GBP|JPY|\$|€|£|¥)\s*([0-9][0-9,]*(?:\.[0-9]{1,2})?)/i,
      currencyGroup: 1,
      amountGroup: 2,
    },
  ];

  for (const p of patterns) {
    const match = normalized.match(p.regex);
    if (!match) continue;

    const amount = Number(match[p.amountGroup].replace(/,/g, ''));
    if (!Number.isFinite(amount)) continue;

    return {
      total: amount,
      currency: normalizeCurrency_(match[p.currencyGroup] || ''),
    };
  }

  return {total: null, currency: null};
}

function normalizeCurrency_(value) {
  const v = String(value || '').toUpperCase();
  const map = {
    '$': 'USD',
    '€': 'EUR',
    '£': 'GBP',
    '¥': 'JPY',
  };
  return map[v] || v || '';
}

function savePdfAttachments_(message, folder) {
  const attachments = message.getAttachments({
    includeInlineImages: false,
    includeAttachments: true,
  });

  return attachments
    .filter(a => isPdf_(a))
    .map((attachment, index) => {
      const original = sanitizeFilename_(attachment.getName() || `attachment-${index + 1}.pdf`);
      const name = `${message.getId()}-${original}`;
      const file = folder.createFile(attachment.copyBlob()).setName(name);
      return file.getUrl();
    });
}

function isPdf_(attachment) {
  const contentType = String(attachment.getContentType() || '').toLowerCase();
  const name = String(attachment.getName() || '').toLowerCase();
  return contentType === 'application/pdf' || name.endsWith('.pdf');
}

function parseSender_(fromValue) {
  const raw = String(fromValue || '').trim();
  const match = raw.match(/^(.*?)\s*<([^>]+)>$/);

  if (match) {
    return {
      name: match[1].replace(/^"|"$/g, '').trim(),
      email: match[2].trim(),
    };
  }

  return {
    name: '',
    email: raw,
  };
}

function loadProcessedMessageIds_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set();

  const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  return new Set(values.flat().filter(Boolean).map(String));
}

function logError_(ss, message, err) {
  const sheet = ss.getSheetByName(CONFIG.errorSheet);
  sheet.appendRow([
    new Date(),
    message ? message.getId() : '',
    message ? message.getSubject() : '',
    err && err.stack ? err.stack : String(err),
  ]);
}

function getSpreadsheet_() {
  const id = getRequiredScriptProperty_('SHEET_ID');
  return SpreadsheetApp.openById(id);
}

function getDriveFolder_() {
  const id = getRequiredScriptProperty_('FOLDER_ID');
  return DriveApp.getFolderById(id);
}

function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

function getRequiredScriptProperty_(name) {
  const value = PropertiesService.getScriptProperties().getProperty(name);
  if (!value) {
    throw new Error(`Missing required Script Property: ${name}`);
  }
  return value;
}

function getScriptProperty_(name, fallback) {
  return PropertiesService.getScriptProperties().getProperty(name) || fallback;
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function clearDataRows_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getMaxColumns()).clearContent();
  }
}

function sanitizeFilename_(name) {
  return String(name)
    .replace(/[\\/:*?"<>|]+/g, '_')
    .slice(0, 180);
}

// ===========================================================================
// v0.2 Safe Intake — review-gated, retry-safe processing (opt-in)
//
// Flow: setupSafeIntake() -> refreshReviewQueue() -> operator edits Decision /
// corrections in the Review sheet -> processApprovedReviewRows().
//
// Google Sheets, Drive and Gmail are separate non-transactional services, so
// exactly-once processing cannot be guaranteed. Instead, every side effect is
// (a) gated on an explicit, re-validated APPROVE, (b) serialized with the
// script lock, (c) made re-runnable: Records is checked by Message ID, PDFs use
// deterministic names and are verified by SHA-256 before reuse, and anything
// ambiguous is surfaced as HOLD / ERROR instead of being overwritten or
// duplicated.
// ===========================================================================

const SAFE_INTAKE = {
  source: 'safe-intake-v0.2',
  reviewSheet: 'Review',
  auditSheet: 'Audit',
  // Bump when anything feeding buildMessageSnapshot_() (e.g. parseTotal_)
  // changes meaning, so approvals given against the old values are invalidated.
  revisionVersion: 'r1',
  decisions: ['APPROVE', 'HOLD', 'SKIP'],
  currencies: ['USD', 'EUR', 'GBP', 'JPY'],
  zeroDecimalCurrencies: ['JPY'],
  lockWaitMs: 10 * 1000,
  maxRuntimeMs: 4 * 60 * 1000,
  maxRowsPerRun: 25,
  maxThreadsLimit: 500,
  recordsExtraHeaders: ['Source', 'Amount Source', 'Review Revision'],
};

const REVIEW_HEADERS = [
  'Message ID',
  'Thread ID',
  'Sender',
  'Email Date',
  'Subject',
  'Detected Total',
  'Detected Currency',
  'PDF Count',
  'Revision',
  'Decision',
  'Corrected Amount',
  'Corrected Currency',
  'Status',
  'Status Reason',
  'Updated At',
  'Saved PDF URLs',
  'Artifact File IDs',
];

const REVIEW_COL = {
  messageId: 1,
  threadId: 2,
  sender: 3,
  emailDate: 4,
  subject: 5,
  detectedTotal: 6,
  detectedCurrency: 7,
  pdfCount: 8,
  revision: 9,
  decision: 10,
  correctedAmount: 11,
  correctedCurrency: 12,
  status: 13,
  reason: 14,
  updatedAt: 15,
  pdfUrls: 16,
  artifactIds: 17,
};

// Columns kept as plain text so Sheets cannot turn IDs into numbers.
const REVIEW_TEXT_COLUMNS = [1, 2, 3, 5, 7, 9, 10, 12, 13, 14, 16, 17];

const REVIEW_STATUS = {
  NEEDS_REVIEW: 'NEEDS_REVIEW',
  HELD: 'HELD',
  SKIPPED: 'SKIPPED',
  IN_PROGRESS: 'IN_PROGRESS',
  DONE: 'DONE',
  LABEL_PENDING: 'LABEL_PENDING',
  ERROR: 'ERROR',
};

// ---------- Public v0.2 entry points ----------

function setupSafeIntake() {
  const config = loadSafeIntakeConfig_({needsFolder: false});
  const result = withScriptLock_(() => {
    setupSafeIntakeSheets_(config.ss);
    return {status: 'OK'};
  });
  report_(config.ss, 'Safe Intake setup', result);
  return result;
}

/**
 * Lists candidate messages in the Review sheet. Writes only to the Review
 * sheet: no Drive, Records or Gmail label changes.
 */
function refreshReviewQueue() {
  const config = loadSafeIntakeConfig_({needsFolder: false});
  const result = withScriptLock_(() => {
    const sheet = setupSafeIntakeSheets_(config.ss).review;
    const rowsById = groupReviewRowsById_(readReviewRows_(sheet));
    const counts = {matched: 0, added: 0, unchanged: 0, invalidated: 0, ambiguous: 0, deferred: 0};
    const seen = new Set();
    const newRows = [];
    const startedAt = Date.now();

    findCandidateMessages_(config.maxThreads).forEach(message => {
      const id = String(message.getId());
      if (seen.has(id)) return;
      seen.add(id);
      counts.matched += 1;
      if (Date.now() - startedAt > SAFE_INTAKE.maxRuntimeMs) {
        // Stay inside the Apps Script time limit; the next refresh continues.
        counts.deferred += 1;
        return;
      }

      const snap = buildMessageSnapshot_(message);
      const existing = rowsById.get(id) || [];

      if (existing.length === 0) {
        newRows.push(newReviewRow_(snap));
        counts.added += 1;
        return;
      }
      if (existing.length > 1) {
        markDuplicateReviewRows_(sheet, id, existing);
        counts.ambiguous += 1;
        return;
      }

      const row = existing[0];
      const mismatch = describeReviewMismatch_(row, snap);
      if (!mismatch) {
        counts.unchanged += 1;
        return;
      }

      if (row.status === REVIEW_STATUS.DONE) {
        // Already processed: never reopen automatically.
        updateReviewRow_(sheet, id, {
          reason: `Processed earlier; the source now differs (${mismatch}). Records were not modified.`,
        });
        counts.unchanged += 1;
        return;
      }

      const mayHaveArtifacts = rowMayHaveArtifacts_(row);
      const hasCorrections = row.correctedAmount !== '' || row.correctedCurrency !== '';
      updateReviewRow_(sheet, id, Object.assign(systemReviewFields_(snap), {
        decision: 'HOLD',
        status: mayHaveArtifacts ? REVIEW_STATUS.ERROR : REVIEW_STATUS.HELD,
        reason: `Re-review required: ${mismatch}. Check the row, then set Decision again.` +
          (hasCorrections ? ' The earlier Corrected Amount / Currency are still filled in and will be used on APPROVE; clear them if they no longer apply.' : '') +
          (mayHaveArtifacts ? ' ' + ARTIFACT_WARNING : ''),
      }));
      counts.invalidated += 1;
    });

    appendReviewRows_(sheet, newRows);
    return Object.assign({status: 'OK'}, counts);
  });
  report_(config.ss, 'Refresh review queue', result);
  return result;
}

/**
 * Processes only Review rows whose Decision is a valid APPROVE and whose source
 * still matches what was reviewed.
 */
function processApprovedReviewRows() {
  const config = loadSafeIntakeConfig_({needsFolder: true});
  const result = withScriptLock_(() => runSafeIntake_(config));
  report_(config.ss, 'Process APPROVE rows', result);
  return result;
}

/**
 * Read-only consistency report (Gmail / Drive / Records are not modified);
 * findings are written to the Audit sheet.
 */
function auditSafeIntake() {
  const config = loadSafeIntakeConfig_({needsFolder: true});
  const result = withScriptLock_(() => {
    assertAuditSheetOwned_(config.ss);
    const sheets = setupSafeIntakeSheets_(config.ss);
    const findings = collectAuditFindings_(sheets, config.folder);
    writeAuditSheet_(config.ss, findings);
    return {status: 'OK', findings};
  });
  report_(config.ss, 'Safe Intake audit',
    result.findings ? {status: result.status, findings: result.findings.length} : result);
  return result;
}

// ---------- Processing ----------

function runSafeIntake_(config) {
  const startedAt = Date.now();
  const sheets = setupSafeIntakeSheets_(config.ss);
  const ctx = {
    config,
    ss: config.ss,
    review: sheets.review,
    records: sheets.records,
    folder: config.folder,
    recordedIds: loadProcessedMessageIds_(sheets.records),
    label: null,
  };
  const result = {
    status: 'OK',
    processed: 0,
    alreadyRecorded: 0,
    labelPending: 0,
    held: 0,
    skipped: 0,
    errors: 0,
    deferred: 0,
  };

  const rows = readReviewRows_(ctx.review);
  const groups = groupReviewRowsById_(rows);
  const duplicatesHandled = new Set();
  let attempted = 0;

  rows.forEach(row => {
    const id = row.messageId;
    const group = groups.get(id);

    if (group.length > 1) {
      if (!duplicatesHandled.has(id)) {
        duplicatesHandled.add(id);
        if (group.some(r => r.status !== REVIEW_STATUS.DONE)) {
          markDuplicateReviewRows_(ctx.review, id, group);
          result.errors += 1;
        }
      }
      return;
    }

    if (row.status === REVIEW_STATUS.DONE) return;

    const decision = normalizeDecision_(row.decisionRaw);
    if (decision === '') return;
    if (decision !== 'APPROVE') {
      result[settleNonApprovedRow_(ctx, row, decision)] += 1;
      return;
    }

    if (attempted >= SAFE_INTAKE.maxRowsPerRun ||
        Date.now() - startedAt > SAFE_INTAKE.maxRuntimeMs) {
      result.deferred += 1;
      return;
    }
    attempted += 1;
    result[processApprovedRowSafely_(ctx, id)] += 1;
  });

  return result;
}

function settleNonApprovedRow_(ctx, row, decision) {
  const id = row.messageId;
  if (!SAFE_INTAKE.decisions.includes(decision)) {
    return holdRow_(ctx, row,
      `Decision "${String(row.decisionRaw).trim()}" is not one of APPROVE / HOLD / SKIP.`);
  }

  const outcome = decision === 'HOLD' ? 'held' : 'skipped';

  if (row.status === REVIEW_STATUS.IN_PROGRESS || row.artifactIds.trim() !== '') {
    // Do not silently "settle" a row whose earlier attempt was interrupted.
    updateReviewRow_(ctx.review, id, {
      reason: `An earlier attempt was interrupted or failed while processing this row and Decision is now ${decision}. ` +
        'PDFs may exist without a Records row: run the Audit report and reconcile manually. Nothing was changed.',
    });
    return outcome;
  }

  const settleable = ['', REVIEW_STATUS.NEEDS_REVIEW, REVIEW_STATUS.HELD, REVIEW_STATUS.SKIPPED];
  const target = decision === 'HOLD' ? REVIEW_STATUS.HELD : REVIEW_STATUS.SKIPPED;
  if (settleable.includes(row.status) && row.status !== target) {
    updateReviewRow_(ctx.review, id, {
      status: target,
      reason: decision === 'HOLD'
        ? 'Held by reviewer.'
        : 'Skipped by reviewer: no PDF, Records or label writes.',
    });
  }
  return outcome;
}

function processApprovedRowSafely_(ctx, id) {
  try {
    return processApprovedRow_(ctx, id);
  } catch (err) {
    logErrorById_(ctx.ss, id, '', err);
    try {
      updateReviewRow_(ctx.review, id, {
        status: REVIEW_STATUS.ERROR,
        reason: `Attempt failed: ${errorMessage_(err)}. Decision stays APPROVE: the next run re-checks ` +
          'Records and Drive before writing, so a retry does not duplicate. If it keeps failing, set HOLD and see the runbook.',
      });
    } catch (writeErr) {
      logErrorById_(ctx.ss, id, '', writeErr);
    }
    return 'errors';
  }
}

function processApprovedRow_(ctx, id) {
  // Re-read the row now: the operator may have edited it since the run began.
  const fresh = readReviewRowsById_(ctx.review, id);
  if (fresh.length !== 1) {
    throw new Error(`Review row for ${id} was moved, deleted or duplicated during the run`);
  }
  const row = fresh[0];
  if (normalizeDecision_(row.decisionRaw) !== 'APPROVE' || row.status === REVIEW_STATUS.DONE) {
    return 'skipped';
  }

  if (ctx.recordedIds.has(id)) {
    // Recorded earlier (previous run, interrupted run, or the legacy path).
    // Only the label may still be missing; apply it only for a still-valid approval.
    const recordedMessage = getMessageOrNull_(id);
    const recordedMismatch = recordedMessage
      ? describeReviewMismatch_(row, buildMessageSnapshot_(recordedMessage))
      : '';
    if (recordedMismatch) {
      return holdRow_(ctx, row, `Already in Records, but ${recordedMismatch}; the Gmail label was not applied. ` +
        'Refresh, re-check and APPROVE again if the label is still wanted.');
    }
    const labelError = applyProcessedLabel_(ctx, recordedMessage, id);
    updateReviewRow_(ctx.review, id, labelError
      ? {
        status: REVIEW_STATUS.LABEL_PENDING,
        reason: `Already in Records; Gmail label failed (${labelError}). The next run retries the label only.`,
      }
      : {
        status: REVIEW_STATUS.DONE,
        reason: 'Already recorded in Records (earlier run or legacy path); no new PDF or Records writes.',
      });
    return labelError ? 'labelPending' : 'alreadyRecorded';
  }

  const message = getMessageOrNull_(id);
  if (!message) {
    return holdRow_(ctx, row, 'Gmail message not found or not accessible; nothing was written.');
  }
  if (message.isInTrash()) {
    return holdRow_(ctx, row, 'Gmail message is in Trash; nothing was written.');
  }

  const snap = buildMessageSnapshot_(message);
  const mismatch = describeReviewMismatch_(row, snap);
  if (mismatch) {
    return holdRow_(ctx, row,
      `Not processed: ${mismatch}. Run "Refresh review queue", re-check the row and APPROVE again.`);
  }

  const amount = resolveEffectiveAmount_(row, snap);
  if (!amount.ok) return holdRow_(ctx, row, amount.reason);

  // Claim the row before the first side effect so an interruption is visible.
  updateReviewRow_(ctx.review, id, {
    status: REVIEW_STATUS.IN_PROGRESS,
    reason: `Processing started ${new Date().toISOString()}`,
  });
  SpreadsheetApp.flush();

  const artifacts = reconcilePdfArtifacts_(ctx, row, snap);
  if (!artifacts.ok) {
    logErrorById_(ctx.ss, id, snap.subject, new Error(artifacts.reason));
    updateReviewRow_(ctx.review, id, {
      decision: 'HOLD',
      status: REVIEW_STATUS.ERROR,
      reason: `${artifacts.reason} Nothing was overwritten or duplicated. Reconcile manually (see runbook), ` +
        'then clear Status and APPROVE again.',
    });
    return 'errors';
  }

  const urls = artifacts.files.map(f => f.getUrl());
  const sender = parseSender_(snap.from);
  ctx.records.appendRow([
    new Date(),
    snap.id,
    snap.threadId,
    sheetText_(sender.name),
    sheetText_(sender.email),
    snap.date,
    sheetText_(snap.subject),
    amount.total,
    amount.currency,
    urls.join('\n'),
    SAFE_INTAKE.source,
    amount.source,
    snap.revision,
  ]);
  SpreadsheetApp.flush();

  ctx.recordedIds = loadProcessedMessageIds_(ctx.records);
  if (!ctx.recordedIds.has(id)) {
    const reason = 'A Records row was appended but its Message ID did not read back verbatim ' +
      '(Sheets may have converted it). Fix that Records row (column B, plain text) before any retry, ' +
      'otherwise a retry would add a second row.';
    logErrorById_(ctx.ss, id, snap.subject, new Error(reason));
    updateReviewRow_(ctx.review, id, {decision: 'HOLD', status: REVIEW_STATUS.ERROR, reason, pdfUrls: urls.join('\n')});
    return 'errors';
  }

  const labelError = applyProcessedLabel_(ctx, message, id);
  const summary = `${urls.length} PDF(s) (created ${artifacts.created}, reused ${artifacts.reused}); ` +
    `amount ${amount.total} ${amount.currency} (${amount.source}).`;
  updateReviewRow_(ctx.review, id, {
    status: labelError ? REVIEW_STATUS.LABEL_PENDING : REVIEW_STATUS.DONE,
    reason: labelError
      ? `Recorded: ${summary} Gmail label failed (${labelError}); the next run retries the label only.`
      : `Processed: ${summary}`,
    pdfUrls: urls.join('\n'),
  });
  return labelError ? 'labelPending' : 'processed';
}

/**
 * Finds or creates one Drive file per PDF attachment using deterministic names.
 * Existing files are reused only if size and SHA-256 match; anything else is a
 * conflict that the caller surfaces without writing.
 */
function reconcilePdfArtifacts_(ctx, row, snap) {
  const ids = String(row.artifactIds || '').split('\n').map(s => s.trim());
  const files = [];
  let created = 0;
  let reused = 0;

  for (const pdf of snap.pdfs) {
    const name = safeIntakeArtifactName_(snap.id, pdf);
    const legacyName = legacyArtifactName_(snap.id, pdf);

    const legacyHits = liveFilesByName_(ctx.folder, legacyName);
    if (legacyHits.length) {
      return {
        ok: false,
        reason: `Found ${legacyHits.length} legacy-named file(s) "${legacyName}" but no Records row ` +
          '(interrupted legacy run?).',
      };
    }

    const candidates = liveFilesByName_(ctx.folder, name);
    const recordedId = ids[pdf.index];
    if (recordedId && !candidates.some(f => f.getId() === recordedId)) {
      const recorded = getFileOrNull_(recordedId);
      if (recorded && !recorded.isTrashed()) {
        if (String(recorded.getName()).indexOf(`${snap.id}_SI`) !== 0) {
          return {
            ok: false,
            reason: `Artifact File IDs on this row points to "${recorded.getName()}", which belongs to another ` +
              'message: clear that cell (do NOT trash the file).',
          };
        }
        candidates.push(recorded);
      }
    }

    let file;
    if (candidates.length > 1) {
      return {ok: false, reason: `Found ${candidates.length} files for "${name}"; cannot tell which one is authoritative.`};
    } else if (candidates.length === 1) {
      if (!fileMatchesPdf_(candidates[0], pdf)) {
        return {ok: false, reason: `Existing Drive file "${candidates[0].getName()}" does not match the attachment (size / SHA-256).`};
      }
      file = candidates[0];
      reused += 1;
    } else {
      // Name is set on the blob so the file is created with its deterministic
      // name in a single call (no unnamed orphan if a later call fails).
      file = ctx.folder.createFile(pdf.attachment.copyBlob().setName(name));
      created += 1;
    }

    files.push(file);
    if (ids[pdf.index] !== file.getId()) {
      ids[pdf.index] = file.getId();
      const recorded = snap.pdfs.map(p => ids[p.index] || '').join('\n');
      updateReviewRow_(ctx.review, snap.id, {artifactIds: recorded});
      SpreadsheetApp.flush();
    }
  }

  return {ok: true, files, created, reused};
}

function safeIntakeArtifactName_(messageId, pdf) {
  return `${messageId}_SI${pdf.index + 1}_${pdf.safeName}`;
}

function legacyArtifactName_(messageId, pdf) {
  // Mirrors savePdfAttachments_() naming.
  return `${messageId}-${pdf.safeName}`;
}

function liveFilesByName_(folder, name) {
  const out = [];
  const it = folder.getFilesByName(name);
  while (it.hasNext()) {
    const f = it.next();
    if (!f.isTrashed()) out.push(f);
  }
  return out;
}

function fileMatchesPdf_(file, pdf) {
  if (file.getSize() !== pdf.size) return false;
  return sha256Hex_(file.getBlob().getBytes()) === pdf.digest;
}

function getFileOrNull_(id) {
  try {
    return DriveApp.getFileById(id);
  } catch (err) {
    return null;
  }
}

function getMessageOrNull_(id) {
  try {
    return GmailApp.getMessageById(id) || null;
  } catch (err) {
    return null;
  }
}

function applyProcessedLabel_(ctx, message, id) {
  try {
    const target = message || GmailApp.getMessageById(id);
    if (!ctx.label) ctx.label = getOrCreateLabel_(ctx.config.labelName);
    target.getThread().addLabel(ctx.label);
    return '';
  } catch (err) {
    logErrorById_(ctx.ss, id, '', err);
    return errorMessage_(err);
  }
}

const ARTIFACT_WARNING = 'An earlier attempt may have left PDFs; run the Audit report before changing this row.';

/** Rows touched by an earlier attempt keep a visible ERROR instead of HELD. */
function rowMayHaveArtifacts_(row) {
  return [REVIEW_STATUS.IN_PROGRESS, REVIEW_STATUS.ERROR, REVIEW_STATUS.LABEL_PENDING].includes(row.status) ||
    String(row.artifactIds || '').trim() !== '';
}

function holdRow_(ctx, row, reason) {
  const mayHaveArtifacts = rowMayHaveArtifacts_(row);
  updateReviewRow_(ctx.review, row.messageId, {
    decision: 'HOLD',
    status: mayHaveArtifacts ? REVIEW_STATUS.ERROR : REVIEW_STATUS.HELD,
    reason: mayHaveArtifacts ? `${reason} ${ARTIFACT_WARNING}` : reason,
  });
  return 'held';
}

// ---------- Snapshot / validation ----------

function buildMessageSnapshot_(message) {
  const id = String(message.getId());
  const threadId = String(message.getThread().getId());
  const from = String(message.getFrom() || '');
  const date = message.getDate();
  const subject = String(message.getSubject() || '');
  const body = String(message.getPlainBody() || '');
  const parsed = parseMessage_(message);

  const pdfs = message.getAttachments({
    includeInlineImages: false,
    includeAttachments: true,
  }).filter(a => isPdf_(a)).map((attachment, index) => {
    const bytes = attachment.copyBlob().getBytes();
    return {
      index,
      attachment,
      name: String(attachment.getName() || ''),
      safeName: sanitizeFilename_(attachment.getName() || `attachment-${index + 1}.pdf`),
      contentType: String(attachment.getContentType() || ''),
      size: bytes.length,
      digest: sha256Hex_(bytes),
    };
  });

  const revisionInput = JSON.stringify({
    v: SAFE_INTAKE.revisionVersion,
    id,
    threadId,
    from,
    date: date ? new Date(date).toISOString() : '',
    subject,
    body: sha256Hex_(body),
    total: parsed.total,
    currency: parsed.currency,
    pdfs: pdfs.map(p => [p.name, p.contentType, p.size, p.digest]),
  });

  return {
    id,
    threadId,
    from,
    date,
    subject,
    total: parsed.total,
    currency: parsed.currency || '',
    pdfs,
    revision: sha256Hex_(revisionInput).slice(0, 32),
  };
}

function describeReviewMismatch_(row, snap) {
  if (String(row.revision).trim() !== snap.revision) {
    return 'source content changed since review (revision mismatch)';
  }
  const edited = [];
  if (String(row.threadId).trim() !== snap.threadId) edited.push('Thread ID');
  if (!sameCellNumber_(row.detectedTotal, snap.total)) edited.push('Detected Total');
  if (String(row.detectedCurrency).trim() !== snap.currency) edited.push('Detected Currency');
  if (!sameCellNumber_(row.pdfCount, snap.pdfs.length)) edited.push('PDF Count');
  if (edited.length) {
    return `system column(s) ${edited.join(', ')} were edited in the Review sheet ` +
      '(put changes in Corrected Amount / Corrected Currency instead)';
  }
  return '';
}

function sameCellNumber_(cell, value) {
  if (value === null || value === undefined) return cell === '' || cell === null;
  return typeof cell === 'number' && cell === value;
}

function normalizeDecision_(raw) {
  return String(raw === null || raw === undefined ? '' : raw).trim().toUpperCase();
}

/**
 * Returns the amount/currency to record, preferring validated corrections.
 * Detected values are best-effort and never accepted without a currency.
 */
function resolveEffectiveAmount_(row, snap) {
  const fail = reason => ({ok: false, reason});

  const corrected = parseCorrectedAmount_(row.correctedAmount);
  if (corrected.error) return fail(corrected.error);
  const correctedCurrency = parseCorrectedCurrency_(row.correctedCurrency);
  if (correctedCurrency.error) return fail(correctedCurrency.error);

  const total = corrected.present ? corrected.value : snap.total;
  const currency = correctedCurrency.present ? correctedCurrency.value : snap.currency;

  if (total === null || total === undefined) {
    return fail('No total was detected and no Corrected Amount was given: enter Corrected Amount and ' +
      'Corrected Currency, then APPROVE again.');
  }
  if (!currency) {
    return fail('A total was detected without a currency: enter Corrected Currency ' +
      `(${SAFE_INTAKE.currencies.join(' / ')}), then APPROVE again.`);
  }
  if (!SAFE_INTAKE.currencies.includes(currency)) {
    return fail(`Currency "${currency}" is not supported; enter Corrected Currency ` +
      `(${SAFE_INTAKE.currencies.join(' / ')}).`);
  }
  if (!(total > 0)) {
    return fail(`Amount ${total} is not positive; enter a Corrected Amount.`);
  }
  if (!hasAllowedDecimals_(total, currency)) {
    const places = SAFE_INTAKE.zeroDecimalCurrencies.includes(currency) ? 0 : 2;
    return fail(`Amount ${total} has more than ${places} decimal place(s) for ${currency}; ` +
      'enter a Corrected Amount.');
  }

  return {
    ok: true,
    total,
    currency,
    source: corrected.present || correctedCurrency.present ? 'corrected' : 'detected',
  };
}

function parseCorrectedAmount_(value) {
  if (value === '' || value === null || value === undefined) return {present: false};
  let n;
  if (typeof value === 'number') {
    n = value;
  } else if (typeof value === 'string') {
    const s = value.trim();
    if (s === '') return {present: false};
    if (!/^(\d{1,3}(,\d{3})+|\d+)(\.\d+)?$/.test(s)) {
      return {error: `Corrected Amount "${s}" is not a plain positive number (e.g. 1234.50).`};
    }
    n = Number(s.replace(/,/g, ''));
  } else {
    return {error: 'Corrected Amount must be a number (the cell holds a date or another type).'};
  }
  if (!Number.isFinite(n) || n <= 0) {
    return {error: `Corrected Amount ${n} must be a number greater than zero.`};
  }
  return {present: true, value: n};
}

function parseCorrectedCurrency_(value) {
  if (value === '' || value === null || value === undefined) return {present: false};
  if (typeof value !== 'string') {
    return {error: 'Corrected Currency must be a currency code such as USD or JPY.'};
  }
  const s = value.trim();
  if (s === '') return {present: false};
  const aliases = {'$': 'USD', 'US$': 'USD', '€': 'EUR', '£': 'GBP', '¥': 'JPY', '￥': 'JPY', '円': 'JPY'};
  const code = aliases[s] || s.toUpperCase();
  if (!SAFE_INTAKE.currencies.includes(code)) {
    return {error: `Corrected Currency "${s}" is not supported (${SAFE_INTAKE.currencies.join(' / ')}).`};
  }
  return {present: true, value: code};
}

function hasAllowedDecimals_(amount, currency) {
  const places = SAFE_INTAKE.zeroDecimalCurrencies.includes(currency) ? 0 : 2;
  const scaled = amount * Math.pow(10, places);
  return Math.abs(scaled - Math.round(scaled)) < 1e-6;
}

// ---------- Review sheet I/O ----------

function setupSafeIntakeSheets_(ss) {
  ensureBaseSheets_(ss);

  const records = ss.getSheetByName(CONFIG.recordsSheet);
  assertRecordsLayout_(records);

  let review = ss.getSheetByName(SAFE_INTAKE.reviewSheet);
  if (!review) review = ss.insertSheet(SAFE_INTAKE.reviewSheet);
  ensureColumns_(review, REVIEW_HEADERS.length);
  if (review.getLastRow() === 0) {
    review.getRange(1, 1, 1, REVIEW_HEADERS.length).setValues([REVIEW_HEADERS]);
    review.setFrozenRows(1);
    if (review.getMaxRows() >= 2) {
      review.getRange(2, REVIEW_COL.decision, review.getMaxRows() - 1, 1)
        .setDataValidation(decisionValidationRule_());
    }
  } else {
    const header = review.getRange(1, 1, 1, REVIEW_HEADERS.length).getValues()[0].map(String);
    if (header.join('\u0001') !== REVIEW_HEADERS.join('\u0001')) {
      throw new Error(`Safe Intake: the "${SAFE_INTAKE.reviewSheet}" sheet has unexpected headers; ` +
        'rename or remove it and run setupSafeIntake() again. Nothing was processed.');
    }
  }
  return {records, review};
}

function assertRecordsLayout_(sheet) {
  const width = 10 + SAFE_INTAKE.recordsExtraHeaders.length;
  ensureColumns_(sheet, width);
  const header = sheet.getRange(1, 1, 1, width).getValues()[0].map(String);
  if (header[1] !== 'Message ID') {
    throw new Error(`Safe Intake: Records column B header must be "Message ID" (found "${header[1]}"); ` +
      'refusing to guess where processed IDs are stored.');
  }
  const extras = header.slice(10);
  if (extras.every(v => v === '')) {
    sheet.getRange(1, 11, 1, extras.length).setValues([SAFE_INTAKE.recordsExtraHeaders]);
  } else if (extras.join('\u0001') !== SAFE_INTAKE.recordsExtraHeaders.join('\u0001')) {
    throw new Error('Safe Intake: Records columns K:M hold unexpected headers ' +
      `(${extras.join(', ')}); expected ${SAFE_INTAKE.recordsExtraHeaders.join(', ')}.`);
  }
  // Keep Message ID / Thread ID as plain text so Sheets cannot coerce them.
  if (sheet.getMaxRows() >= 2) {
    sheet.getRange(2, 2, sheet.getMaxRows() - 1, 2).setNumberFormat('@');
  }
}

function ensureColumns_(sheet, count) {
  const max = sheet.getMaxColumns();
  if (max < count) sheet.insertColumnsAfter(max, count - max);
}

function ensureRows_(sheet, lastRowNeeded) {
  const max = sheet.getMaxRows();
  if (lastRowNeeded > max) sheet.insertRowsAfter(max, lastRowNeeded - max);
}

function decisionValidationRule_() {
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(SAFE_INTAKE.decisions, true)
    .setAllowInvalid(false)
    .setHelpText('APPROVE = write PDF/Records/label, HOLD = keep for review, SKIP = never process.')
    .build();
}

function newReviewRow_(snap) {
  return [
    snap.id,
    snap.threadId,
    sheetText_(snap.from),
    snap.date,
    sheetText_(snap.subject),
    snap.total === null || snap.total === undefined ? '' : snap.total,
    snap.currency,
    snap.pdfs.length,
    snap.revision,
    '',
    '',
    '',
    REVIEW_STATUS.NEEDS_REVIEW,
    initialReviewHint_(snap),
    new Date(),
    '',
    '',
  ];
}

function systemReviewFields_(snap) {
  return {
    threadId: snap.threadId,
    sender: sheetText_(snap.from),
    emailDate: snap.date,
    subject: sheetText_(snap.subject),
    detectedTotal: snap.total === null || snap.total === undefined ? '' : snap.total,
    detectedCurrency: snap.currency,
    pdfCount: snap.pdfs.length,
    revision: snap.revision,
  };
}

function initialReviewHint_(snap) {
  let hint;
  if (snap.total === null || snap.total === undefined) {
    hint = 'No total detected: enter Corrected Amount and Corrected Currency before APPROVE.';
  } else if (!snap.currency) {
    hint = 'Total detected without currency: enter Corrected Currency before APPROVE.';
  } else {
    hint = `Check the best-effort detected total ${snap.total} ${snap.currency}, then set Decision.`;
  }
  if (!snap.pdfs.length) hint += ' No PDF attachment.';
  return hint;
}

function appendReviewRows_(sheet, rows) {
  if (!rows.length) return;
  const start = sheet.getLastRow() + 1;
  ensureRows_(sheet, start + rows.length - 1);
  formatReviewRows_(sheet, start, rows.length);
  sheet.getRange(start, 1, rows.length, REVIEW_HEADERS.length).setValues(rows);
}

function formatReviewRows_(sheet, start, count) {
  REVIEW_TEXT_COLUMNS.forEach(col => sheet.getRange(start, col, count, 1).setNumberFormat('@'));
  sheet.getRange(start, REVIEW_COL.decision, count, 1).setDataValidation(decisionValidationRule_());
}

function readReviewRows_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, REVIEW_HEADERS.length).getValues()
    .map((values, i) => reviewRowFromValues_(values, i + 2))
    .filter(row => row.messageId !== '');
}

function readReviewRowsById_(sheet, id) {
  return findReviewRowNumbers_(sheet, id).map(rowNumber => reviewRowFromValues_(
    sheet.getRange(rowNumber, 1, 1, REVIEW_HEADERS.length).getValues()[0], rowNumber));
}

function reviewRowFromValues_(v, rowNumber) {
  const at = col => v[col - 1];
  return {
    rowNumber,
    messageId: String(at(REVIEW_COL.messageId)).trim(),
    threadId: at(REVIEW_COL.threadId),
    detectedTotal: at(REVIEW_COL.detectedTotal),
    detectedCurrency: at(REVIEW_COL.detectedCurrency),
    pdfCount: at(REVIEW_COL.pdfCount),
    revision: at(REVIEW_COL.revision),
    decisionRaw: at(REVIEW_COL.decision),
    correctedAmount: at(REVIEW_COL.correctedAmount),
    correctedCurrency: at(REVIEW_COL.correctedCurrency),
    status: String(at(REVIEW_COL.status)).trim(),
    artifactIds: String(at(REVIEW_COL.artifactIds)),
  };
}

function groupReviewRowsById_(rows) {
  const map = new Map();
  rows.forEach(row => {
    if (!map.has(row.messageId)) map.set(row.messageId, []);
    map.get(row.messageId).push(row);
  });
  return map;
}

function findReviewRowNumbers_(sheet, id) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const out = [];
  sheet.getRange(2, REVIEW_COL.messageId, last - 1, 1).getValues().forEach((v, i) => {
    if (String(v[0]).trim() === id) out.push(i + 2);
  });
  return out;
}

/** Writes fields to the single Review row for `id`, located at write time. */
function updateReviewRow_(sheet, id, fields) {
  const rows = findReviewRowNumbers_(sheet, id);
  if (rows.length !== 1) {
    throw new Error(`Expected exactly one Review row for ${id}, found ${rows.length}; not writing.`);
  }
  updateReviewRowAt_(sheet, rows[0], id, fields);
}

function updateReviewRowAt_(sheet, rowNumber, id, fields) {
  const current = String(sheet.getRange(rowNumber, REVIEW_COL.messageId).getValue()).trim();
  if (current !== id) {
    throw new Error(`Review row ${rowNumber} no longer holds ${id}; not writing.`);
  }
  const byCol = new Map([[REVIEW_COL.updatedAt, new Date()]]);
  Object.keys(fields).forEach(key => {
    const col = REVIEW_COL[key];
    if (!col) throw new Error(`Unknown Review field ${key}`);
    byCol.set(col, fields[key]);
  });

  // Write each run of adjacent columns in one call (operator columns that are
  // not being changed are never rewritten).
  const cols = Array.from(byCol.keys()).sort((a, b) => a - b);
  let i = 0;
  while (i < cols.length) {
    let j = i;
    while (j + 1 < cols.length && cols[j + 1] === cols[j] + 1) j += 1;
    const start = cols[i];
    const width = cols[j] - start + 1;
    for (let c = start; c <= cols[j]; c++) {
      if (REVIEW_TEXT_COLUMNS.includes(c)) sheet.getRange(rowNumber, c).setNumberFormat('@');
    }
    sheet.getRange(rowNumber, start, 1, width).setValues([cols.slice(i, j + 1).map(c => byCol.get(c))]);
    i = j + 1;
  }

  // The script lock does not cover the Sheets UI: if rows were sorted or
  // inserted meanwhile, the write may have hit another message's row.
  const after = String(sheet.getRange(rowNumber, REVIEW_COL.messageId).getValue()).trim();
  if (after !== id) {
    throw new Error(`Review rows moved while writing ${id}: row ${rowNumber} now holds "${after}" and may ` +
      `have received fields meant for ${id} (${Object.keys(fields).join(', ')}). Run the Audit report.`);
  }
}

function markDuplicateReviewRows_(sheet, id, rows) {
  rows.forEach(row => updateReviewRowAt_(sheet, row.rowNumber, id, {
    decision: 'HOLD',
    status: REVIEW_STATUS.ERROR,
    reason: `${rows.length} Review rows share this Message ID; delete the extra rows, then refresh. Nothing was processed.`,
  }));
}

// ---------- Audit ----------

function collectAuditFindings_(sheets, folder) {
  const findings = [];
  const add = (kind, messageId, detail) => findings.push({kind, messageId, detail});

  const recordIds = [];
  const last = sheets.records.getLastRow();
  if (last >= 2) {
    sheets.records.getRange(2, 2, last - 1, 1).getValues()
      .forEach(v => { if (v[0] !== '') recordIds.push(String(v[0])); });
  }
  const recorded = new Set(recordIds);
  const recordCounts = countBy_(recordIds);
  recordCounts.forEach((n, id) => {
    if (n > 1) add('DUPLICATE_RECORDS', id, `${n} Records rows share this Message ID.`);
  });

  const reviewRows = readReviewRows_(sheets.review);
  groupReviewRowsById_(reviewRows).forEach((rows, id) => {
    if (rows.length > 1) add('DUPLICATE_REVIEW_ROWS', id, `${rows.length} Review rows share this Message ID.`);
  });
  reviewRows.forEach(row => {
    if (row.status === REVIEW_STATUS.IN_PROGRESS) {
      add('INTERRUPTED', row.messageId, `Review row ${row.rowNumber} is IN_PROGRESS (an earlier run stopped mid-way).`);
    } else if (row.status === REVIEW_STATUS.ERROR) {
      add('ERROR_ROW', row.messageId, `Review row ${row.rowNumber} is in ERROR; see Status Reason.`);
    } else if (row.status === REVIEW_STATUS.LABEL_PENDING) {
      add('LABEL_PENDING', row.messageId, `Review row ${row.rowNumber}: recorded, Gmail label not yet applied.`);
    }
    if ((row.status === REVIEW_STATUS.DONE || row.status === REVIEW_STATUS.LABEL_PENDING) &&
        !recorded.has(row.messageId)) {
      add('DONE_WITHOUT_RECORD', row.messageId,
        `Review row ${row.rowNumber} is ${row.status} but Records has no row for this Message ID.`);
    }
  });

  const names = new Map();
  const it = folder.getFiles();
  while (it.hasNext()) {
    const file = it.next();
    if (file.isTrashed()) continue;
    const name = String(file.getName());
    const match = name.match(/^([A-Za-z0-9]+)_SI(\d+)_/);
    if (!match) continue;
    names.set(name, (names.get(name) || 0) + 1);
    if (!recorded.has(match[1])) {
      add('ORPHAN_ARTIFACT', match[1], `Drive file "${name}" (${file.getUrl()}) has no Records row.`);
    }
  }
  names.forEach((n, name) => {
    if (n > 1) add('DUPLICATE_ARTIFACT', name.split('_SI')[0], `${n} live Drive files are named "${name}".`);
  });

  return findings;
}

const AUDIT_HEADERS = ['Checked At', 'Kind', 'Message ID', 'Detail'];

/** Fails before any write if an "Audit" sheet exists that this script did not create. */
function assertAuditSheetOwned_(ss) {
  const sheet = ss.getSheetByName(SAFE_INTAKE.auditSheet);
  if (!sheet || sheet.getLastRow() === 0) return;
  const header = sheet.getRange(1, 1, 1, AUDIT_HEADERS.length).getValues()[0].map(String);
  if (header.join('\u0001') !== AUDIT_HEADERS.join('\u0001') ||
      sheet.getLastColumn() > AUDIT_HEADERS.length) {
    throw new Error(`Safe Intake: a "${SAFE_INTAKE.auditSheet}" sheet with other content exists; ` +
      'rename it so the audit report does not overwrite it. Nothing was changed.');
  }
}

function writeAuditSheet_(ss, findings) {
  const sheet = ensureSheet_(ss, SAFE_INTAKE.auditSheet, AUDIT_HEADERS);
  clearDataRows_(sheet);
  const now = new Date();
  const rows = findings.length
    ? findings.map(f => [now, f.kind, f.messageId, f.detail])
    : [[now, 'OK', '', 'No inconsistencies found.']];
  ensureRows_(sheet, rows.length + 1);
  sheet.getRange(2, 3, rows.length, 1).setNumberFormat('@');
  sheet.getRange(2, 1, rows.length, 4).setValues(rows);
}

function countBy_(values) {
  const map = new Map();
  values.forEach(v => map.set(v, (map.get(v) || 0) + 1));
  return map;
}

// ---------- Shared v0.2 helpers ----------

function loadSafeIntakeConfig_(options) {
  const props = PropertiesService.getScriptProperties();
  const read = name => String(props.getProperty(name) || '').trim();
  const configError = msg => new Error(`Safe Intake configuration error: ${msg}. Nothing was changed.`);

  const sheetId = read('SHEET_ID');
  if (!sheetId) throw configError('missing Script Property SHEET_ID');
  const folderId = read('FOLDER_ID');
  if (!folderId) throw configError('missing Script Property FOLDER_ID');

  let maxThreads = 50;
  const rawMax = read('MAX_THREADS');
  if (rawMax) {
    if (!/^\d+$/.test(rawMax) || Number(rawMax) < 1 || Number(rawMax) > SAFE_INTAKE.maxThreadsLimit) {
      throw configError(`MAX_THREADS must be a whole number from 1 to ${SAFE_INTAKE.maxThreadsLimit} (found "${rawMax}")`);
    }
    maxThreads = Number(rawMax);
  }

  let ss;
  try {
    ss = SpreadsheetApp.openById(sheetId);
  } catch (err) {
    throw configError(`cannot open the spreadsheet in SHEET_ID (${errorMessage_(err)})`);
  }

  let folder = null;
  if (options.needsFolder) {
    try {
      folder = DriveApp.getFolderById(folderId);
    } catch (err) {
      throw configError(`cannot open the Drive folder in FOLDER_ID (${errorMessage_(err)})`);
    }
  }

  return {
    ss,
    folder,
    maxThreads,
    labelName: read('PROCESSED_LABEL') || CONFIG.defaultProcessedLabel,
  };
}

/** Runs fn under the script lock; returns LOCKED (and changes nothing) if busy. */
function withScriptLock_(fn) {
  const lock = LockService.getScriptLock();
  let acquired = false;
  try {
    acquired = lock.tryLock(SAFE_INTAKE.lockWaitMs);
  } catch (err) {
    acquired = false;
  }
  if (!acquired) {
    return {
      status: 'LOCKED',
      message: 'Another Safe Intake run holds the script lock. Nothing was changed; try again shortly.',
    };
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function sha256Hex_(value) {
  const bytes = typeof value === 'string'
    ? Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8)
    : Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value);
  return bytes.map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}

/** Stops email-controlled text from being interpreted as a Sheets formula. */
function sheetText_(value) {
  const s = String(value === null || value === undefined ? '' : value);
  return /^[=+\-@]/.test(s) ? `'${s}` : s;
}

function errorMessage_(err) {
  return err && err.message ? err.message : String(err);
}

function logErrorById_(ss, id, subject, err) {
  try {
    ss.getSheetByName(CONFIG.errorSheet).appendRow([
      new Date(),
      id,
      sheetText_(subject),
      err && err.stack ? err.stack : String(err),
    ]);
  } catch (logErr) {
    Logger.log(`Could not write to Errors sheet: ${errorMessage_(logErr)}`);
  }
}

function report_(ss, title, result) {
  Logger.log(`${title}: ${JSON.stringify(result)}`);
  try {
    ss.toast(JSON.stringify(result), title, 10);
  } catch (err) {
    // Toasts only work when the spreadsheet is open; the log line is enough.
  }
}
