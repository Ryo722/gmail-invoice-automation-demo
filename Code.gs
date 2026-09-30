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
    .addItem('Setup sheets', 'setupSheets')
    .addItem('Preview matching emails', 'previewInvoiceEmails')
    .addItem('Process emails', 'processInvoiceEmails')
    .addToUi();
}

function setupSheets() {
  const ss = getSpreadsheet_();
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

function findCandidateMessages_() {
  const maxThreads = Number(getScriptProperty_('MAX_THREADS', '50'));
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
