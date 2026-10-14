'use strict';

// Acceptance tests for v0.2 Safe Intake. They execute the real Code.gs inside
// the node:vm Apps Script mock (tests/harness/gas-env.js).

const test = require('node:test');
const assert = require('node:assert/strict');
const {createEnv} = require('./harness/gas-env');
const fx = require('./fixtures');

const REVIEW_HEADERS = [
  'Message ID', 'Thread ID', 'Sender', 'Email Date', 'Subject',
  'Detected Total', 'Detected Currency', 'PDF Count', 'Revision',
  'Decision', 'Corrected Amount', 'Corrected Currency',
  'Status', 'Status Reason', 'Updated At', 'Saved PDF URLs', 'Artifact File IDs',
];

const plain = v => JSON.parse(JSON.stringify(v));

function seeded(messages) {
  const env = createEnv();
  (messages || [fx.enInvoice(), fx.enReceiptNoAmount(), fx.jpInvoiceYen(), fx.jpReceiptEnSuffix(), fx.nonMatching()])
    .forEach(m => env.addMessage(m));
  return env;
}

function review(env) {
  const rows = env.sheet('Review').rows();
  assert.deepEqual(rows[0], REVIEW_HEADERS);
  return rows.slice(1).map(r => Object.fromEntries(REVIEW_HEADERS.map((h, i) => [h, r[i] === undefined ? '' : r[i]])));
}

function reviewRow(env, id) {
  const matches = review(env).filter(r => String(r['Message ID']) === id);
  assert.equal(matches.length, 1, `exactly one Review row for ${id}`);
  return matches[0];
}

/** Simulates an operator typing into the Review sheet. */
function edit(env, id, values) {
  const sheet = env.sheet('Review');
  const rows = sheet.rows();
  const idx = rows.findIndex((r, i) => i > 0 && String(r[0]) === id);
  assert.ok(idx > 0, `row for ${id}`);
  Object.entries(values).forEach(([header, value]) => {
    const col = REVIEW_HEADERS.indexOf(header);
    assert.ok(col >= 0, header);
    sheet._cells[idx][col] = value;
  });
}

function records(env) {
  return env.sheet('Records').rows().slice(1);
}

function noSideEffects(env) {
  assert.deepEqual(env.sideEffectCounts(), {createFile: 0, appendRecords: 0, addLabel: 0, createLabel: 0});
}

function refreshAndApprove(env, id, extra = {}) {
  assert.equal(env.call('refreshReviewQueue').status, 'OK');
  edit(env, id, Object.assign({Decision: 'APPROVE'}, extra));
}

const EN = fx.enInvoice().id;
const EN_NO_AMOUNT = fx.enReceiptNoAmount().id;
const JP_YEN = fx.jpInvoiceYen().id;
const JP_EN_SUFFIX = fx.jpReceiptEnSuffix().id;

// ---------------------------------------------------------------------------

test('setupSafeIntake creates the Review sheet, decision dropdown and Records audit columns (idempotent)', () => {
  const env = createEnv();
  env.call('setupSafeIntake');
  env.call('setupSafeIntake');
  assert.deepEqual(env.sheet('Review').rows(), [REVIEW_HEADERS]);
  const validation = env.sheet('Review')._validations.find(v => v.col === 10);
  assert.ok(validation, 'Decision column has data validation');
  assert.deepEqual(plain(validation.rule.list), ['APPROVE', 'HOLD', 'SKIP']);
  assert.equal(validation.rule.allowInvalid, false);
  assert.deepEqual(env.sheet('Records').rows()[0].slice(10), ['Source', 'Amount Source', 'Review Revision']);
  noSideEffects(env);
});

test('refreshReviewQueue lists candidates for review with no Drive / Records / Gmail side effects', () => {
  const env = seeded();
  const result = plain(env.call('refreshReviewQueue'));
  assert.equal(result.status, 'OK');
  assert.equal(result.added, 4);

  const rows = review(env);
  assert.deepEqual(rows.map(r => r['Message ID']), [EN, EN_NO_AMOUNT, JP_YEN, JP_EN_SUFFIX]);
  rows.forEach(r => {
    assert.equal(r.Decision, '', 'no decision is pre-filled');
    assert.equal(r.Status, 'NEEDS_REVIEW');
    assert.match(String(r.Revision), /^[0-9a-f]{32}$/);
  });
  assert.equal(reviewRow(env, EN)['Detected Total'], 123.45);
  assert.equal(reviewRow(env, EN)['Detected Currency'], 'USD');
  assert.equal(reviewRow(env, JP_YEN)['Detected Total'], 12000);
  assert.equal(reviewRow(env, JP_YEN)['Detected Currency'], 'JPY');
  assert.match(reviewRow(env, EN_NO_AMOUNT)['Status Reason'], /Corrected Amount/);
  assert.match(reviewRow(env, JP_EN_SUFFIX)['Status Reason'], /Corrected Amount/);
  noSideEffects(env);
  assert.equal(env.sheet('Records').getLastRow(), 1, 'Records has only its header');
});

test('explicit APPROVE processes one message: deterministic PDF, one Records row, label, DONE', () => {
  const env = seeded();
  refreshAndApprove(env, EN);
  const result = plain(env.call('processApprovedReviewRows'));
  assert.equal(result.status, 'OK');
  assert.equal(result.processed, 1);
  assert.equal(result.errors, 0);

  const files = env.folder().liveFiles();
  assert.deepEqual(files.map(f => f.getName()), [`${EN}_SI1_invoice-1001.pdf`]);
  assert.equal(files[0]._blob.getDataAsString(), fx.PDF_A);

  const recs = records(env);
  assert.equal(recs.length, 1);
  assert.deepEqual(recs[0].slice(1, 13), [
    EN, fx.enInvoice().threadId, 'Demo Vendor Billing', 'billing@vendor.example',
    new Date(fx.enInvoice().date), 'Invoice - Demo Vendor #1001', 123.45, 'USD',
    files[0].getUrl(), 'safe-intake-v0.2', 'detected', reviewRow(env, EN).Revision,
  ]);
  assert.equal(env.countCalls('thread.addLabel'), 1);

  const row = reviewRow(env, EN);
  assert.equal(row.Status, 'DONE');
  assert.equal(row.Decision, 'APPROVE');
  assert.equal(row['Saved PDF URLs'], files[0].getUrl());
  assert.equal(row['Artifact File IDs'], files[0].getId());

  // Rows without a decision were not touched.
  assert.equal(reviewRow(env, JP_YEN).Status, 'NEEDS_REVIEW');
});

test('HOLD, SKIP and blank decisions never cause side effects', () => {
  const env = seeded();
  env.call('refreshReviewQueue');
  edit(env, EN, {Decision: 'HOLD'});
  edit(env, JP_YEN, {Decision: 'SKIP'});
  const result = plain(env.call('processApprovedReviewRows'));
  assert.equal(result.processed, 0);
  noSideEffects(env);
  assert.equal(reviewRow(env, EN).Status, 'HELD');
  assert.equal(reviewRow(env, JP_YEN).Status, 'SKIPPED');
  assert.equal(reviewRow(env, EN_NO_AMOUNT).Status, 'NEEDS_REVIEW');
});

test('missing amount: APPROVE without correction is held; a valid correction is then processed', () => {
  const env = seeded();
  refreshAndApprove(env, EN_NO_AMOUNT);
  let result = plain(env.call('processApprovedReviewRows'));
  assert.equal(result.held, 1);
  noSideEffects(env);
  let row = reviewRow(env, EN_NO_AMOUNT);
  assert.equal(row.Decision, 'HOLD');
  assert.equal(row.Status, 'HELD');
  assert.match(row['Status Reason'], /Corrected Amount/);

  edit(env, EN_NO_AMOUNT, {Decision: 'APPROVE', 'Corrected Amount': '1,045.60', 'Corrected Currency': 'usd'});
  result = plain(env.call('processApprovedReviewRows'));
  assert.equal(result.processed, 1);
  const recs = records(env);
  assert.equal(recs.length, 1);
  assert.deepEqual([recs[0][7], recs[0][8], recs[0][11]], [1045.6, 'USD', 'corrected']);
  assert.equal(reviewRow(env, EN_NO_AMOUNT).Status, 'DONE');
});

test('detected total without currency is not silently accepted', () => {
  const env = seeded([fx.enInvoice({body: 'Total: 99'})]);
  refreshAndApprove(env, EN);
  const result = plain(env.call('processApprovedReviewRows'));
  assert.equal(result.held, 1);
  assert.match(reviewRow(env, EN)['Status Reason'], /Corrected Currency/);
  noSideEffects(env);
});

test('Japanese fixtures: ¥ total is detected; 円 total needs a validated correction', () => {
  const env = seeded();
  env.call('refreshReviewQueue');
  edit(env, JP_YEN, {Decision: 'APPROVE'});
  edit(env, JP_EN_SUFFIX, {Decision: 'APPROVE', 'Corrected Amount': 5500, 'Corrected Currency': '円'});
  const result = plain(env.call('processApprovedReviewRows'));
  assert.equal(result.processed, 2);
  const recs = records(env);
  assert.deepEqual(recs.map(r => [r[1], r[7], r[8], r[11]]), [
    [JP_YEN, 12000, 'JPY', 'detected'],
    [JP_EN_SUFFIX, 5500, 'JPY', 'corrected'],
  ]);
  assert.equal(recs[0][3], 'デモ商事 経理部');
  assert.deepEqual(env.folder().liveFiles().map(f => f.getName()), [`${JP_YEN}_SI1_請求書_2026-001.pdf`]);
  assert.equal(reviewRow(env, JP_EN_SUFFIX)['Saved PDF URLs'], '');
});

test('malformed decisions and corrections are held with a reason and no side effects', () => {
  const cases = [
    [{Decision: 'yes'}, /Decision/],
    [{Decision: 'APPROVE', 'Corrected Amount': '12.3.4'}, /Corrected Amount/],
    [{Decision: 'APPROVE', 'Corrected Amount': 'abc'}, /Corrected Amount/],
    [{Decision: 'APPROVE', 'Corrected Amount': -5}, /Corrected Amount/],
    [{Decision: 'APPROVE', 'Corrected Amount': 0}, /Corrected Amount/],
    [{Decision: 'APPROVE', 'Corrected Amount': 10.005}, /decimal/],
    [{Decision: 'APPROVE', 'Corrected Amount': new Date('2026-01-01')}, /Corrected Amount/],
    [{Decision: 'APPROVE', 'Corrected Currency': 'BTC'}, /Corrected Currency/],
    [{Decision: 'APPROVE', 'Corrected Amount': 100.5, 'Corrected Currency': 'JPY'}, /decimal/],
  ];
  for (const [values, reason] of cases) {
    const env = seeded();
    refreshAndApprove(env, EN, values);
    const result = plain(env.call('processApprovedReviewRows'));
    assert.equal(result.processed, 0, JSON.stringify(values));
    noSideEffects(env);
    const row = reviewRow(env, EN);
    assert.equal(row.Decision, 'HOLD', JSON.stringify(values));
    assert.equal(row.Status, 'HELD');
    assert.match(row['Status Reason'], reason, JSON.stringify(values));
  }
});

test('stale approval: source changed after review is held at process time, then invalidated on refresh', () => {
  const env = seeded();
  refreshAndApprove(env, EN);
  const oldRevision = reviewRow(env, EN).Revision;
  env.messages.get(EN)._spec.body = 'Grand Total: USD 999.00';

  let result = plain(env.call('processApprovedReviewRows'));
  assert.equal(result.held, 1);
  noSideEffects(env);
  assert.match(reviewRow(env, EN)['Status Reason'], /changed since review/i);

  result = plain(env.call('refreshReviewQueue'));
  assert.equal(result.invalidated, 1);
  const row = reviewRow(env, EN);
  assert.equal(row.Decision, 'HOLD');
  assert.equal(row['Detected Total'], 999);
  assert.notEqual(row.Revision, oldRevision);

  edit(env, EN, {Decision: 'APPROVE'});
  result = plain(env.call('processApprovedReviewRows'));
  assert.equal(result.processed, 1);
  assert.equal(records(env)[0][7], 999);
});

test('refresh preserves decisions and corrections for unchanged rows and HOLDs changed rows', () => {
  const env = seeded();
  env.call('refreshReviewQueue');
  edit(env, EN, {Decision: 'APPROVE'});
  edit(env, EN_NO_AMOUNT, {Decision: 'HOLD', 'Corrected Amount': 45, 'Corrected Currency': 'EUR'});
  edit(env, JP_YEN, {Decision: 'APPROVE', 'Corrected Amount': 11000});
  env.messages.get(JP_YEN)._spec.attachments[0].content = '%PDF-1.4 re-issued';
  env.addMessage(fx.enInvoice({id: '18f0a1b2c3d4e601', threadId: '18f0a1b2c3d4e601', subject: 'Invoice #1002'}));

  const result = plain(env.call('refreshReviewQueue'));
  assert.deepEqual([result.added, result.unchanged, result.invalidated], [1, 3, 1]);
  assert.equal(reviewRow(env, EN).Decision, 'APPROVE');
  assert.deepEqual(
    [reviewRow(env, EN_NO_AMOUNT).Decision, reviewRow(env, EN_NO_AMOUNT)['Corrected Amount'], reviewRow(env, EN_NO_AMOUNT)['Corrected Currency']],
    ['HOLD', 45, 'EUR']);
  const changed = reviewRow(env, JP_YEN);
  assert.equal(changed.Decision, 'HOLD');
  assert.equal(changed.Status, 'HELD');
  assert.match(changed['Status Reason'], /changed/i);
  assert.equal(changed['Corrected Amount'], 11000, 'correction kept visible for re-review');
  assert.equal(reviewRow(env, '18f0a1b2c3d4e601').Status, 'NEEDS_REVIEW');
  assert.equal(review(env).length, 5);
  noSideEffects(env);
});

test('editing a system column instead of a correction column fails visible', () => {
  const env = seeded();
  refreshAndApprove(env, EN, {'Detected Total': 1});
  const result = plain(env.call('processApprovedReviewRows'));
  assert.equal(result.held, 1);
  assert.match(reviewRow(env, EN)['Status Reason'], /Corrected Amount/);
  noSideEffects(env);
});

test('repeat invocation does not duplicate rows, PDFs or labels', () => {
  const env = seeded();
  refreshAndApprove(env, EN);
  env.call('processApprovedReviewRows');
  const before = env.sideEffectCounts();
  const second = plain(env.call('processApprovedReviewRows'));
  assert.equal(second.processed, 0);
  env.call('refreshReviewQueue');
  const third = plain(env.call('processApprovedReviewRows'));
  assert.equal(third.processed, 0);
  assert.deepEqual(env.sideEffectCounts(), before);
  assert.equal(records(env).length, 1);
  assert.equal(env.folder().liveFiles().length, 1);
  assert.equal(reviewRow(env, EN).Status, 'DONE');
});

test('message already in Records (e.g. via legacy) is marked DONE without new artifacts', () => {
  const env = seeded();
  env.call('processInvoiceEmails'); // legacy run records everything
  const files = env.folder().liveFiles().length;
  const recs = records(env).length;
  refreshAndApprove(env, EN);
  const result = plain(env.call('processApprovedReviewRows'));
  assert.equal(result.alreadyRecorded, 1);
  assert.equal(env.folder().liveFiles().length, files);
  assert.equal(records(env).length, recs);
  assert.equal(reviewRow(env, EN).Status, 'DONE');
  assert.match(reviewRow(env, EN)['Status Reason'], /already/i);
});

test('lock unavailable: both v0.2 operations return LOCKED and change nothing', () => {
  const env = seeded();
  env.call('refreshReviewQueue');
  edit(env, EN, {Decision: 'APPROVE'});
  const snapshot = JSON.stringify(env.sheet('Review').rows());
  env.lock.unavailable = true;
  assert.equal(env.call('processApprovedReviewRows').status, 'LOCKED');
  assert.equal(env.call('refreshReviewQueue').status, 'LOCKED');
  assert.equal(JSON.stringify(env.sheet('Review').rows()), snapshot);
  noSideEffects(env);
});

test('concurrent run while processing is refused by the script lock; exactly one artifact set', () => {
  const env = seeded();
  refreshAndApprove(env, EN);
  const nested = [];
  env.faults['folder.createFile.before'] = () => {
    nested.push(plain(env.call('processApprovedReviewRows')));
    nested.push(plain(env.call('refreshReviewQueue')));
  };
  const result = plain(env.call('processApprovedReviewRows'));
  assert.equal(result.processed, 1);
  assert.deepEqual(nested.map(r => r.status), ['LOCKED', 'LOCKED']);
  assert.equal(env.folder().liveFiles().length, 1);
  assert.equal(records(env).length, 1);
  assert.equal(env.lock.holder, null, 'lock released');
});

test('failure after PDF creation but before Records append: retry reuses the PDF, no duplicate', () => {
  const env = seeded();
  refreshAndApprove(env, EN);
  let thrown = 0;
  env.faults['sheet.appendRow'] = ({sheet}) => {
    if (sheet === 'Records' && thrown++ === 0) throw new Error('Injected: Sheets unavailable');
  };
  let result = plain(env.call('processApprovedReviewRows'));
  assert.equal(result.errors, 1);
  assert.equal(env.folder().liveFiles().length, 1);
  assert.equal(records(env).length, 0);
  let row = reviewRow(env, EN);
  assert.equal(row.Status, 'ERROR');
  assert.equal(row.Decision, 'APPROVE', 'transient failure stays approved for a safe retry');
  assert.equal(env.sheet('Errors').rows().length, 2);

  result = plain(env.call('processApprovedReviewRows'));
  assert.equal(result.processed, 1);
  assert.equal(env.countCalls('folder.createFile'), 1, 'no second PDF');
  assert.equal(env.folder().liveFiles().length, 1);
  assert.equal(records(env).length, 1);
  row = reviewRow(env, EN);
  assert.equal(row.Status, 'DONE');
  assert.match(row['Status Reason'], /reused 1/);
});

test('ambiguous Drive failure (file created, call throws) and interrupted run recover without duplicates', () => {
  const env = seeded();
  refreshAndApprove(env, EN);
  env.faults['folder.createFile.after'] = 1;
  let result = plain(env.call('processApprovedReviewRows'));
  assert.equal(result.errors, 1);
  assert.equal(env.folder().liveFiles().length, 1);

  // Simulate a hard interruption (execution timeout) leaving IN_PROGRESS.
  edit(env, EN, {Status: 'IN_PROGRESS'});
  result = plain(env.call('processApprovedReviewRows'));
  assert.equal(result.processed, 1);
  assert.equal(env.countCalls('folder.createFile'), 1);
  assert.equal(records(env).length, 1);
});

test('conflicting prior artifacts are never overwritten or duplicated: ERROR + HOLD', () => {
  const scenarios = {
    'same name, different content': env => {
      env.folder().createFile(env.Utilities.newBlob('%PDF-1.4 something else', 'application/pdf', `${EN}_SI1_invoice-1001.pdf`));
    },
    'two files with the deterministic name': env => {
      for (let i = 0; i < 2; i++) {
        env.folder().createFile(env.Utilities.newBlob(fx.PDF_A, 'application/pdf', `${EN}_SI1_invoice-1001.pdf`));
      }
    },
    'legacy orphan without a Records row': env => {
      env.folder().createFile(env.Utilities.newBlob(fx.PDF_A, 'application/pdf', `${EN}-invoice-1001.pdf`));
    },
  };
  for (const [name, arrange] of Object.entries(scenarios)) {
    const env = seeded();
    refreshAndApprove(env, EN);
    arrange(env);
    const filesBefore = env.folder().liveFiles().length;
    const createsBefore = env.countCalls('folder.createFile');
    const result = plain(env.call('processApprovedReviewRows'));
    assert.equal(result.errors, 1, name);
    assert.equal(env.countCalls('folder.createFile'), createsBefore, name);
    assert.equal(env.folder().liveFiles().length, filesBefore, name);
    assert.equal(records(env).length, 0, name);
    const row = reviewRow(env, EN);
    assert.equal(row.Status, 'ERROR', name);
    assert.equal(row.Decision, 'HOLD', name);
    assert.match(row['Status Reason'], /reconcil/i, name);

    // A second run does not make it worse.
    env.call('processApprovedReviewRows');
    assert.equal(env.countCalls('folder.createFile'), createsBefore, name);
  }
});

test('a trashed artifact with the deterministic name is ignored (operator rollback)', () => {
  const env = seeded();
  refreshAndApprove(env, EN);
  env.folder().createFile(env.Utilities.newBlob('%PDF old', 'application/pdf', `${EN}_SI1_invoice-1001.pdf`)).setTrashed(true);
  const result = plain(env.call('processApprovedReviewRows'));
  assert.equal(result.processed, 1);
  assert.equal(env.folder().liveFiles().length, 1);
});

test('label failure after Records append is retried later without duplicating records', () => {
  const env = seeded();
  refreshAndApprove(env, EN);
  env.faults['thread.addLabel'] = 1;
  let result = plain(env.call('processApprovedReviewRows'));
  assert.equal(result.labelPending, 1);
  assert.equal(reviewRow(env, EN).Status, 'LABEL_PENDING');
  assert.equal(records(env).length, 1);

  result = plain(env.call('processApprovedReviewRows'));
  assert.equal(result.alreadyRecorded, 1);
  assert.equal(reviewRow(env, EN).Status, 'DONE');
  assert.equal(records(env).length, 1);
  assert.equal(env.folder().liveFiles().length, 1);
});

test('bad configuration fails before any side effect and releases nothing it did not take', () => {
  const cases = [
    [env => { delete env.properties.SHEET_ID; }, /SHEET_ID/],
    [env => { env.properties.SHEET_ID = 'missing-sheet'; }, /SHEET_ID/],
    [env => { delete env.properties.FOLDER_ID; }, /FOLDER_ID/],
    [env => { env.properties.FOLDER_ID = 'missing-folder'; }, /FOLDER_ID/],
    [env => { env.properties.MAX_THREADS = 'abc'; }, /MAX_THREADS/],
    [env => { env.properties.MAX_THREADS = '0'; }, /MAX_THREADS/],
    [env => { env.properties.MAX_THREADS = '501'; }, /MAX_THREADS/],
  ];
  for (const [arrange, pattern] of cases) {
    const env = seeded();
    env.call('refreshReviewQueue');
    edit(env, EN, {Decision: 'APPROVE'});
    arrange(env);
    assert.throws(() => env.call('processApprovedReviewRows'), pattern);
    noSideEffects(env);
    assert.equal(env.lock.holder, null);
  }
});

test('unexpected sheet layouts are refused instead of guessed', () => {
  let env = seeded();
  env.call('setupSheets');
  env.sheet('Records').getRange(1, 11).setValue('Notes');
  assert.throws(() => env.call('refreshReviewQueue'), /Records/);

  env = seeded();
  env.spreadsheets.get('sheet-1').insertSheet('Review').appendRow(['Something', 'Else']);
  assert.throws(() => env.call('refreshReviewQueue'), /Review/);
  noSideEffects(env);
});

test('duplicate Review rows for one message are refused', () => {
  const env = seeded();
  refreshAndApprove(env, EN);
  const sheet = env.sheet('Review');
  sheet.appendRow(sheet.rows()[1]);
  const result = plain(env.call('processApprovedReviewRows'));
  assert.equal(result.errors, 1);
  noSideEffects(env);
  review(env).filter(r => r['Message ID'] === EN).forEach(r => {
    assert.equal(r.Decision, 'HOLD');
    assert.equal(r.Status, 'ERROR');
  });
});

test('missing or trashed source message is held', () => {
  for (const mutate of [env => env.messages.delete(EN), env => { env.messages.get(EN)._spec.inTrash = true; }]) {
    const env = seeded();
    refreshAndApprove(env, EN);
    mutate(env);
    const result = plain(env.call('processApprovedReviewRows'));
    assert.equal(result.held, 1);
    noSideEffects(env);
  }
});

test('numeric-looking message IDs survive Sheets coercion (no duplicate on re-run)', () => {
  const id = '9990000000000001';
  const env = seeded([fx.enInvoice({id, threadId: id})]);
  refreshAndApprove(env, id);
  assert.equal(typeof env.sheet('Review').rows()[1][0], 'string');
  env.call('processApprovedReviewRows');
  assert.equal(records(env)[0][1], id);
  env.call('processApprovedReviewRows');
  env.call('processInvoiceEmails'); // legacy must also see it as processed
  assert.equal(records(env).length, 1);
  assert.equal(env.folder().liveFiles().length, 1);
});

test('formula-like email text is stored as inert text in the Review sheet', () => {
  const env = seeded([fx.enInvoice({subject: '=IMPORTXML("http://x.invalid") Invoice'})]);
  env.call('refreshReviewQueue');
  assert.ok(String(reviewRow(env, EN).Subject).startsWith("'="));
});

test('bounded execution defers remaining approved rows to the next run', () => {
  const env = seeded();
  env.call('refreshReviewQueue');
  edit(env, EN, {Decision: 'APPROVE'});
  edit(env, JP_YEN, {Decision: 'APPROVE'});
  env.evalInScript('SAFE_INTAKE.maxRowsPerRun = 1');
  let result = plain(env.call('processApprovedReviewRows'));
  assert.deepEqual([result.processed, result.deferred], [1, 1]);
  result = plain(env.call('processApprovedReviewRows'));
  assert.deepEqual([result.processed, result.deferred], [1, 0]);
  assert.equal(records(env).length, 2);
});

test('auditSafeIntake reports interrupted rows and orphan artifacts without mutating Gmail/Drive/Records', () => {
  const env = seeded();
  refreshAndApprove(env, EN);
  env.faults['sheet.appendRow'] = ({sheet}) => { if (sheet === 'Records') throw new Error('boom'); };
  env.call('processApprovedReviewRows');
  env.faults = {};
  edit(env, EN, {Decision: 'SKIP', Status: 'IN_PROGRESS'});
  env.call('processApprovedReviewRows');
  assert.equal(reviewRow(env, EN).Status, 'IN_PROGRESS', 'interrupted row is not silently marked SKIPPED');
  assert.match(reviewRow(env, EN)['Status Reason'], /interrupted.*Audit/i);

  const before = env.sideEffectCounts();
  const result = plain(env.call('auditSafeIntake'));
  assert.deepEqual(env.sideEffectCounts(), before);
  const kinds = result.findings.map(f => f.kind).sort();
  assert.deepEqual(kinds, ['INTERRUPTED', 'ORPHAN_ARTIFACT']);
  assert.ok(env.sheet('Audit').getLastRow() >= 3);
});

test('legacy processInvoiceEmails is NOT review-gated (documented) but skips v0.2-recorded messages', () => {
  const env = seeded();
  refreshAndApprove(env, EN);
  edit(env, JP_YEN, {Decision: 'SKIP'});
  env.call('processApprovedReviewRows');
  const result = plain(env.call('processInvoiceEmails'));
  assert.equal(result.skipped, 1, 'the v0.2-recorded message is not duplicated');
  assert.equal(result.processed, 3, 'legacy path still processes unreviewed / SKIPped messages');
});

test('onOpen exposes the v0.2 entry points without wiring them to the legacy process path', () => {
  const env = createEnv();
  env.call('onOpen');
  const items = [];
  env.ui.menus.forEach(m => {
    m.items.forEach(i => i.fn && items.push(i));
    m.subMenus.forEach(s => s.items.forEach(i => i.fn && items.push(i)));
  });
  const fns = items.map(i => i.fn);
  for (const fn of ['setupSafeIntake', 'refreshReviewQueue', 'processApprovedReviewRows', 'auditSafeIntake']) {
    assert.ok(fns.includes(fn), fn);
  }
  const legacy = items.find(i => i.fn === 'processInvoiceEmails');
  assert.match(legacy.label, /legacy|no review/i);
  const src = env.evalInScript('processApprovedReviewRows.toString() + refreshReviewQueue.toString()');
  assert.doesNotMatch(src, /processInvoiceEmails\(/);
});

// ---------------------------------------------------------------------------
// Added after independent review (cycle 1).

test('review fix: a pre-existing, foreign "Audit" sheet is refused and left untouched', () => {
  const env = seeded();
  const audit = env.spreadsheets.get('sheet-1').insertSheet('Audit');
  audit.appendRow(['My notes', 'Owner']);
  audit.appendRow(['keep me', 'operator']);
  assert.throws(() => env.call('auditSafeIntake'), /Audit/);
  assert.deepEqual(audit.rows(), [['My notes', 'Owner'], ['keep me', 'operator']]);
});

test('review fix: Review rows moved during a status write fail visible; Audit flags the stray DONE', () => {
  const env = seeded();
  refreshAndApprove(env, EN);
  const sheet = env.sheet('Review');
  let swapped = false;
  env.faults['range.setValues'] = ({sheet: name, values}) => {
    if (name === 'Review' && !swapped && values[0].includes('DONE')) {
      swapped = true; // operator sorts the sheet at the worst moment
      [sheet._cells[1], sheet._cells[2]] = [sheet._cells[2], sheet._cells[1]];
    }
  };
  const result = plain(env.call('processApprovedReviewRows'));
  env.faults = {};
  assert.equal(result.errors, 1);
  assert.equal(reviewRow(env, EN).Status, 'ERROR');
  assert.equal(reviewRow(env, EN_NO_AMOUNT).Status, 'DONE', 'the stray write landed on the other row');
  assert.ok(env.sheet('Errors').rows().some(r => /moved/.test(String(r[3]))));

  const audit = plain(env.call('auditSafeIntake'));
  assert.ok(audit.findings.some(f => f.kind === 'DONE_WITHOUT_RECORD' && f.messageId === EN_NO_AMOUNT));

  const retry = plain(env.call('processApprovedReviewRows'));
  assert.equal(retry.alreadyRecorded, 1);
  assert.equal(reviewRow(env, EN).Status, 'DONE');
  assert.equal(records(env).length, 1);
  assert.equal(env.folder().liveFiles().length, 1);
});

test('review fix: a row with leftover PDFs stays visibly ERROR through HOLD / SKIP', () => {
  const env = seeded();
  refreshAndApprove(env, EN);
  env.faults['sheet.appendRow'] = ({sheet}) => { if (sheet === 'Records') throw new Error('boom'); };
  env.call('processApprovedReviewRows');
  env.faults = {};
  assert.notEqual(reviewRow(env, EN)['Artifact File IDs'], '');

  edit(env, EN, {'Corrected Amount': 'abc'});
  env.call('processApprovedReviewRows');
  let row = reviewRow(env, EN);
  assert.deepEqual([row.Decision, row.Status], ['HOLD', 'ERROR']);
  assert.match(row['Status Reason'], /Audit/);

  edit(env, EN, {Decision: 'SKIP'});
  env.call('processApprovedReviewRows');
  row = reviewRow(env, EN);
  assert.equal(row.Status, 'ERROR');
  assert.doesNotMatch(row['Status Reason'], /no PDF/);
});

test('review fix: an already-recorded message is not labelled when its approval is stale', () => {
  const env = seeded();
  env.call('refreshReviewQueue');
  env.call('processInvoiceEmails'); // legacy records (and labels) everything
  edit(env, EN, {Decision: 'APPROVE'});
  env.messages.get(EN)._spec.body = 'Grand Total: USD 1.00';
  const labelsBefore = env.countCalls('thread.addLabel');
  const result = plain(env.call('processApprovedReviewRows'));
  assert.equal(result.held, 1);
  assert.equal(env.countCalls('thread.addLabel'), labelsBefore);
  assert.equal(reviewRow(env, EN).Decision, 'HOLD');
});

test('review fix: Records read-back mismatch is ERROR + HOLD, never silently DONE', () => {
  const env = seeded();
  refreshAndApprove(env, EN);
  env.faults['sheet.appendRow'] = ({sheet, values}) => {
    if (sheet === 'Records') values[1] = 'mangled-by-sheets';
  };
  const result = plain(env.call('processApprovedReviewRows'));
  assert.equal(result.errors, 1);
  const row = reviewRow(env, EN);
  assert.deepEqual([row.Decision, row.Status], ['HOLD', 'ERROR']);
  assert.match(row['Status Reason'], /read back/);
});

test('review fix: refresh of DONE, previously-failed and duplicated rows', () => {
  const env = seeded();
  env.call('refreshReviewQueue');
  edit(env, EN, {Decision: 'APPROVE'});
  env.call('processApprovedReviewRows');
  assert.equal(reviewRow(env, EN).Status, 'DONE');
  // JP_YEN: an earlier attempt failed after creating a PDF.
  edit(env, JP_YEN, {Decision: 'APPROVE', 'Corrected Amount': 11000, Status: 'ERROR', 'Artifact File IDs': 'file-x'});
  const sheet = env.sheet('Review');
  sheet.appendRow(sheet.rows().find(r => r[0] === JP_EN_SUFFIX));

  env.messages.get(EN)._spec.body = 'Grand Total: USD 5.00';
  env.messages.get(JP_YEN)._spec.body = 'ご請求金額: ¥13,000';
  const result = plain(env.call('refreshReviewQueue'));
  assert.equal(result.ambiguous, 1);

  const done = reviewRow(env, EN);
  assert.deepEqual([done.Status, done.Decision], ['DONE', 'APPROVE']);
  assert.match(done['Status Reason'], /Records were not modified/);

  const failed = reviewRow(env, JP_YEN);
  assert.deepEqual([failed.Status, failed.Decision], ['ERROR', 'HOLD']);
  assert.match(failed['Status Reason'], /Corrected Amount.*still filled in/);
  assert.match(failed['Status Reason'], /Audit/);

  review(env).filter(r => r['Message ID'] === JP_EN_SUFFIX)
    .forEach(r => assert.deepEqual([r.Status, r.Decision], ['ERROR', 'HOLD']));
});

test('review fix: runtime budget defers work in both refresh and process', () => {
  const env = seeded();
  env.call('refreshReviewQueue');
  edit(env, EN, {Decision: 'APPROVE'});
  env.evalInScript('SAFE_INTAKE.maxRuntimeMs = -1');
  assert.equal(plain(env.call('processApprovedReviewRows')).deferred, 1);
  env.addMessage(fx.enInvoice({id: '18f0a1b2c3d4e601', threadId: '18f0a1b2c3d4e601'}));
  const refresh = plain(env.call('refreshReviewQueue'));
  assert.deepEqual([refresh.added, refresh.deferred], [0, 5]);
  noSideEffects(env);
});

test('review fix: audit reports duplicate Records rows and duplicate artifacts', () => {
  const env = seeded();
  refreshAndApprove(env, EN);
  env.call('processApprovedReviewRows');
  const recs = env.sheet('Records');
  recs.appendRow(recs.rows()[1]);
  const name = env.folder().liveFiles()[0].getName();
  env.folder().createFile(env.Utilities.newBlob(fx.PDF_A, 'application/pdf', name));
  const kinds = plain(env.call('auditSafeIntake')).findings.map(f => f.kind).sort();
  assert.deepEqual(kinds, ['DUPLICATE_ARTIFACT', 'DUPLICATE_RECORDS']);
});

test('review fix: e-notation-looking hex Message IDs are kept as text', () => {
  const id = '1800000000e00001';
  const env = seeded([fx.enInvoice({id, threadId: id})]);
  refreshAndApprove(env, id);
  env.call('processApprovedReviewRows');
  env.call('processApprovedReviewRows');
  assert.equal(records(env).length, 1);
  assert.equal(records(env)[0][1], id);
});

test('review fix (round 2): a file ID belonging to another message is reported, not adopted or trashed', () => {
  const env = seeded();
  refreshAndApprove(env, EN);
  env.call('processApprovedReviewRows');
  const enFile = env.folder().liveFiles()[0];
  edit(env, JP_YEN, {Decision: 'APPROVE', 'Artifact File IDs': enFile.getId()});
  const creates = env.countCalls('folder.createFile');
  const result = plain(env.call('processApprovedReviewRows'));
  assert.equal(result.errors, 1);
  assert.equal(env.countCalls('folder.createFile'), creates);
  assert.equal(enFile.isTrashed(), false);
  assert.equal(records(env).length, 1);
  assert.match(reviewRow(env, JP_YEN)['Status Reason'], /belongs to another message/);
});
