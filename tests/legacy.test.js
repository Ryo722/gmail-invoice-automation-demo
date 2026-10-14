'use strict';

// Regression tests for the pre-v0.2 public entry points. These were written
// against main@f194ed9 and must keep passing unchanged.

const test = require('node:test');
const assert = require('node:assert/strict');
const {createEnv} = require('./harness/gas-env');
const fx = require('./fixtures');

function seeded() {
  const env = createEnv();
  env.addMessage(fx.enInvoice());
  env.addMessage(fx.enReceiptNoAmount());
  env.addMessage(fx.jpInvoiceYen());
  env.addMessage(fx.nonMatching());
  return env;
}

test('legacy: setupSheets creates Records / Preview / Errors with original headers', () => {
  const env = createEnv();
  env.call('setupSheets');
  assert.deepEqual(env.sheet('Records').rows()[0], [
    'Processed At', 'Message ID', 'Thread ID', 'Sender Name', 'Sender Email',
    'Email Date', 'Subject', 'Detected Total', 'Currency', 'Saved PDF URLs',
  ]);
  assert.deepEqual(env.sheet('Preview').rows()[0], [
    'Message ID', 'Sender', 'Email Date', 'Subject', 'Detected Total', 'Currency', 'PDF Count',
  ]);
  assert.deepEqual(env.sheet('Errors').rows()[0], ['Occurred At', 'Message ID', 'Subject', 'Error']);
});

test('legacy: previewInvoiceEmails lists matching messages without side effects', () => {
  const env = seeded();
  const result = env.call('previewInvoiceEmails');
  assert.equal(result.matchedMessages, 3);

  const rows = env.sheet('Preview').rows().slice(1);
  assert.deepEqual(rows.map(r => r[0]), [fx.enInvoice().id, fx.enReceiptNoAmount().id, fx.jpInvoiceYen().id]);
  assert.deepEqual(rows.map(r => [r[4], r[5], r[6]]), [
    [123.45, 'USD', 1],
    ['', '', 1],
    [12000, 'JPY', 1],
  ]);
  assert.deepEqual(env.sideEffectCounts(), {createFile: 0, appendRecords: 0, addLabel: 0, createLabel: 0});
});

test('legacy: processInvoiceEmails writes records, PDFs and labels, then is duplicate-safe', () => {
  const env = seeded();
  const first = env.call('processInvoiceEmails');
  assert.deepEqual({...first}, {matchedMessages: 3, processed: 3, skipped: 0, failed: 0});

  const records = env.sheet('Records').rows().slice(1);
  assert.equal(records.length, 3);
  assert.deepEqual(records[0].slice(1, 10), [
    fx.enInvoice().id, fx.enInvoice().threadId, 'Demo Vendor Billing', 'billing@vendor.example',
    new Date(fx.enInvoice().date), 'Invoice - Demo Vendor #1001', 123.45, 'USD',
    'https://drive.example.invalid/file/file-1',
  ]);
  // Legacy (unreviewed) path records a missing total as blank.
  assert.equal(records[1][7], '');
  assert.deepEqual(env.folder().liveFiles().map(f => f.getName()), [
    `${fx.enInvoice().id}-invoice-1001.pdf`,
    `${fx.enReceiptNoAmount().id}-receipt.pdf`,
    `${fx.jpInvoiceYen().id}-請求書_2026-001.pdf`,
  ]);
  assert.equal(env.countCalls('thread.addLabel'), 3);
  assert.ok(env.labels.has('portfolio/processed'));

  const second = env.call('processInvoiceEmails');
  assert.deepEqual({...second}, {matchedMessages: 3, processed: 0, skipped: 3, failed: 0});
  assert.equal(env.sheet('Records').rows().length, 4);
  assert.equal(env.folder().liveFiles().length, 3);
});

test('legacy: processInvoiceEmails logs per-message failures to Errors and continues', () => {
  const env = seeded();
  env.faults['folder.createFile.before'] = 1;
  const result = env.call('processInvoiceEmails');
  assert.equal(result.failed, 1);
  assert.equal(result.processed, 2);
  const errors = env.sheet('Errors').rows().slice(1);
  assert.equal(errors.length, 1);
  assert.equal(errors[0][1], fx.enInvoice().id);
});

test('legacy: missing SHEET_ID still throws the original error', () => {
  const env = seeded();
  delete env.properties.SHEET_ID;
  assert.throws(() => env.call('processInvoiceEmails'), /Missing required Script Property: SHEET_ID/);
});

test('legacy: parseTotal_ best-effort extraction is unchanged', () => {
  const env = createEnv();
  const parse = t => ({...env.call('parseTotal_', t)});
  assert.deepEqual(parse('Grand Total: USD 123.45'), {total: 123.45, currency: 'USD'});
  assert.deepEqual(parse('Amount due: €1,234.50'), {total: 1234.5, currency: 'EUR'});
  assert.deepEqual(parse('Total: 99'), {total: 99, currency: ''});
  assert.deepEqual(parse('ご請求金額: ¥12,000'), {total: 12000, currency: 'JPY'});
  assert.deepEqual(parse('お支払い金額: 5,500円'), {total: null, currency: null});
  assert.deepEqual(parse('no amount here'), {total: null, currency: null});
});

test('legacy: onOpen keeps the original menu entry points', () => {
  const env = createEnv();
  env.call('onOpen');
  const fns = [];
  env.ui.menus.forEach(m => {
    m.items.forEach(i => i.fn && fns.push(i.fn));
    m.subMenus.forEach(s => s.items.forEach(i => i.fn && fns.push(i.fn)));
  });
  for (const fn of ['setupSheets', 'previewInvoiceEmails', 'processInvoiceEmails']) {
    assert.ok(fns.includes(fn), `menu still exposes ${fn}`);
  }
});
