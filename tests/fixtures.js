'use strict';

// Synthetic fixtures only. No real people, vendors, addresses or amounts.

const PDF_A = '%PDF-1.4 synthetic invoice A';
const PDF_B = '%PDF-1.4 synthetic receipt B';

function enInvoice(overrides = {}) {
  return Object.assign({
    id: '18f0a1b2c3d4e5f6',
    threadId: '18f0a1b2c3d4e5f6',
    from: 'Demo Vendor Billing <billing@vendor.example>',
    date: '2026-10-01T09:00:00Z',
    subject: 'Invoice - Demo Vendor #1001',
    body: 'Demo invoice for portfolio testing.\n\nGrand Total: USD 123.45\n\nSynthetic test data only.',
    attachments: [{name: 'invoice-1001.pdf', contentType: 'application/pdf', content: PDF_A}],
  }, overrides);
}

function enReceiptNoAmount(overrides = {}) {
  return Object.assign({
    id: '18f0a1b2c3d4e5f7',
    threadId: '18f0a1b2c3d4e5f7',
    from: 'receipts@shop.example',
    date: '2026-10-02T10:00:00Z',
    subject: 'Your receipt from Demo Shop',
    body: 'Thanks for your order. The amount is shown in the attached file.',
    attachments: [{name: 'receipt.pdf', contentType: 'application/pdf', content: PDF_B}],
  }, overrides);
}

function jpInvoiceYen(overrides = {}) {
  return Object.assign({
    id: '18f0a1b2c3d4e5f8',
    threadId: '18f0a1b2c3d4e5f8',
    from: '"デモ商事 経理部" <keiri@demo-shoji.example>',
    date: '2026-10-03T01:00:00Z',
    subject: '【請求書】Invoice No. 2026-001 デモ商事',
    body: 'いつもお世話になっております。\nご請求金額: ¥12,000（税込）\nお支払期限: 2026年10月31日\n※テスト用の架空データです。',
    attachments: [{name: '請求書_2026-001.pdf', contentType: 'application/pdf', content: '%PDF-1.4 synthetic JP'}],
  }, overrides);
}

function jpReceiptEnSuffix(overrides = {}) {
  // "5,500円" is not recognised by the best-effort parser -> needs correction.
  return Object.assign({
    id: '18f0a1b2c3d4e5f9',
    threadId: '18f0a1b2c3d4e5f9',
    from: 'no-reply@demo-store.example',
    date: '2026-10-04T02:00:00Z',
    subject: '領収書 Receipt #A-77',
    body: 'お支払い金額: 5,500円\nテスト用の架空データです。',
    attachments: [],
  }, overrides);
}

function nonMatching(overrides = {}) {
  return Object.assign({
    id: '18f0a1b2c3d4e600',
    threadId: '18f0a1b2c3d4e600',
    from: 'news@example.invalid',
    date: '2026-10-05T00:00:00Z',
    subject: 'Weekly newsletter',
    body: 'Total: USD 1.00',
    attachments: [],
  }, overrides);
}

module.exports = {enInvoice, enReceiptNoAmount, jpInvoiceYen, jpReceiptEnSuffix, nonMatching, PDF_A, PDF_B};
