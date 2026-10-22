// ════════════════════════════════════════════════════════════════════
// INVOICE SHEET HELPERS — shared by /api/invoice/* and the importer
//
// 🧾 Invoices columns B..S (row-array index in brackets):
//   B[0] Invoice #        C[1] Date Issued     D[2] Client        E[3] Description
//   F[4] Amount excl HST  G[5] HST             H[6] Total         I[7] HST? (Yes/No)
//   J[8] Due Date         K[9] Status          L[10] Date Paid    M[11] Notes
//   N[12] Revenue Category   O[13] Deposit Amount   P[14] Deposit Date Received
//   Q[15] Balance Due
//   R[16] Lines (JSON)        — added Oct 2026 (Migration 16): the line items,
//                               so an invoice can be reopened and edited
//   S[17] Payments (JSON)     — every payment: {date, method, amount, kind, ref, pending}
//
// Status values (K) — unchanged so the importer keeps working:
//   Unpaid · Awaiting Deposit · Deposit Received · Paid · Overdue · Cancelled
//   The UI shows 'Unpaid' as "Open".
//
// Money cells come back formatted ("$1,234.00") — always read through num().
// ════════════════════════════════════════════════════════════════════

import { readRange, writeRange, appendRows } from './_sheets.js';
import { num, isoDate, round2 } from './_payroll_sheet.js';

export const INVOICES_TAB = '🧾 Invoices';
export const TXN_TAB = '📒 Transactions';
export const INV_READ_RANGE = `'${INVOICES_TAB}'!B12:S`;

export const INV_COL = {
  invNum: 0, date: 1, client: 2, desc: 3, sub: 4, hst: 5, total: 6, hstFlag: 7,
  due: 8, status: 9, datePaid: 10, notes: 11, category: 12,
  depositAmount: 13, depositDate: 14, balanceDue: 15, lines: 16, payments: 17,
};

/** Payment methods that are cash-in-hand: the sale is booked immediately. */
export const IMMEDIATE_METHODS = new Set(['cash']);
/** Methods that will appear on a bank/card statement: booked by the importer's match. */
export const BANK_METHODS = new Set(['e-transfer', 'etransfer', 'cheque', 'card', 'debit', 'credit', 'bank']);

export function normalizeMethod(m) {
  const s = String(m || '').trim().toLowerCase().replace(/\s+/g, '-');
  if (s === 'etransfer' || s === 'e-transfer' || s === 'interac') return 'e-transfer';
  if (s === 'cash') return 'cash';
  if (s === 'cheque' || s === 'check') return 'cheque';
  if (s === 'card' || s === 'debit' || s === 'credit' || s === 'stripe' || s === 'square') return 'card';
  return s || 'e-transfer';
}

function safeJSON(v, fallback) {
  if (v == null || v === '') return fallback;
  try { const x = JSON.parse(v); return x == null ? fallback : x; } catch { return fallback; }
}

export function parseInvoiceRow(row, sheetRow) {
  const total = num(row[INV_COL.total]);
  const payments = safeJSON(row[INV_COL.payments], []);
  const lines = safeJSON(row[INV_COL.lines], null);
  const paidSoFar = round2((Array.isArray(payments) ? payments : []).reduce((s, p) => s + (Number(p.amount) || 0), 0));
  return {
    sheetRow,
    invNum: String(row[INV_COL.invNum] ?? '').trim(),
    date: isoDate(row[INV_COL.date]) || String(row[INV_COL.date] || ''),
    client: String(row[INV_COL.client] || ''),
    desc: String(row[INV_COL.desc] || ''),
    sub: num(row[INV_COL.sub]),
    hst: num(row[INV_COL.hst]),
    total,
    hstFlag: String(row[INV_COL.hstFlag] || 'Yes'),
    due: isoDate(row[INV_COL.due]) || String(row[INV_COL.due] || ''),
    status: String(row[INV_COL.status] || 'Unpaid').trim(),
    datePaid: isoDate(row[INV_COL.datePaid]),
    notes: String(row[INV_COL.notes] || ''),
    category: String(row[INV_COL.category] || ''),
    depositAmount: num(row[INV_COL.depositAmount]),
    depositDate: isoDate(row[INV_COL.depositDate]),
    balanceDue: row[INV_COL.balanceDue] == null || row[INV_COL.balanceDue] === '' ? round2(total - paidSoFar) : num(row[INV_COL.balanceDue]),
    lines: Array.isArray(lines) ? lines : null,
    payments: Array.isArray(payments) ? payments : [],
    paidSoFar,
  };
}

export async function loadInvoices(env, userId) {
  const res = await readRange(env, userId, INV_READ_RANGE);
  if (!res.ok) return { ok: false, error: res.error, invoices: [] };
  const invoices = [];
  for (let i = 0; i < res.values.length; i++) {
    const row = res.values[i];
    if (!row || String(row[INV_COL.invNum] ?? '').trim() === '') continue;
    invoices.push(parseInvoiceRow(row, 12 + i));
  }
  return { ok: true, invoices };
}

export async function findInvoice(env, userId, { invNum, sheetRow }) {
  const res = await loadInvoices(env, userId);
  if (!res.ok) return { ok: false, error: res.error };
  let inv = null;
  if (sheetRow) inv = res.invoices.find(i => i.sheetRow === Number(sheetRow)) || null;
  if (!inv && invNum) {
    const key = String(invNum).trim().toUpperCase();
    const hits = res.invoices.filter(i => i.invNum.toUpperCase() === key);
    inv = hits.length ? hits[hits.length - 1] : null;      // newest row wins on duplicates
    if (hits.length > 1) inv.duplicates = hits.length;
  }
  return { ok: true, invoice: inv, all: res.invoices };
}

/**
 * Status + balance from the payments list and total. Pure.
 *   nothing paid      → 'Unpaid' (or 'Awaiting Deposit' when a deposit is required)
 *   some paid         → 'Deposit Received'
 *   paid in full      → 'Paid'
 */
export function deriveStatus({ total, payments, depositRequired = 0, current = '' }) {
  if (String(current).toLowerCase() === 'cancelled') return { status: 'Cancelled', balanceDue: 0, paid: 0 };
  const paid = round2((payments || []).reduce((s, p) => s + (Number(p.amount) || 0), 0));
  const balanceDue = Math.max(0, round2(total - paid));
  let status;
  if (total > 0 && paid >= total - 0.005) status = 'Paid';
  else if (paid > 0) status = 'Deposit Received';
  else if (depositRequired > 0) status = 'Awaiting Deposit';
  else status = 'Unpaid';
  return { status, balanceDue, paid };
}

/**
 * Write the whole data row for an invoice (B..S) at sheetRow. `inv` is the
 * normalized shape used by the create/update endpoints.
 */
export async function writeInvoiceRow(env, userId, sheetRow, inv) {
  const row = [[
    inv.invNum, inv.date || '', inv.client || '', inv.desc || '',
    round2(inv.sub), round2(inv.hst), round2(inv.total), inv.hstFlag || 'Yes',
    inv.due || '', inv.status || 'Unpaid', inv.datePaid || '', inv.notes || '',
    inv.category || '', inv.depositAmount ? round2(inv.depositAmount) : '', inv.depositDate || '',
    round2(inv.balanceDue ?? inv.total),
    inv.lines ? JSON.stringify(inv.lines) : '',
    inv.payments && inv.payments.length ? JSON.stringify(inv.payments) : '',
  ]];
  // Sheets rejects a write past the grid; fall back to B..Q on an unmigrated sheet.
  let w = await writeRange(env, userId, `'${INVOICES_TAB}'!B${sheetRow}:S${sheetRow}`, row);
  if (!w.ok && /grid|range|column/i.test(w.error || '')) {
    w = await writeRange(env, userId, `'${INVOICES_TAB}'!B${sheetRow}:Q${sheetRow}`, [row[0].slice(0, 16)]);
    if (w.ok) w.truncated = true;
  }
  return w;
}

export async function appendInvoiceRow(env, userId, inv) {
  const row = [[
    inv.invNum, inv.date || '', inv.client || '', inv.desc || '',
    round2(inv.sub), round2(inv.hst), round2(inv.total), inv.hstFlag || 'Yes',
    inv.due || '', inv.status || 'Unpaid', inv.datePaid || '', inv.notes || '',
    inv.category || '', inv.depositAmount ? round2(inv.depositAmount) : '', inv.depositDate || '',
    round2(inv.balanceDue ?? inv.total),
    inv.lines ? JSON.stringify(inv.lines) : '',
    inv.payments && inv.payments.length ? JSON.stringify(inv.payments) : '',
  ]];
  let a = await appendRows(env, userId, `'${INVOICES_TAB}'!B12:S`, row);
  if (!a.ok && /grid|range|column/i.test(a.error || '')) {
    a = await appendRows(env, userId, `'${INVOICES_TAB}'!B12:Q`, [row[0].slice(0, 16)]);
    if (a.ok) a.truncated = true;
  }
  if (a.ok) {
    const m = /!B(\d+)/.exec(a.updates?.updatedRange || '');
    a.sheetRow = m ? parseInt(m[1], 10) : null;
  }
  return a;
}

/**
 * Book a cash payment against an invoice as a revenue row in 📒 Transactions.
 * The payment is gross; HST is split out in the invoice's own ratio so a
 * $50 deposit on a $565 (13%) invoice books $44.25 revenue + $5.75 HST.
 * Account 'Cash' keeps it off the bank reconciliation. Source Ref is unique
 * per (invoice, kind, date) so a re-send cannot double-book.
 */
export async function bookCashPayment(env, userId, inv, payment) {
  const gross = round2(Number(payment.amount) || 0);
  if (gross <= 0) return { ok: false, error: 'Payment amount must be positive' };
  const hstShare = inv.total > 0 && inv.hst > 0 ? round2(gross * (inv.hst / inv.total)) : 0;
  const net = round2(gross - hstShare);
  const kindLabel = payment.kind === 'deposit' ? 'deposit' : payment.kind === 'final' ? 'final payment' : 'payment';
  const ref = payment.ref || `INV-${inv.invNum}-${(payment.kind || 'pay').toUpperCase()}-${String(payment.date || '').replace(/-/g, '')}`;
  const row = [[
    payment.date,                                   // B Date
    inv.client || '',                               // C Party
    `Invoice ${inv.invNum} — ${kindLabel} (cash)`,  // D Description
    net,                                            // E Amount (signed, + = money in, excl HST)
    inv.category || 'Sales',                        // F Category
    hstShare > 0 ? 'Yes' : 'No',                    // G HST?
    hstShare,                                       // H HST
    'Cash',                                         // I Account
    'Invoice',                                      // J Source
    ref,                                            // K Source Ref (dedup)
    inv.invNum,                                     // L Related Invoice #
    'N/A',                                          // M Match Status (cash never hits a statement)
  ]];
  // Dedup: refuse if this ref already exists in the ledger
  const existing = await readRange(env, userId, `'${TXN_TAB}'!K12:K`);
  if (existing.ok && existing.values.some(r => String(r?.[0] || '') === ref)) {
    return { ok: true, duplicate: true, ref, net, hstShare };
  }
  const a = await appendRows(env, userId, `'${TXN_TAB}'!B12:M`, row);
  if (!a.ok) return { ok: false, error: 'Transactions write failed: ' + a.error };
  return { ok: true, ref, net, hstShare, range: a.updates?.updatedRange };
}
