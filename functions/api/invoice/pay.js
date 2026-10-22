// ════════════════════════════════════════════════════════════════════
// POST /api/invoice/pay
//
// Record a payment against an invoice.
//   { invNum, sheetRow?, amount, method: 'cash'|'e-transfer'|'cheque'|'card',
//     date: 'YYYY-MM-DD', kind?: 'deposit'|'final'|'payment' }
//
// cash          → books a revenue row in 📒 Transactions NOW (Account: Cash,
//                 HST split in the invoice's ratio, Related Invoice # set).
// e-transfer /
// cheque / card → recorded on the invoice as pending; the statement importer
//                 writes the ledger row when the deposit shows up and matches.
//
// Either way the invoice's Payments JSON (S), Balance Due (Q), Status (K),
// Deposit cols (O/P, first deposit) and Date Paid (L) are updated.
// ════════════════════════════════════════════════════════════════════
import { authenticateRequest, json, options } from '../../_shared.js';
import { findInvoice, writeInvoiceRow, bookCashPayment, deriveStatus, normalizeMethod, IMMEDIATE_METHODS } from '../../_invoices.js';
import { round2 } from '../../_payroll_sheet.js';

export const onRequestOptions = () => options();

export async function onRequestPost({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  const userId = auth.userId;

  let body;
  try { body = await request.json(); } catch { return json({ ok: false, error: 'Invalid JSON' }, 400); }

  const amount = round2(parseFloat(body.amount) || 0);
  const method = normalizeMethod(body.method);
  const date = String(body.date || '').trim();
  const kind = ['deposit', 'final', 'payment'].includes(body.kind) ? body.kind : 'payment';
  if (!(amount > 0)) return json({ ok: false, error: 'amount must be greater than 0' }, 400);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return json({ ok: false, error: 'date must be YYYY-MM-DD' }, 400);

  const found = await findInvoice(env, userId, { invNum: body.invNum, sheetRow: body.sheetRow });
  if (!found.ok) return json({ ok: false, error: found.error });
  const inv = found.invoice;
  if (!inv) return json({ ok: false, error: `Invoice ${body.invNum || ''} not found` }, 404);
  if (inv.status.toLowerCase() === 'cancelled') return json({ ok: false, error: 'Invoice is cancelled' }, 400);

  const alreadyPaid = inv.paidSoFar;
  const remaining = Math.max(0, round2(inv.total - alreadyPaid));
  if (amount > remaining + 0.01 && !body.allowOverpay) {
    return json({ ok: false, error: `That would overpay the invoice: $${remaining.toFixed(2)} is still owing (total $${inv.total.toFixed(2)}, paid $${alreadyPaid.toFixed(2)}). Edit the invoice first if the price changed.`, remaining }, 400);
  }

  const payment = { date, method, amount, kind, pending: !IMMEDIATE_METHODS.has(method), recordedAt: new Date().toISOString() };

  // Cash: book it now. Bank methods: the importer books it on match.
  let booked = null;
  if (IMMEDIATE_METHODS.has(method)) {
    booked = await bookCashPayment(env, userId, inv, payment);
    if (!booked.ok) return json({ ok: false, error: booked.error });
    payment.ref = booked.ref;
    payment.pending = false;
  }

  const payments = [...inv.payments, payment];
  const { status, balanceDue, paid } = deriveStatus({ total: inv.total, payments, depositRequired: inv.depositAmount, current: inv.status });

  const updated = {
    ...inv,
    payments,
    balanceDue,
    status,
    datePaid: status === 'Paid' ? date : inv.datePaid,
    // Keep the legacy deposit columns meaningful: first deposit's amount/date.
    depositAmount: inv.depositAmount || (kind === 'deposit' ? amount : 0),
    depositDate: inv.depositDate || (kind === 'deposit' ? date : ''),
  };
  const w = await writeInvoiceRow(env, userId, inv.sheetRow, updated);
  if (!w.ok) return json({ ok: false, error: 'Invoice update failed: ' + w.error, ledgerBooked: !!booked });

  return json({
    ok: true,
    invNum: inv.invNum, sheetRow: inv.sheetRow,
    payment, status, balanceDue, paid,
    ledger: booked ? { ref: booked.ref, net: booked.net, hst: booked.hstShare, duplicate: !!booked.duplicate } : null,
    note: booked ? null : `${method} payment recorded on the invoice; it will be booked to the ledger when the ${method === 'card' ? 'card payout' : 'bank deposit'} is imported and matched.`,
    truncated: !!w.truncated,
  });
}
