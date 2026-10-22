// ════════════════════════════════════════════════════════════════════
// POST /api/invoice/create
// Create an invoice row in 🧾 Invoices — or UPDATE an existing one in place
// when body.invoice.sheetRow is given (the wizard reopened it).
//
// CASH BASIS: creating/updating an invoice does NOT write to the ledger.
// Revenue reaches 📒 Transactions either when a bank deposit is imported
// and matched (/api/import/confirm-match) or when a cash payment is
// recorded (/api/invoice/pay). Updating an invoice never touches payments
// already recorded; it recomputes status/balance from them.
//
// Stored per row: the header fields (B..Q), the line items as JSON (R)
// so the invoice can be reopened and edited, and payments as JSON (S).
// ════════════════════════════════════════════════════════════════════

import { authenticateRequest, json, options } from '../../_shared.js';
import { findInvoice, writeInvoiceRow, appendInvoiceRow, deriveStatus } from '../../_invoices.js';
import { round2 } from '../../_payroll_sheet.js';

export const onRequestOptions = () => options();

export async function onRequestPost({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  const userId = auth.userId;

  let body;
  try { body = await request.json(); } catch { return json({ ok: false, error: 'Invalid JSON' }, 400); }

  const inv = body.invoice || {};
  const invNum = String(inv.invNum || '').trim();
  if (!invNum) return json({ ok: false, error: 'Invoice number is required' }, 400);

  const category = inv.category || 'Consulting Revenue';
  const sub = round2(parseFloat(inv.sub) || 0);
  const hstAmt = round2(parseFloat(inv.hstAmt) || 0);
  const total = round2(parseFloat(inv.total) || (sub + hstAmt));
  const lines = Array.isArray(inv.lines) ? inv.lines.filter(l => l && (l.desc || l.qty || l.rate)) : null;

  // Legacy deposit fields from the wizard's "deposit required" toggle.
  // Amount = what's required up front (gross). Date = received date, or ''.
  const depositRequired = round2(parseFloat(inv.depositAmount) || 0);
  const depositDate = String(inv.depositDate || '').trim();

  // ── UPDATE path ──────────────────────────────────────────────────
  if (inv.sheetRow) {
    const found = await findInvoice(env, userId, { sheetRow: inv.sheetRow });
    if (!found.ok) return json({ ok: false, error: found.error });
    const existing = found.invoice;
    if (!existing) return json({ ok: false, error: `No invoice at row ${inv.sheetRow}` }, 404);

    // Payments already recorded stay as they are. If the legacy deposit date
    // is being set for the first time and there's no payment yet, treat it as
    // a pending e-transfer deposit so status/balance line up with the old flow.
    let payments = existing.payments.slice();
    if (depositRequired > 0 && depositDate && payments.length === 0 && !existing.depositDate) {
      payments.push({ date: depositDate, method: 'e-transfer', amount: Math.min(depositRequired, total), kind: 'deposit', pending: true, recordedAt: new Date().toISOString() });
    }
    const { status, balanceDue, paid } = deriveStatus({ total, payments, depositRequired, current: existing.status });

    const updated = {
      ...existing,
      invNum, date: inv.dateVal || existing.date, client: inv.client || existing.client,
      desc: inv.desc || existing.desc, sub, hst: hstAmt, total,
      hstFlag: inv.hst || existing.hstFlag || 'Yes',
      due: inv.dueVal || existing.due,
      notes: inv.notes != null ? String(inv.notes) : existing.notes,
      category,
      depositAmount: depositRequired || existing.depositAmount || 0,
      depositDate: existing.depositDate || depositDate || '',
      payments, balanceDue, status,
      datePaid: status === 'Paid' ? (existing.datePaid || payments[payments.length - 1]?.date || '') : '',
      lines: lines || existing.lines,
    };
    const w = await writeInvoiceRow(env, userId, existing.sheetRow, updated);
    if (!w.ok) return json({ ok: false, error: 'Invoice update failed: ' + w.error });
    return json({ ok: true, updated: true, invNum, sheetRow: existing.sheetRow, status, balanceDue, paid, truncated: !!w.truncated });
  }

  // ── CREATE path ──────────────────────────────────────────────────
  const payments = [];
  if (depositRequired > 0 && depositDate) {
    payments.push({ date: depositDate, method: 'e-transfer', amount: Math.min(depositRequired, total), kind: 'deposit', pending: true, recordedAt: new Date().toISOString() });
  }
  const { status, balanceDue, paid } = deriveStatus({ total, payments, depositRequired });

  const row = {
    invNum, date: inv.dateVal || '', client: inv.client || '', desc: inv.desc || '',
    sub, hst: hstAmt, total, hstFlag: inv.hst || 'Yes', due: inv.dueVal || '',
    status, datePaid: '', notes: inv.notes ? String(inv.notes) : '', category,
    depositAmount: depositRequired, depositDate, balanceDue, lines, payments,
  };
  const a = await appendInvoiceRow(env, userId, row);
  if (!a.ok) return json({ ok: false, error: 'Invoice write failed: ' + a.error });

  return json({ ok: true, created: true, invNum, sheetRow: a.sheetRow, status, balanceDue, paid, truncated: !!a.truncated });
}
