// ════════════════════════════════════════════════════════════════════
// POST /api/invoice/create
// Create an invoice row in the Invoices tab.
//
// CASH BASIS (Phase 2A): invoice creation does NOT write to the ledger.
// Revenue is recognized only when the deposit is matched to this invoice
// via /api/import/confirm-match (which writes the Transactions row using
// the Revenue Category stored on the invoice).
// ════════════════════════════════════════════════════════════════════

import { appendRows, writeRange } from '../../_sheets.js';
import { authenticateRequest, json, options } from '../../_shared.js';

const INVOICES_TAB = '🧾 Invoices';

export const onRequestOptions = () => options();

export async function onRequestPost({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  const userId = auth.userId;

  let body;
  try {
    body = await request.json();
  } catch {
    return json({ ok: false, error: 'Invalid JSON' }, 400);
  }

  const inv = body.invoice || {};
  const category = inv.category || 'Consulting Revenue';

  // Deposit / down payment fields. Stored gross (incl HST). Date may be blank
  // when the deposit is required-but-not-yet-received (kitchen-table flow).
  const depositAmount = parseFloat(inv.depositAmount) || 0;
  const depositDate = (inv.depositDate || '').trim();
  const total = parseFloat(inv.total) || 0;
  const balanceDue = depositAmount > 0
    ? Math.max(0, Math.round((total - Math.min(depositAmount, total)) * 100) / 100)
    : total;

  // Status state machine:
  //   - No deposit field set                                → 'Unpaid'   (legacy)
  //   - Deposit amount > 0, no date received yet            → 'Awaiting Deposit'
  //   - Deposit amount > 0 and date received                → 'Deposit Received'
  //   - 'Paid' is set later by the bank-match flow.
  let status = 'Unpaid';
  if (depositAmount > 0 && depositDate) status = 'Deposit Received';
  else if (depositAmount > 0) status = 'Awaiting Deposit';

  // Columns B-K: InvNum, Date, Client, Description, Subtotal, HST, Total, HSTFlag, Due, Status
  const invRow = [[
    inv.invNum || '',
    inv.dateVal || '',
    inv.client || '',
    inv.desc || '',
    parseFloat(inv.sub) || 0,
    parseFloat(inv.hstAmt) || 0,
    total,
    inv.hst || 'Yes',
    inv.dueVal || '',
    status,
  ]];

  const invResult = await appendRows(env, userId, `'${INVOICES_TAB}'!B12:K`, invRow);
  if (!invResult.ok) return json({ ok: false, error: 'Invoice write failed: ' + invResult.error });

  // Resolve the row number we just appended to so we can write the additive
  // columns (Revenue Category, deposit fields). Sheets returns updatedRange
  // looking like "'🧾 Invoices'!B45:K45".
  const match = /!B(\d+):K(\d+)/.exec(invResult.updates.updatedRange);
  if (match) {
    const row = parseInt(match[1]);
    // Col N — Revenue Category (existing behaviour)
    const catResult = await writeRange(env, userId, `'${INVOICES_TAB}'!N${row}`, [[category]]);
    if (!catResult.ok) console.warn('Invoice category write failed:', catResult.error);

    // Cols O, P, Q — Deposit Amount, Deposit Date Received, Balance Due.
    // Only write when a deposit is actually flagged so legacy-shaped sheets
    // (no headers in O/P/Q) don't get noise from no-deposit invoices.
    if (depositAmount > 0) {
      const depResult = await writeRange(
        env, userId,
        `'${INVOICES_TAB}'!O${row}:Q${row}`,
        [[depositAmount, depositDate || '', balanceDue]]
      );
      if (!depResult.ok) console.warn('Invoice deposit write failed:', depResult.error);
    }
  }

  return json({ ok: true, invNum: inv.invNum, status, balanceDue });
}
