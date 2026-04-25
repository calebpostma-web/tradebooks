// ════════════════════════════════════════════════════════════════════
// POST /api/import/confirm-match
//
// User confirms that a deposit (row in 📒 Transactions) corresponds to an
// open invoice. This endpoint:
//   1. Writes the invoice number into Transactions col L (Related Invoice)
//   2. Sets Match Status (col M) to "Matched"
//   3. Updates the invoice Status in 🧾 Invoices:
//        - 'deposit' leg: Unpaid/Awaiting Deposit → 'Deposit Received'
//          (back-fills Deposit Date Received in col P)
//        - 'final'   leg: anything → 'Paid' (sets Date Paid in col L)
//      The leg is taken from the request body when the front-end has it
//      (passed through from /api/import/statement's proposed match), or
//      inferred from the current invoice status as a fallback.
//
// Request body:
//   { txnRow: <int, 1-indexed sheet row>, invNum: <string>, leg?: 'deposit'|'final', txnDate?: 'YYYY-MM-DD' }
//
// Reject (mark "Unmatched, skip") uses txnRow + invNum: "" and just sets
// Match Status to Unmatched (no invoice update).
// ════════════════════════════════════════════════════════════════════

import { readRange, writeRange } from '../../_sheets.js';
import { authenticateRequest, json, options } from '../../_shared.js';

const TXN_TAB = '📒 Transactions';
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

  const txnRow = parseInt(body.txnRow, 10);
  const invNum = (body.invNum || '').trim();
  const legHint = (body.leg || '').toLowerCase();    // 'deposit' | 'final' | ''
  const txnDate = (body.txnDate || '').trim();        // ISO date for back-fill

  if (!txnRow || txnRow < 12) return json({ ok: false, error: 'Invalid txnRow' }, 400);

  // If invNum is empty, just mark the transaction Unmatched (user dismissed the suggestion).
  if (!invNum) {
    const res = await writeRange(env, userId, `'${TXN_TAB}'!L${txnRow}:M${txnRow}`, [['', 'Unmatched']]);
    if (!res.ok) return json({ ok: false, error: 'Failed to update Transactions: ' + res.error });
    return json({ ok: true, action: 'dismissed' });
  }

  // Read out to col Q so we get the current status + deposit fields.
  const invResult = await readRange(env, userId, `'${INVOICES_TAB}'!B12:Q500`);
  if (!invResult.ok) return json({ ok: false, error: 'Failed to read Invoices: ' + invResult.error });

  let invoiceRow = null;
  let invoiceCategory = '';
  let currentStatus = '';
  let depositAmount = 0;
  let depositDate = '';
  for (let i = 0; i < invResult.values.length; i++) {
    if (String(invResult.values[i][0] || '') === invNum) {
      invoiceRow = 12 + i;
      const r = invResult.values[i];
      // Col indices within B-Q: N=12 (Rev Cat), K=9 (Status), O=13 (Dep Amt), P=14 (Dep Date)
      invoiceCategory = String(r[12] || '') || 'Consulting Revenue';
      currentStatus = String(r[9] || '').trim();
      depositAmount = parseFloat(r[13]) || 0;
      depositDate = String(r[14] || '').trim();
      break;
    }
  }
  if (!invoiceRow) return json({ ok: false, error: `Invoice #${invNum} not found` }, 404);

  // Decide which leg this is. Trust the front-end hint if provided; otherwise
  // infer from current status: if a deposit is configured and not yet received,
  // this is the deposit leg.
  let leg = legHint;
  if (!leg) {
    const s = currentStatus.toLowerCase();
    if (s === 'awaiting deposit' || (depositAmount > 0 && !depositDate && s !== 'deposit received')) {
      leg = 'deposit';
    } else {
      leg = 'final';
    }
  }

  // Rewrite Transaction category (F) to the invoice's Revenue Category so the
  // deposit buckets correctly in Year-End (instead of generic 'Income Received'),
  // and populate Related Invoice (L) + Match Status (M).
  const txnCatUpdate = await writeRange(env, userId, `'${TXN_TAB}'!F${txnRow}`, [[invoiceCategory]]);
  if (!txnCatUpdate.ok) return json({ ok: false, error: 'Failed to update Transaction category: ' + txnCatUpdate.error });

  const txnUpdate = await writeRange(env, userId, `'${TXN_TAB}'!L${txnRow}:M${txnRow}`, [[invNum, 'Matched']]);
  if (!txnUpdate.ok) return json({ ok: false, error: 'Failed to update Transactions: ' + txnUpdate.error });

  // Update the invoice. Two cases:
  if (leg === 'deposit') {
    // First leg — flip status to 'Deposit Received' and back-fill the deposit date (col P).
    // We do this in two writes because Status (K) and Deposit Date (P) aren't contiguous.
    const statusUpdate = await writeRange(env, userId, `'${INVOICES_TAB}'!K${invoiceRow}`, [['Deposit Received']]);
    if (!statusUpdate.ok) return json({ ok: false, error: 'Failed to update Invoice status: ' + statusUpdate.error });
    if (txnDate) {
      const depDateUpdate = await writeRange(env, userId, `'${INVOICES_TAB}'!P${invoiceRow}`, [[txnDate]]);
      if (!depDateUpdate.ok) console.warn('Deposit date back-fill failed:', depDateUpdate.error);
    }
  } else {
    // Final leg — flip status to 'Paid' and set Date Paid (col L).
    const statusUpdate = await writeRange(env, userId, `'${INVOICES_TAB}'!K${invoiceRow}`, [['Paid']]);
    if (!statusUpdate.ok) return json({ ok: false, error: 'Failed to update Invoice status: ' + statusUpdate.error });
    if (txnDate) {
      const datePaidUpdate = await writeRange(env, userId, `'${INVOICES_TAB}'!L${invoiceRow}`, [[txnDate]]);
      if (!datePaidUpdate.ok) console.warn('Date Paid write failed:', datePaidUpdate.error);
    }
  }

  return json({
    ok: true,
    action: 'matched',
    invNum,
    txnRow,
    invoiceRow,
    category: invoiceCategory,
    leg,
    newStatus: leg === 'deposit' ? 'Deposit Received' : 'Paid',
  });
}
