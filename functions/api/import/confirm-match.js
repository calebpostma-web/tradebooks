// ════════════════════════════════════════════════════════════════════
// POST /api/import/confirm-match
//
// User confirms that a deposit (row in 📒 Transactions) corresponds to an
// open invoice. This endpoint:
//   1. Writes the invoice number into Transactions col L (Related Invoice)
//   2. Sets Match Status (col M) to "Matched"
//   3. Flips the invoice's Status column to "Paid" in 🧾 Invoices
//
// Request body:
//   { txnRow: <int, 1-indexed sheet row>, invNum: <string> }
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

  if (!txnRow || txnRow < 12) return json({ ok: false, error: 'Invalid txnRow' }, 400);

  // If invNum is empty, just mark the transaction Unmatched (user dismissed the suggestion).
  if (!invNum) {
    const res = await writeRange(env, userId, `'${TXN_TAB}'!L${txnRow}:M${txnRow}`, [['', 'Unmatched']]);
    if (!res.ok) return json({ ok: false, error: 'Failed to update Transactions: ' + res.error });
    return json({ ok: true, action: 'dismissed' });
  }

  // Find the invoice row (scan B12:K500 for InvNum in col B).
  const invResult = await readRange(env, userId, `'${INVOICES_TAB}'!B12:K500`);
  if (!invResult.ok) return json({ ok: false, error: 'Failed to read Invoices: ' + invResult.error });

  let invoiceRow = null;
  for (let i = 0; i < invResult.values.length; i++) {
    if (String(invResult.values[i][0] || '') === invNum) {
      invoiceRow = 12 + i;
      break;
    }
  }
  if (!invoiceRow) return json({ ok: false, error: `Invoice #${invNum} not found` }, 404);

  // Write the match onto Transactions row (L = Related Invoice, M = Match Status).
  const txnUpdate = await writeRange(env, userId, `'${TXN_TAB}'!L${txnRow}:M${txnRow}`, [[invNum, 'Matched']]);
  if (!txnUpdate.ok) return json({ ok: false, error: 'Failed to update Transactions: ' + txnUpdate.error });

  // Flip Invoice status (col K = Status).
  const invUpdate = await writeRange(env, userId, `'${INVOICES_TAB}'!K${invoiceRow}`, [['Paid']]);
  if (!invUpdate.ok) return json({ ok: false, error: 'Failed to update Invoice status: ' + invUpdate.error });

  return json({ ok: true, action: 'matched', invNum, txnRow, invoiceRow });
}
