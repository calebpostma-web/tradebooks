// ════════════════════════════════════════════════════════════════════
// POST /api/invoice/create
// Create an invoice row in the Invoices tab.
//
// CASH BASIS (Phase 2A): invoice creation does NOT write to the ledger.
// Revenue is recognized only when the deposit is matched to this invoice
// via /api/import/confirm-match (which writes the Transactions row).
// ════════════════════════════════════════════════════════════════════

import { appendRows } from '../../_sheets.js';
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

  // Invoices tab row: B-K = InvNum, Date, Client, Description, Subtotal, HST, Total, HSTFlag, Due, Status
  const invRow = [[
    inv.invNum || '',
    inv.dateVal || '',
    inv.client || '',
    inv.desc || '',
    parseFloat(inv.sub) || 0,
    parseFloat(inv.hstAmt) || 0,
    parseFloat(inv.total) || 0,
    inv.hst || 'Yes',
    inv.dueVal || '',
    'Unpaid',
  ]];

  const invResult = await appendRows(env, userId, `'${INVOICES_TAB}'!B12:K`, invRow);
  if (!invResult.ok) return json({ ok: false, error: 'Invoice write failed: ' + invResult.error });

  return json({ ok: true, invNum: inv.invNum });
}
