// ════════════════════════════════════════════════════════════════════
// POST /api/invoice/create
// Create an invoice row and the corresponding accrual-basis income row.
// Replaces Apps Script's handleInvoice().
// ════════════════════════════════════════════════════════════════════

import { appendRows, writeRange, generateRef } from '../../_sheets.js';
import { authenticateRequest, json, options } from '../../_shared.js';

const INVOICES_TAB = '🧾 Invoices';
const INCOME_TAB = '💰 Income';

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

  // Income tab row (accrual basis: income is recognized at issue, not payment)
  // B-G = Date, Client, Description, InvNum, Amount, Category
  const ref = generateRef('INV', inv.dateVal, inv.sub, inv.client);
    const incRow = [[
          inv.dateVal || '',
          inv.client || '',
          inv.desc || '',
          inv.invNum || '',
          parseFloat(inv.sub) || 0,
          inv.category || 'Consulting Revenue',
        ]];

  const incResult = await appendRows(env, userId, `'${INCOME_TAB}'!B12:G`, incRow);
    if (!incResult.ok) return json({ ok: false, error: 'Income write failed: ' + incResult.error });

  // Metadata columns on income row: J=Source, K=Ref
  const match = /!B(\d+):G(\d+)/.exec(incResult.updates.updatedRange);
    if (match) {
          const row = parseInt(match[1]);
          const metaRange = `'${INCOME_TAB}'!J${row}:K${row}`;
          const metaResult = await writeRange(env, userId, metaRange, [['Invoice', ref]]);
          if (!metaResult.ok) return json({ ok: false, error: 'Income meta write failed: ' + metaResult.error });
    }

  return json({ ok: true, invNum: inv.invNum });
}
