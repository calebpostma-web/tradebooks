// ════════════════════════════════════════════════════════════════════
// GET /api/invoice/list
//
// Every invoice number in the 🧾 Invoices tab, plus the next free
// sequential number. Used by the Invoice Wizard to (a) offer a real
// "next number" instead of a random one and (b) warn before a duplicate.
//
// Invoice numbers are free text (e.g. "1042", "C-1042", "PCT-10377").
// The sequential counter looks at the LAST run of digits in each number,
// so "C-1042" still advances the counter to 1043. Numbers with no digits
// are listed but don't affect the counter.
//
// Response: { ok, invoices:[{invNum, date, client, total, status, sheetRow}],
//             numbers:[...], nextNumber:"1043" }
// ════════════════════════════════════════════════════════════════════

import { authenticateRequest, json, options } from '../../_shared.js';
import { loadInvoices } from '../../_invoices.js';

export const onRequestOptions = () => options();

export async function onRequestGet({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  const userId = auth.userId;
  const url = new URL(request.url);
  const start = parseInt(url.searchParams.get('start') || '', 10) || 1001;

  const res = await loadInvoices(env, userId);
  if (!res.ok) return json({ ok: false, error: res.error });

  const invoices = [];
  let maxSeq = start - 1;
  for (const inv of res.invoices) {
    invoices.push({
      invNum: inv.invNum, date: inv.date, client: inv.client, total: inv.total,
      status: inv.status, balanceDue: inv.balanceDue, paidSoFar: inv.paidSoFar,
      sheetRow: inv.sheetRow, hasLines: !!inv.lines,
    });
    const m = /(\d+)(?!.*\d)/.exec(inv.invNum);          // last run of digits
    if (m) { const n = parseInt(m[1], 10); if (n > maxSeq && n < 1e9) maxSeq = n; }
  }
  invoices.sort((a, b) => b.sheetRow - a.sheetRow);          // newest first

  return json({
    ok: true,
    invoices,
    numbers: invoices.map(i => i.invNum),
    nextNumber: String(maxSeq + 1),
  });
}
