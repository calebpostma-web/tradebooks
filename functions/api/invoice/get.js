// ════════════════════════════════════════════════════════════════════
// GET /api/invoice/get?num=PCT-10377   (or ?row=45)
// One invoice with its lines and payments, so the wizard can reopen it.
// ════════════════════════════════════════════════════════════════════
import { authenticateRequest, json, options } from '../../_shared.js';
import { findInvoice } from '../../_invoices.js';

export const onRequestOptions = () => options();

export async function onRequestGet({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  const url = new URL(request.url);
  const invNum = (url.searchParams.get('num') || '').trim();
  const sheetRow = parseInt(url.searchParams.get('row') || '', 10) || 0;
  if (!invNum && !sheetRow) return json({ ok: false, error: 'num or row is required' }, 400);
  const res = await findInvoice(env, auth.userId, { invNum, sheetRow });
  if (!res.ok) return json({ ok: false, error: res.error });
  if (!res.invoice) return json({ ok: false, error: `Invoice ${invNum || '#' + sheetRow} not found` }, 404);
  return json({ ok: true, invoice: res.invoice });
}
