// ════════════════════════════════════════════════════════════════════
// POST /api/intake/paste   { html?: string, text?: string, subject?: string, dryRun?: bool }
//
// Book a pctires.ca order from an email the user pasted into the app.
// Same parser + booking as the email Worker. dryRun returns the parsed
// order without writing anything, so the user can see what would land.
// ════════════════════════════════════════════════════════════════════
import { authenticateRequest, json, options } from '../../_shared.js';
import { parseOrderEmail } from '../../_pctires_order.js';
import { bookOrder, logIntake } from '../../_order_intake.js';

export const onRequestOptions = () => options();

export async function onRequestPost({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  let body; try { body = await request.json(); } catch { return json({ ok: false, error: 'Invalid JSON' }, 400); }

  const parsed = parseOrderEmail({ html: body.html || '', text: body.text || '', subject: body.subject || '' });
  if (!parsed.ok) return json({ ok: false, error: parsed.error }, 400);
  if (body.dryRun) return json({ ok: true, dryRun: true, parsed });

  const result = await bookOrder(env, auth.userId, parsed);
  await logIntake(env, {
    userId: auth.userId, to: 'paste', from: 'paste', subject: body.subject || parsed.orderNo,
    kind: 'order', orderNo: parsed.orderNo,
    status: !result.ok ? 'error' : result.duplicate ? 'duplicate' : 'booked',
    detail: result.ok ? (result.duplicate ? `Already in Invoices (row ${result.invoice.sheetRow})` : `Invoices row ${result.invoice.sheetRow}; ${result.ledger.length} ledger row(s)`) : result.error,
    payload: { orderNo: parsed.orderNo, client: parsed.customer?.name, total: parsed.charged?.total, status: parsed.status, tdg: parsed.tdg?.totalCost },
  });
  return json({ ...result, parsed: { orderNo: parsed.orderNo, client: parsed.customer.name, total: parsed.charged.total, status: parsed.status, tdgCost: parsed.tdg.totalCost } }, result.ok ? 200 : 500);
}
