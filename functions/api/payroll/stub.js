// ════════════════════════════════════════════════════════════════════
// GET /api/payroll/stub
//
//   ?list=1                → recent committed pay runs (newest first, ≤ 40)
//   ?row=N                 → stub data (JSON) for Payroll sheet row N
//   ?row=N&format=html     → the rendered stub as a standalone HTML page
//                            (the app opens this in a new tab for Print /
//                            Save as PDF; the same HTML is what gets emailed)
// ════════════════════════════════════════════════════════════════════

import { authenticateRequest, json, options } from '../../_shared.js';
import { loadPayrollRows } from '../../_payroll_sheet.js';
import { buildStubData, renderPayStub } from '../../_paystub.js';

export const onRequestOptions = () => options();

export async function onRequestGet({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  const userId = auth.userId;
  const url = new URL(request.url);

  if (url.searchParams.get('list')) {
    const res = await loadPayrollRows(env, userId);
    if (!res.ok) return json({ ok: false, error: res.error });
    const runs = res.rows
      .filter(r => r.status.toLowerCase() !== 'cancelled')
      .sort((a, b) => b.payDate.localeCompare(a.payDate) || b.sheetRow - a.sheetRow)
      .slice(0, 40)
      .map(r => ({
        sheetRow: r.sheetRow, payDate: r.payDate, employee: r.employee, period: r.period,
        hours: r.hours, gross: r.gross, netPay: r.netPay, status: r.status,
        totalRemittance: r.totalRemittance, stub: r.stub,
      }));
    return json({ ok: true, runs });
  }

  const row = parseInt(url.searchParams.get('row') || '', 10);
  if (!row || row < 12) return json({ ok: false, error: 'row (Payroll sheet row ≥ 12) is required' }, 400);

  const data = await buildStubData(env, userId, row);
  if (!data.ok) return json(data, 404);

  if (url.searchParams.get('format') === 'html') {
    return new Response(renderPayStub(data), {
      headers: { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store' },
    });
  }
  return json(data);
}
