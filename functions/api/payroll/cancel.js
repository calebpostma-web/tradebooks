// ════════════════════════════════════════════════════════════════════
// POST /api/payroll/cancel  { sheetRow }
//
// Cancel a committed pay run that turned out wrong — the "delete it and
// start over" Andrea asked for, done the bookkeeping way: nothing is
// erased. The Payroll row's Status becomes 'Cancelled' (every reader
// already skips Cancelled rows: YTD, stubs, T4, remittances) and a
// reversing Transactions row (+net pay, ref …-VOID) is added so the
// books net to zero. Then the run can be redone with Preview → Commit.
//
// Refused for Remitted runs — CRA already has that money; correcting a
// remitted run is an amendment, not a cancel.
// ════════════════════════════════════════════════════════════════════

import { authenticateRequest, json, options } from '../../_shared.js';
import { readRange, writeRange, appendRows } from '../../_sheets.js';
import { PAYROLL_TAB, TXN_TAB, loadPayrollRows, num, round2, resolveWageCategory, SheetReadError, sheetReadErrorBody } from '../../_payroll_sheet.js';

export const onRequestOptions = () => options();

export async function onRequestPost({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  const userId = auth.userId;
  let body = {}; try { body = await request.json(); } catch { /* empty */ }
  const sheetRow = parseInt(body.sheetRow, 10);
  if (!sheetRow || sheetRow < 12) return json({ ok: false, error: 'sheetRow is required' }, 400);

  try {
    const res = await loadPayrollRows(env, userId);
    if (!res.ok) throw new SheetReadError('past pay runs (Payroll tab)', res);
    const run = res.rows.find(r => r.sheetRow === sheetRow);
    if (!run) return json({ ok: false, error: `No pay run at Payroll row ${sheetRow}` }, 404);
    const st = run.status.toLowerCase();
    if (st === 'cancelled') return json({ ok: false, error: 'That pay run is already cancelled.' }, 409);
    if (st === 'remitted') return json({ ok: false, error: 'That pay run has already been remitted to CRA, so it can\'t be cancelled here. Talk to Caleb — it needs a CRA correction.' }, 409);

    // Find the net-pay Transactions row this run wrote (Source 'Payroll',
    // ref PAY-…-<payDate>, party = employee, amount = −net) to copy its account.
    const tx = await readRange(env, userId, `'${TXN_TAB}'!B12:M`);
    if (!tx.ok) throw new SheetReadError('the Transactions tab', tx);
    let account = 'BMO', origRef = '', category = '';
    for (const row of tx.values) {
      if (!row) continue;
      const ref = String(row[9] || '');
      if (String(row[8] || '') !== 'Payroll' || !ref.startsWith('PAY-') || !ref.endsWith(run.payDate)) continue;
      if (String(row[1] || '') !== run.employee) continue;
      if (Math.abs(num(row[3]) + run.netPay) > 0.005) continue;
      account = row[7] || account; origRef = ref; category = row[4] || ''; break;
    }
    const voidRef = (origRef || `PAY-${run.payDate}`) + '-VOID';
    if (tx.values.some(r => r && String(r[9] || '') === voidRef)) {
      return json({ ok: false, error: 'A reversal for this run already exists in Transactions.' }, 409);
    }

    const w = await writeRange(env, userId, `'${PAYROLL_TAB}'!Q${sheetRow}`, [['Cancelled']]);
    if (!w.ok) return json({ ok: false, error: 'Could not mark the run Cancelled: ' + w.error }, 502);

    const reason = String(body.reason || '').slice(0, 120);
    const a = await appendRows(env, userId, `'${TXN_TAB}'!B12:M`, [[
      run.payDate, run.employee,
      `Payroll CANCELLED — reverses ${origRef || 'pay run ' + run.payDate}${reason ? ' (' + reason + ')' : ''}`,
      round2(run.netPay), category || await resolveWageCategory(env, userId), 'No', 0, account, 'Payroll', voidRef, '', 'N/A',
    ]]);
    if (!a.ok) return json({ ok: false, partial: true, error: 'Run marked Cancelled, but the reversing Transactions row failed: ' + a.error }, 502);

    return json({ ok: true, sheetRow, payDate: run.payDate, employee: run.employee, reversed: round2(run.netPay), ref: voidRef });
  } catch (err) {
    if (err instanceof SheetReadError) return json(sheetReadErrorBody(err), 503);
    throw err;
  }
}
