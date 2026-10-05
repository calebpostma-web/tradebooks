// ════════════════════════════════════════════════════════════════════
// POST /api/payroll/calculate
//
// Preview a pay run for a given employee and period. No writes. Used to
// show the user exactly what gets committed before they hit Run.
//
// Request body:
//   {
//     employeeId:   string,  // employee from D1 profile
//     periodStart:  'YYYY-MM-DD',
//     periodEnd:    'YYYY-MM-DD',
//     payDate:      'YYYY-MM-DD',  // used for age, YTD calendar year, remittance due
//     adjustment?:  number,  // optional extra gross (bonus, advance, overtime)
//   }
//
// Pipeline:
//   1. Load employee from D1
//   2. Read Work Log entries where Employee==name AND date in [periodStart, periodEnd]
//   3. Sum gross = Σ(hours × rate) + adjustment
//   4. Read existing Payroll rows for this employee in the calendar year of payDate;
//      aggregate YTD gross / CPP / fed tax / ON tax BEFORE this run
//   5. Run calculatePayRun from _payroll.js
//   6. Return preview: gross, deductions, net, breakdown, workLogEntries
//
// ════════════════════════════════════════════════════════════════════

import { authenticateRequest, json, options } from '../../_shared.js';
import { calculatePayRun, remittanceDueDate } from '../../_payroll.js';
import { loadEmployee, loadWorkLogInRange, loadYtdState, loadRemitter } from '../../_payroll_sheet.js';

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

  const { employeeId, periodStart, periodEnd, payDate } = body;
  const adjustment = parseFloat(body.adjustment) || 0;

  if (!employeeId || !periodStart || !periodEnd || !payDate) {
    return json({ ok: false, error: 'employeeId, periodStart, periodEnd, payDate are required' }, 400);
  }

  // 1. Load employee
  const employee = await loadEmployee(env, userId, employeeId);
  if (!employee) return json({ ok: false, error: `Employee ${employeeId} not found` }, 404);

  // 2. Read Work Log entries in range for this employee
  const wlEntries = await loadWorkLogInRange(env, userId, employee.name, periodStart, periodEnd);
  const wlGross = wlEntries.reduce((s, e) => s + (e.hours * e.rate), 0);
  const grossPay = Math.round((wlGross + adjustment) * 100) / 100;

  // 3. Read Payroll rows for YTD state (calendar year of payDate)
  const ytd = await loadYtdState(env, userId, employee.name, payDate);
  const remitter = await loadRemitter(env, userId);

  // 4. Run the engine
  const result = calculatePayRun({
    employee,
    payDate,
    grossPay,
    ytd,
  });

  return json({
    ok: true,
    employee: {
      id: employee.id,
      name: employee.name,
      dob: employee.dob,
      relationship: employee.relationship,
      familyEiExempt: employee.familyEiExempt,
    },
    period: { start: periodStart, end: periodEnd, payDate },
    workLog: {
      entries: wlEntries,
      sumGross: Math.round(wlGross * 100) / 100,
      adjustment,
    },
    ytdBefore: ytd,
    ytdAfter: {
      gross: Math.round((ytd.gross + result.gross) * 100) / 100,
      cppBase: Math.round((ytd.cppBase + result.cpp + result.cpp2) * 100) / 100,
      fedTax: Math.round((ytd.fedTax + result.fedTax) * 100) / 100,
      onTax: Math.round((ytd.onTax + result.onTax) * 100) / 100,
    },
    calculation: result,
    remitter,
    remittanceDue: (result.cpp + result.cpp2 + result.ei + result.fedTax + result.onTax) > 0
      ? remittanceDueDate(payDate, remitter)
      : null,
  });
}
