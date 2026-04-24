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

import { readRange } from '../../_sheets.js';
import { authenticateRequest, json, options } from '../../_shared.js';
import { calculatePayRun, remittanceDueDate } from '../../_payroll.js';

const WORK_LOG_TAB = '📝 Work Log';
const PAYROLL_TAB = '💼 Payroll';

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
    remittanceDue: (result.cpp + result.cpp2 + result.fedTax + result.onTax) > 0
      ? remittanceDueDate(payDate)
      : null,
  });
}

// ─── Helpers ─────────────────────────────────────────────────────────

async function loadEmployee(env, userId, employeeId) {
  try {
    const row = await env.DB.prepare('SELECT employees FROM profiles WHERE user_id = ?')
      .bind(userId).first();
    if (!row || !row.employees) return null;
    const list = JSON.parse(row.employees);
    if (!Array.isArray(list)) return null;
    return list.find(e => e.id === employeeId) || null;
  } catch {
    return null;
  }
}

/**
 * Load Work Log rows for one employee where Date ∈ [start, end].
 * Work Log layout B-I: Date | Employee | Business | Task | Hours | Rate | Notes | Entry Audit
 */
async function loadWorkLogInRange(env, userId, employeeName, startIso, endIso) {
  const result = await readRange(env, userId, `'${WORK_LOG_TAB}'!B12:I1000`);
  if (!result.ok) return [];
  const startT = Date.parse(startIso);
  const endT = Date.parse(endIso);
  const entries = [];
  for (const row of result.values) {
    if (!row || !row[0]) continue;
    const [date, name, business, task, hours, rate, notes, audit] = row;
    if (name !== employeeName) continue;
    const t = Date.parse(date);
    if (isNaN(t) || t < startT || t > endT) continue;
    entries.push({
      date,
      business: business || '',
      task: task || '',
      hours: parseFloat(hours) || 0,
      rate: parseFloat(rate) || 0,
      notes: notes || '',
      audit: audit || '',
    });
  }
  return entries;
}

/**
 * Sum existing Payroll rows for this employee where Pay Date is in the
 * calendar year of payDateIso AND strictly before payDateIso.
 *
 * Payroll layout B-Q (indices 0–15 in the row array):
 *  0 Pay Date, 1 Employee, 2 Age, 3 Business, 4 Work Description,
 *  5 Hours, 6 Rate, 7 Gross, 8 CPP, 9 EI, 10 Fed Tax, 11 ON Tax,
 *  12 Net Pay, 13 YTD Gross, 14 Remittance Due, 15 Status
 */
async function loadYtdState(env, userId, employeeName, payDateIso) {
  const ytd = { gross: 0, cppBase: 0, cpp2: 0, fedTax: 0, onTax: 0 };
  const result = await readRange(env, userId, `'${PAYROLL_TAB}'!B12:Q500`);
  if (!result.ok) return ytd;

  const payT = Date.parse(payDateIso);
  const payYear = new Date(payDateIso).getUTCFullYear();

  for (const row of result.values) {
    if (!row || !row[0]) continue;
    const [rowDate, rowName, , , , , , gross, cpp, , fedTax, onTax, , , , status] = row;
    if (rowName !== employeeName) continue;
    if (status && String(status).toLowerCase() === 'cancelled') continue;
    const rowT = Date.parse(rowDate);
    if (isNaN(rowT)) continue;
    if (new Date(rowDate).getUTCFullYear() !== payYear) continue;
    if (rowT >= payT) continue; // only rows BEFORE the current run

    ytd.gross   += parseFloat(gross)   || 0;
    ytd.cppBase += parseFloat(cpp)     || 0;  // combined CPP base + CPP2 in one column
    ytd.fedTax  += parseFloat(fedTax)  || 0;
    ytd.onTax   += parseFloat(onTax)   || 0;
  }

  // Round the aggregates
  ytd.gross   = Math.round(ytd.gross   * 100) / 100;
  ytd.cppBase = Math.round(ytd.cppBase * 100) / 100;
  ytd.fedTax  = Math.round(ytd.fedTax  * 100) / 100;
  ytd.onTax   = Math.round(ytd.onTax   * 100) / 100;
  return ytd;
}
