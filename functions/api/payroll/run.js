// ════════════════════════════════════════════════════════════════════
// POST /api/payroll/run
//
// COMMITS a pay run to the sheet. Same inputs as /calculate plus optional
// idempotency override. Writes:
//   1. One row to 💼 Payroll  (Pay Date, Employee, Age, Business, Work
//      Description, Hours, Rate, Gross, CPP, EI, Fed Tax, ON Tax, Net Pay,
//      YTD Gross, Remittance Due, Status)
//   2. One row to 📒 Transactions  (-Net Pay, category Wages & Salaries,
//      account = BMO, ref = PAY-<empShort>-<payDate>)
//
// Idempotency: we reject if a Payroll row for (employee, payDate) already
// exists, unless body.overwrite === true.
//
// Status convention:
//   - Deductions == 0 (typical for under-18): Status = 'Paid',   Remittance Due = ''
//   - Deductions  > 0:                        Status = 'Paid',   Remittance Due = <15th of next month>
//     (becomes 'Remitted' when /api/payroll/remit fires in Stage 4)
// ════════════════════════════════════════════════════════════════════

import { readRange, appendRows } from '../../_sheets.js';
import { authenticateRequest, json, options } from '../../_shared.js';
import { calculatePayRun, remittanceDueDate } from '../../_payroll.js';

const WORK_LOG_TAB = '📝 Work Log';
const PAYROLL_TAB = '💼 Payroll';
const TXN_TAB = '📒 Transactions';
const WAGE_CATEGORY = 'Wages & Salaries';

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
  const overwrite = !!body.overwrite;
  const account = body.account || 'BMO';

  if (!employeeId || !periodStart || !periodEnd || !payDate) {
    return json({ ok: false, error: 'employeeId, periodStart, periodEnd, payDate are required' }, 400);
  }

  const employee = await loadEmployee(env, userId, employeeId);
  if (!employee) return json({ ok: false, error: `Employee ${employeeId} not found` }, 404);

  // Idempotency check — existing Payroll row for this (employee, payDate)?
  const existingRow = await findExistingPayrollRow(env, userId, employee.name, payDate);
  if (existingRow && !overwrite) {
    return json({
      ok: false,
      error: `A pay run for ${employee.name} on ${payDate} already exists (row ${existingRow.sheetRow}). Pass overwrite:true to commit anyway.`,
      existingRow,
    }, 409);
  }

  // Pull work log + YTD — same as calculate.
  const wlEntries = await loadWorkLogInRange(env, userId, employee.name, periodStart, periodEnd);
  const wlGross = wlEntries.reduce((s, e) => s + (e.hours * e.rate), 0);
  const grossPay = Math.round((wlGross + adjustment) * 100) / 100;
  const totalHours = wlEntries.reduce((s, e) => s + e.hours, 0);

  if (grossPay <= 0) {
    return json({ ok: false, error: 'Gross pay is $0 — no Work Log entries in range and no adjustment.' }, 400);
  }

  const ytd = await loadYtdState(env, userId, employee.name, payDate);

  const result = calculatePayRun({ employee, payDate, grossPay, ytd });

  // Determine remittance due + status
  const totalDeductions = result.cpp + result.cpp2 + result.fedTax + result.onTax;
  const remitDue = totalDeductions > 0 ? remittanceDueDate(payDate) : '';
  const status = 'Paid';  // employee paid; CRA remittance tracked separately via Status flip

  // Build Payroll row (16 data columns B-Q)
  const businesses = [...new Set(wlEntries.map(e => e.business).filter(Boolean))];
  const businessDisplay = businesses.length === 0 ? 'Postma' : businesses.length === 1 ? businesses[0] : 'Multi';
  const workDescription = buildWorkDescription(wlEntries, adjustment);
  const avgRate = totalHours > 0 ? round2(wlGross / totalHours) : '';

  const payrollRow = [[
    payDate,                      // B Pay Date
    employee.name,                // C Employee
    result.age ?? '',             // D Age
    businessDisplay,              // E Business
    workDescription,              // F Work Description
    totalHours || '',             // G Hours
    avgRate,                      // H Rate
    result.gross,                 // I Gross
    round2(result.cpp + result.cpp2), // J CPP (combined base+CPP2)
    result.ei,                    // K EI
    result.fedTax,                // L Fed Tax
    result.onTax,                 // M ON Tax
    result.netPay,                // N Net Pay
    round2(ytd.gross + result.gross),  // O YTD Gross
    remitDue,                     // P Remittance Due
    status,                       // Q Status
  ]];

  const payResult = await appendRows(env, userId, `'${PAYROLL_TAB}'!B12:Q`, payrollRow);
  if (!payResult.ok) return json({ ok: false, error: 'Payroll write failed: ' + payResult.error });

  // Build Transactions row for the NET pay that left BMO.
  // Deductions remain as an implicit liability until /api/payroll/remit
  // writes the CRA remittance Transaction(s) separately.
  const empShort = (employee.id || 'emp').slice(0, 6);
  const ref = `PAY-${empShort}-${payDate}`;
  const txnRow = [[
    payDate,                      // B Date
    employee.name,                // C Party
    `Payroll — ${workDescription}`, // D Description
    -result.netPay,               // E Amount (signed; negative = money out)
    WAGE_CATEGORY,                // F Category
    'No',                         // G HST?
    0,                            // H HST Amount
    account,                      // I Account
    'Payroll',                    // J Source
    ref,                          // K Ref
    '',                           // L Related Invoice (N/A)
    'N/A',                        // M Match Status
  ]];
  const txnResult = await appendRows(env, userId, `'${TXN_TAB}'!B12:M`, txnRow);
  if (!txnResult.ok) {
    // Payroll row already landed — return partial success with the error so
    // the user knows what happened. Manual cleanup if they want to roll back.
    return json({
      ok: false,
      error: 'Payroll row written, but Transactions write failed: ' + txnResult.error,
      partial: true, payrollUpdate: payResult.updates,
    });
  }

  return json({
    ok: true,
    employee: { id: employee.id, name: employee.name },
    payDate, period: { start: periodStart, end: periodEnd },
    gross: result.gross, netPay: result.netPay,
    deductions: {
      cpp: result.cpp, cpp2: result.cpp2, ei: result.ei,
      fedTax: result.fedTax, onTax: result.onTax, total: round2(totalDeductions),
    },
    remittanceDue: remitDue || null,
    status,
    workLogEntries: wlEntries.length,
    payrollRow: payResult.updates?.updatedRange,
    transactionRow: txnResult.updates?.updatedRange,
    ytdAfter: {
      gross: round2(ytd.gross + result.gross),
      cpp: round2(ytd.cppBase + result.cpp + result.cpp2),
      fedTax: round2(ytd.fedTax + result.fedTax),
      onTax: round2(ytd.onTax + result.onTax),
    },
  });
}

// ─── Helpers ─────────────────────────────────────────────────────────

function round2(n) { return Math.round((Number(n) || 0) * 100) / 100; }

function buildWorkDescription(entries, adjustment) {
  if (!entries.length) {
    return adjustment > 0 ? `Adjustment / bonus only` : `—`;
  }
  const tasks = [...new Set(entries.map(e => e.task).filter(Boolean))];
  let desc;
  if (tasks.length === 1) desc = tasks[0];
  else if (tasks.length <= 3) desc = tasks.join('; ');
  else desc = `${tasks.slice(0, 3).join('; ')}; +${tasks.length - 3} more`;
  if (adjustment > 0) desc += ` (+ adjustment)`;
  return desc;
}

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
      date, business: business || '', task: task || '',
      hours: parseFloat(hours) || 0, rate: parseFloat(rate) || 0,
      notes: notes || '', audit: audit || '',
    });
  }
  return entries;
}

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
    if (rowT >= payT) continue;

    ytd.gross   += parseFloat(gross)   || 0;
    ytd.cppBase += parseFloat(cpp)     || 0;
    ytd.fedTax  += parseFloat(fedTax)  || 0;
    ytd.onTax   += parseFloat(onTax)   || 0;
  }

  ytd.gross   = round2(ytd.gross);
  ytd.cppBase = round2(ytd.cppBase);
  ytd.fedTax  = round2(ytd.fedTax);
  ytd.onTax   = round2(ytd.onTax);
  return ytd;
}

async function findExistingPayrollRow(env, userId, employeeName, payDateIso) {
  const result = await readRange(env, userId, `'${PAYROLL_TAB}'!B12:Q500`);
  if (!result.ok) return null;
  for (let i = 0; i < result.values.length; i++) {
    const row = result.values[i];
    if (!row || !row[0]) continue;
    const [rowDate, rowName] = row;
    if (rowName !== employeeName) continue;
    const rowT = Date.parse(rowDate);
    if (isNaN(rowT)) continue;
    if (new Date(rowDate).toISOString().slice(0, 10) === payDateIso) {
      return { sheetRow: 12 + i, rowValues: row };
    }
  }
  return null;
}
