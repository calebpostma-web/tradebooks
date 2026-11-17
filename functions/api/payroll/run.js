// ════════════════════════════════════════════════════════════════════
// POST /api/payroll/run
//
// COMMITS a pay run to the sheet. Same inputs as /calculate plus optional
// idempotency override. Writes:
//   1. One row to 💼 Payroll  (Pay Date, Employee, Age, Business, Work
//      Description, Hours, Rate, Gross, CPP, EI, Fed Tax, ON Tax, Net Pay,
//      YTD Gross, Remittance Due, Status, Period, Employer CPP, Employer EI, Stub)
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

import { appendRows } from '../../_sheets.js';
import { authenticateRequest, json, options } from '../../_shared.js';
import { calculatePayRun, remittanceDueDate, applyCppRecovery } from '../../_payroll.js';
import {
  PAYROLL_TAB, TXN_TAB, round2,
  loadEmployee, loadWorkLogInRange, loadYtdState, findExistingPayrollRow, resolveWageCategory, loadRemitter,
  loadPeriodChecks,
} from '../../_payroll_sheet.js';
import { SheetReadError, sheetReadErrorBody } from '../../_payroll_sheet.js';

export const onRequestOptions = () => options();

export async function onRequestPost(ctx) {
  try {
    return await handlePost(ctx);
  } catch (err) {
    if (err instanceof SheetReadError) return json(sheetReadErrorBody(err), 503);
    throw err;
  }
}

async function handlePost({ request, env }) {
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
  const WAGE_CATEGORY = await resolveWageCategory(env, userId);

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

  // No hours, only an Adjustment: allowed for a real bonus or extra pay, but
  // the screen has to say so. Otherwise it's almost always hours dated
  // outside the period (Andrea's 300 h dated on the pay day).
  const isBonus = !!body.isBonus;
  if (wlEntries.length === 0 && !isBonus) {
    return json({ ok: false, needsBonusConfirm: true,
      error: 'There are no Work Log hours in this period, only the Adjustment amount. If these are hours worked, log them in the Work Log dated inside this period (the last day the work was done) and Preview again. If this really is a bonus or extra pay with no hours, tick "This is a bonus / extra pay" and commit again.' }, 400);
  }

  // Never pay the same hours twice.
  const checks = await loadPeriodChecks(env, userId, employee.name, periodStart, periodEnd, wlEntries,
    existingRow && overwrite ? existingRow.sheetRow : null);
  if (checks.alreadyPaid.length) {
    const p = checks.alreadyPaid[0].paidBy;
    return json({ ok: false, alreadyPaid: checks.alreadyPaid,
      error: `${checks.alreadyPaid.length} Work Log entr${checks.alreadyPaid.length === 1 ? 'y' : 'ies'} in this period ` +
        `${checks.alreadyPaid.length === 1 ? 'was' : 'were'} already paid by the ${fmtLong(p.payDate)} pay (${fmtLong(p.start)} to ${fmtLong(p.end)}). ` +
        `Start this period ${fmtLong(new Date(Date.parse(p.end) + 86400000).toISOString().slice(0, 10))} or later, or cancel that earlier run first if it was wrong.` }, 409);
  }

  const ytd = await loadYtdState(env, userId, employee.name, payDate);
  const remitter = await loadRemitter(env, userId);

  const result = calculatePayRun({ employee, payDate, grossPay, ytd: { ...ytd, cppBase: ytd.cppRequired }, periodStart, periodEnd });
  applyCppRecovery(result, ytd.cppShortfall);

  // Determine remittance due + status
  const totalDeductions = result.totalDeductions;            // includes any CPP recovery
  const remitDue = result.totalRemittance > 0 ? remittanceDueDate(payDate, remitter) : '';
  const status = 'Paid';  // employee paid; CRA remittance tracked separately via Status flip

  // Build Payroll row (20 data columns B-U)
  const businesses = [...new Set(wlEntries.map(e => e.business).filter(Boolean))];
  const businessDisplay = businesses.length === 0 ? 'Postma' : businesses.length === 1 ? businesses[0] : 'Multi';
  const workDescription = buildWorkDescription(wlEntries, adjustment, isBonus)
    + (result.cppRecovery > 0 ? ` · CPP incl. $${result.cppRecovery.toFixed(2)} correction (earlier under-deduction)` : '');
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
    round2(result.cpp + result.cpp2 + result.cppRecovery), // J CPP deducted (base+CPP2+any recovery)
    result.ei,                    // K EI
    result.fedTax,                // L Fed Tax
    result.onTax,                 // M ON Tax
    result.netPay,                // N Net Pay
    round2(ytd.gross + result.gross),  // O YTD Gross
    remitDue,                     // P Remittance Due
    status,                       // Q Status
    `${periodStart} → ${periodEnd}`, // R Period
    result.employerCpp,           // S Employer CPP = the CORRECT contribution for this pay (no recovery)
    result.employerEi,            // T Employer EI (1.4× employee EI)
    '',                           // U Stub (Drive link, filled when a stub is generated)
  ]];

  let payResult = await appendRows(env, userId, `'${PAYROLL_TAB}'!B12:U`, payrollRow);
  if (!payResult.ok && /grid|range|column/i.test(payResult.error || '')) {
    // Sheet not yet migrated to the 21-column layout (user hasn't clicked
    // "Update sheet"). Write the classic 16 columns so the pay still lands;
    // readers fall back to the CRA identities for the employer share.
    payResult = await appendRows(env, userId, `'${PAYROLL_TAB}'!B12:Q`, [payrollRow[0].slice(0, 16)]);
  }
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
      cpp: result.cpp, cpp2: result.cpp2, cppRecovery: result.cppRecovery, ei: result.ei,
      fedTax: result.fedTax, onTax: result.onTax, total: round2(totalDeductions),
    },
    employer: {
      cpp: result.employerCpp, ei: result.employerEi, share: result.employerShare,
      totalRemittance: result.totalRemittance, employerCost: result.employerCost,
    },
    remittanceDue: remitDue || null,
    remitter,
    status,
    workLogEntries: wlEntries.length,
    payrollRow: payResult.updates?.updatedRange,
    transactionRow: txnResult.updates?.updatedRange,
    ytdAfter: {
      gross: round2(ytd.gross + result.gross),
      cpp: round2(ytd.cppBase + result.cpp + result.cpp2 + result.cppRecovery),
      fedTax: round2(ytd.fedTax + result.fedTax),
      onTax: round2(ytd.onTax + result.onTax),
    },
  });
}

// ─── Helpers ─────────────────────────────────────────────────────────

function fmtLong(iso) {
  const d = new Date(iso + 'T00:00:00Z');
  return isNaN(d) ? iso : d.toLocaleDateString('en-CA', { month: 'short', day: 'numeric', year: 'numeric', timeZone: 'UTC' });
}

function buildWorkDescription(entries, adjustment, isBonus = false) {
  if (!entries.length) {
    return adjustment > 0 ? (isBonus ? `Bonus / extra pay — no hours` : `Adjustment only — no hours logged`) : `—`;
  }
  const tasks = [...new Set(entries.map(e => e.task).filter(Boolean))];
  let desc;
  if (tasks.length === 1) desc = tasks[0];
  else if (tasks.length <= 3) desc = tasks.join('; ');
  else desc = `${tasks.slice(0, 3).join('; ')}; +${tasks.length - 3} more`;
  if (adjustment > 0) desc += ` (+ adjustment)`;
  return desc;
}
