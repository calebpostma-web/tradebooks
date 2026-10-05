// ════════════════════════════════════════════════════════════════════
// POST /api/payroll/remit
//
// Marks a set of Payroll rows as Remitted and writes a single matching
// negative Transaction for the CRA source deduction payment.
//
// Request body:
//   {
//     payrollRows: number[],     // sheet rows in 💼 Payroll (1-indexed) to flip Paid → Remitted
//     remitDate:   'YYYY-MM-DD', // date the remittance actually hit CRA
//     account:     'BMO' | 'Cash' | ...,  // defaults 'BMO'
//     notes?:      string,
//   }
//
// Server re-reads the Payroll rows (don't trust client-supplied totals),
// verifies each is currently Paid with deductions owed, sums CPP+Fed+ON,
// writes ONE Transactions row (ref CRA-REMIT-YYYYMMDD), then flips each
// Payroll row's Status column from 'Paid' to 'Remitted'.
// ════════════════════════════════════════════════════════════════════

import { writeRange, appendRows } from '../../_sheets.js';
import { authenticateRequest, json, options } from '../../_shared.js';
import { PAYROLL_TAB, TXN_TAB, round2, loadPayrollRows, resolveWageCategory } from '../../_payroll_sheet.js';

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

  const payrollRows = Array.isArray(body.payrollRows) ? body.payrollRows.map(n => parseInt(n, 10)).filter(n => n >= 12) : [];
  const remitDate = body.remitDate;
  const account = body.account || 'BMO';
  const notes = body.notes || '';

  if (!payrollRows.length) return json({ ok: false, error: 'payrollRows is required and must be non-empty' }, 400);
  if (!remitDate) return json({ ok: false, error: 'remitDate is required' }, 400);

  // Read Payroll tab and verify each row is Paid with a deduction total.
  const res = await loadPayrollRows(env, userId);
  if (!res.ok) return json({ ok: false, error: 'Failed to read Payroll: ' + res.error });
  const bySheetRow = new Map(res.rows.map(r => [r.sheetRow, r]));
  const WAGE_CATEGORY = await resolveWageCategory(env, userId);

  const pickedRows = [];
  const issues = [];
  for (const rowNum of payrollRows) {
    const r = bySheetRow.get(rowNum);
    if (!r) { issues.push(`Row ${rowNum} not found`); continue; }
    const { payDate, employee, status } = r;
    if (status.toLowerCase() === 'remitted') {
      issues.push(`Row ${rowNum} (${employee} ${payDate}) is already Remitted`);
      continue;
    }
    if (status.toLowerCase() !== 'paid') {
      issues.push(`Row ${rowNum} (${employee} ${payDate}) has Status '${status}' — only 'Paid' rows can be remitted`);
      continue;
    }
    const total = round2(r.cpp + r.fedTax + r.onTax);
    if (total <= 0) {
      issues.push(`Row ${rowNum} (${employee} ${payDate}) has $0 deductions — nothing to remit`);
      continue;
    }
    pickedRows.push({
      sheetRow: rowNum, payDate, employee, cpp: r.cpp, fedTax: r.fedTax, onTax: r.onTax, total,
    });
  }

  if (!pickedRows.length) {
    return json({ ok: false, error: 'No valid rows to remit', issues }, 400);
  }

  const totalCpp = round2(pickedRows.reduce((s, r) => s + r.cpp, 0));
  const totalFed = round2(pickedRows.reduce((s, r) => s + r.fedTax, 0));
  const totalOn  = round2(pickedRows.reduce((s, r) => s + r.onTax, 0));
  const totalAmount = round2(totalCpp + totalFed + totalOn);

  // Build remittance description: list distinct pay months from the picked rows.
  const months = [...new Set(pickedRows.map(r => monthLabel(r.payDate)))].sort();
  const payRunCount = pickedRows.length;
  const description = `CRA source deductions — ${months.join(', ')} (${payRunCount} pay run${payRunCount === 1 ? '' : 's'})`
    + (notes ? ` · ${notes}` : '');
  const ref = `CRA-REMIT-${remitDate.replace(/-/g, '')}`;

  // Write the Transaction row (negative; money out to CRA)
  const txnRow = [[
    remitDate,                    // B Date
    'CRA',                        // C Party
    description,                  // D Description
    -totalAmount,                 // E Amount (signed)
    WAGE_CATEGORY,                // F Category (total wage cost = net + remittance)
    'No',                         // G HST?
    0,                            // H HST Amount
    account,                      // I Account
    'Payroll',                    // J Source
    ref,                          // K Ref
    '',                           // L Related Invoice
    'N/A',                        // M Match Status
  ]];
  const txnResult = await appendRows(env, userId, `'${TXN_TAB}'!B12:M`, txnRow);
  if (!txnResult.ok) return json({ ok: false, error: 'Transaction write failed: ' + txnResult.error });

  // Flip Payroll Status column (col Q = 17th col, row-indexed) on each picked row.
  // Use individual writeRange calls; Google Sheets batchUpdate for a list of
  // non-contiguous single-cell writes is overkill here given the row counts.
  const flipErrors = [];
  for (const r of pickedRows) {
    const statusRange = `'${PAYROLL_TAB}'!Q${r.sheetRow}`;
    const res = await writeRange(env, userId, statusRange, [['Remitted']]);
    if (!res.ok) flipErrors.push(`Row ${r.sheetRow}: ${res.error}`);
  }

  return json({
    ok: true,
    remitDate, account, ref, description,
    totals: { cpp: totalCpp, fedTax: totalFed, onTax: totalOn, total: totalAmount },
    rowsMarkedRemitted: pickedRows.map(r => r.sheetRow),
    payRunCount,
    months,
    transactionRange: txnResult.updates?.updatedRange,
    issues,             // non-fatal — rows that were skipped
    flipErrors,         // non-fatal — Transaction was written; these rows may still show 'Paid' and need manual fix
  });
}

// ── Helpers ──

function monthLabel(iso) {
  try {
    const d = new Date(iso);
    return d.toLocaleString('en-CA', { month: 'short', year: 'numeric' });
  } catch {
    return iso;
  }
}
