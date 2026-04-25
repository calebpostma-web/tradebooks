// ════════════════════════════════════════════════════════════════════
// GET /api/yearend/checklist?fy=FY2026&fye=March+31
//
// Computes year-end readiness for a given fiscal year. Returns a structured
// list of items with status (green / amber / red) so the front-end can
// render a checklist with "fix this" hints.
//
// The work is intentionally batched into parallel reads — checklist needs
// to feel snappy in the UI even when poking 5+ different sheet ranges.
//
// Items returned (each with key, label, status, detail, optional action):
//   - transactions:    rows in window, % categorized
//   - invoices:        outstanding count + total
//   - hst_returns:     # of quarterly HST remittances logged (target 4)
//   - cra_payroll:     # of payroll source-deduction remittances logged
//   - cra_corp_tax:    # of corp tax instalments + final logged
//   - statements_bmo:  per-month upload coverage (target 12)
//   - statements_amex: per-month upload coverage (target 12)
//   - payroll:         # of pay runs + whether T4s likely needed
//   - receipts:        count of receipts uploaded (rough)
// ════════════════════════════════════════════════════════════════════

import { getGoogleAccessToken } from '../../_google.js';
import { readRange, getSpreadsheetMetadata, resolveTabName } from '../../_sheets.js';
import { authenticateRequest, json, options, fiscalYearOf, fiscalYearMonths } from '../../_shared.js';
import { findOrCreateFolder, listFolderFiles } from '../../_drive.js';

// Logical tab names — these get resolved against the user's actual sheet
// metadata so we work with both modern (📒 Transactions) and legacy
// (🧾 Transactions, no-emoji, etc) tab namings.
const TAB_TXN = 'Transactions';
const TAB_INV = 'Invoices';
const TAB_REM = 'CRA Remittances';
const TAB_PAY = 'Payroll';
const RECEIPTS_PARENT = 'AI Bookkeeper Receipts';
const YEAREND_PARENT = 'AI Bookkeeper Year-End';

export const onRequestOptions = () => options();

export async function onRequest({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  const userId = auth.userId;

  const url = new URL(request.url);
  const fyParam = url.searchParams.get('fy');
  const fye = url.searchParams.get('fye') || 'December 31';

  const fy = resolveFiscalYear(fyParam, fye);
  const tok = await getGoogleAccessToken(env, userId);
  if (!tok.ok) return json({ ok: false, error: tok.error || 'Drive token unavailable' }, 401);

  // Resolve actual tab names from sheet metadata (handles legacy emoji prefixes
  // and freshly-created sheets uniformly). Bail with a helpful error if the
  // sheet is unreachable so we don't fire 4 broken reads.
  const meta = await getSpreadsheetMetadata(env, userId);
  if (!meta.ok) return json({ ok: false, error: meta.error || 'Could not read sheet metadata' });
  const tabTxn = resolveTabName(meta.sheets, TAB_TXN);
  const tabInv = resolveTabName(meta.sheets, TAB_INV);
  const tabRem = resolveTabName(meta.sheets, TAB_REM);
  const tabPay = resolveTabName(meta.sheets, TAB_PAY);

  // Run sheet reads + Drive list in parallel. Each branch handles its own
  // errors — we don't want a single Drive hiccup to kill the whole checklist.
  // For tabs that don't exist we synthesize a "no data" result rather than
  // hitting the API with a known-bad range.
  const [txnRows, invRows, remRows, payRows, statementFiles, receiptFiles] = await Promise.all([
    tabTxn ? safeReadRange(env, userId, `'${tabTxn}'!B12:M1000`) : missingTabResult(TAB_TXN),
    tabInv ? safeReadRange(env, userId, `'${tabInv}'!B12:Q500`)  : missingTabResult(TAB_INV),
    tabRem ? safeReadRange(env, userId, `'${tabRem}'!B12:J500`)  : missingTabResult(TAB_REM),
    tabPay ? safeReadRange(env, userId, `'${tabPay}'!B12:Q500`)  : missingTabResult(TAB_PAY),
    listStatementFiles(tok.accessToken, fy),
    listReceiptFiles(tok.accessToken, fy),
  ]);

  const items = [
    checkTransactions(txnRows, fy),
    checkInvoices(invRows, fy),
    checkHstReturns(remRows, fy),
    checkCraPayroll(remRows, fy),
    checkCraCorpTax(remRows, fy),
    checkStatements('BMO', statementFiles, fy),
    checkStatements('AMEX', statementFiles, fy),
    checkPayroll(payRows, fy),
    checkReceipts(receiptFiles, fy),
  ];

  // Roll-up status: green if everything green, amber if any amber, red if any red.
  const overall = items.some(i => i.status === 'red') ? 'red'
                : items.some(i => i.status === 'amber') ? 'amber'
                : 'green';

  return json({
    ok: true,
    fiscalYear: fy.fyLabel,
    startDate: fy.startISO,
    endDate: fy.endISO,
    overall,
    items,
  });
}

// ── Per-item checks ────────────────────────────────────────────────

function checkTransactions(rows, fy) {
  if (!rows.ok) return errorItem('transactions', 'Transactions', rows.error);
  // B Date, C Party, D Description, E Amount, F Category, G HST?, H HST, I Account, J Source, K Ref, L Inv, M MatchStatus
  const inWindow = rows.values.filter(r => isInWindow(r[0], fy));
  const total = inWindow.length;
  const categorized = inWindow.filter(r => r[4] && String(r[4]).trim()).length;
  const pct = total ? Math.round((categorized / total) * 100) : 100;
  let status = 'green', detail;
  if (total === 0) {
    status = 'red';
    detail = `No transactions in this fiscal year. Import your bank/AMEX statements.`;
  } else if (pct < 100) {
    status = pct >= 90 ? 'amber' : 'red';
    detail = `${categorized} of ${total} transactions categorized (${pct}%). Review the Transactions tab.`;
  } else {
    detail = `All ${total} transactions are categorized.`;
  }
  return {
    key: 'transactions', label: 'Transactions categorized', status, detail,
    action: total === 0
      ? { label: 'Open Statement Importer →', tab: 'importer' }
      : (pct < 100 ? { label: 'Open sheet to fix →', sheetTab: TXN_TAB } : null),
  };
}

function checkInvoices(rows, fy) {
  if (!rows.ok) return errorItem('invoices', 'Invoices', rows.error);
  // B InvNum, C Date, D Client, ..., K Status (idx 9), Q Balance Due (idx 15)
  const inWindow = rows.values.filter(r => r[0] && isInWindow(r[1], fy));
  const total = inWindow.length;
  const outstanding = inWindow.filter(r => {
    const s = String(r[9] || '').toLowerCase();
    return s !== 'paid' && s !== 'cancelled' && s !== '';
  });
  const outstandingTotal = outstanding.reduce((sum, r) => sum + (parseFloat(r[6]) || 0), 0);
  let status, detail;
  if (total === 0) {
    status = 'amber';
    detail = `No invoices issued in this fiscal year. If that's expected (cash sales only), nothing to do.`;
  } else if (outstanding.length === 0) {
    status = 'green';
    detail = `${total} invoices issued, all marked Paid or closed.`;
  } else {
    status = 'amber';
    detail = `${total} invoices issued; ${outstanding.length} still outstanding ($${outstandingTotal.toFixed(2)}). Year-end is a good time to write off uncollectibles.`;
  }
  return {
    key: 'invoices', label: 'Invoices reconciled', status, detail,
    action: outstanding.length > 0 ? { label: 'Open sheet →', sheetTab: INV_TAB } : null,
  };
}

function checkHstReturns(rows, fy) {
  if (!rows.ok) return errorItem('hst_returns', 'HST quarterly remittances', rows.error);
  // CRA Remittances cols (B-J): B Date, C Type, D Period, E Amount, F Conf, G Acct, H PDF, I Notes, J Ref
  const hstInWindow = rows.values.filter(r => String(r[1]).toLowerCase() === 'hst' && isInWindow(r[0], fy));
  const total = hstInWindow.length;
  const withReceipt = hstInWindow.filter(r => r[6]).length;
  let status, detail;
  if (total === 0) {
    status = 'red';
    detail = `No HST remittances logged for this fiscal year. Log all 4 quarters via CRA Tax Filing → Payments Log.`;
  } else if (total < 4) {
    status = 'amber';
    detail = `${total} of 4 quarterly HST remittances logged. Missing ${4 - total}.`;
  } else {
    status = withReceipt === total ? 'green' : 'amber';
    detail = total === withReceipt
      ? `All ${total} HST remittances logged with PDF receipts.`
      : `${total} HST remittances logged; ${total - withReceipt} missing PDF receipt.`;
  }
  return {
    key: 'hst_returns', label: 'HST quarterly remittances', status, detail,
    action: status !== 'green' ? { label: 'Log a payment →', tab: 'taxfiling', subTab: 'payments' } : null,
  };
}

function checkCraPayroll(rows, fy) {
  if (!rows.ok) return errorItem('cra_payroll', 'Payroll remittances', rows.error);
  const inWindow = rows.values.filter(r => String(r[1]).toLowerCase().includes('payroll') && isInWindow(r[0], fy));
  // Most small ON businesses are monthly remitters → 12 PD7As/year. Quarterly
  // remitters (under $3k average monthly withholding) → 4. We can't know which
  // without asking, so just flag if unusually low.
  const total = inWindow.length;
  let status = 'green', detail;
  if (total === 0) {
    status = 'amber';
    detail = `No payroll source-deduction remittances logged. If you didn't run payroll this year, leave as-is. Otherwise log them via Payments Log.`;
  } else {
    detail = `${total} payroll source-deduction remittances logged.`;
  }
  return {
    key: 'cra_payroll', label: 'Payroll source deductions (PD7A)', status, detail,
    action: total === 0 ? { label: 'Log a payment →', tab: 'taxfiling', subTab: 'payments' } : null,
  };
}

function checkCraCorpTax(rows, fy) {
  if (!rows.ok) return errorItem('cra_corp_tax', 'Corporate tax', rows.error);
  const inWindow = rows.values.filter(r => String(r[1]).toLowerCase().includes('corporate tax') && isInWindow(r[0], fy));
  const total = inWindow.length;
  // Corp tax payments are: instalments (typically 4/yr if owing >$3k), then a
  // final settle-up after T2 filing. We can't validate the exact number without
  // knowing instalment status; just report what's logged.
  let status = total > 0 ? 'green' : 'amber';
  let detail = total > 0
    ? `${total} corporate tax payment(s) logged for this year.`
    : `No corporate tax payments logged. Add any instalments + final balance via Payments Log.`;
  return {
    key: 'cra_corp_tax', label: 'Corporate tax payments', status, detail,
    action: total === 0 ? { label: 'Log a payment →', tab: 'taxfiling', subTab: 'payments' } : null,
  };
}

function checkStatements(bank, statementFiles, fy) {
  if (!statementFiles.ok) return errorItem(`statements_${bank.toLowerCase()}`, `${bank} statements`, statementFiles.error);
  const months = fiscalYearMonths(fy.startDate, fy.endDate);
  const re = new RegExp(`^${bank}_(\\d{4}-\\d{2})\\.`, 'i');
  const uploaded = new Set();
  for (const f of statementFiles.files) {
    const m = f.name.match(re);
    if (m) uploaded.add(m[1]);
  }
  const target = months.length;  // 12 for full FY
  const have = months.filter(({ year, month }) => uploaded.has(`${year}-${String(month).padStart(2,'0')}`)).length;
  const missing = months.filter(({ year, month }) => !uploaded.has(`${year}-${String(month).padStart(2,'0')}`))
                        .map(({ year, month }) => `${year}-${String(month).padStart(2,'0')}`);
  let status = 'green', detail;
  if (have === 0) {
    status = 'red';
    detail = `No ${bank} statements uploaded. MNP needs all ${target} monthly PDFs as backup.`;
  } else if (have < target) {
    status = 'amber';
    detail = `${have} of ${target} monthly ${bank} statements uploaded. Missing: ${missing.slice(0, 6).join(', ')}${missing.length > 6 ? '…' : ''}`;
  } else {
    detail = `All ${target} monthly ${bank} statements uploaded.`;
  }
  return {
    key: `statements_${bank.toLowerCase()}`, label: `${bank} monthly statements`, status, detail,
    action: status !== 'green' ? { label: 'Upload statements →', tab: 'yearend' } : null,
  };
}

function checkPayroll(rows, fy) {
  if (!rows.ok) return errorItem('payroll', 'Payroll runs + T4s', rows.error);
  // Payroll cols B-Q: B PayDate, C Employee, ... Q Status
  const inWindow = rows.values.filter(r => r[0] && isInWindow(r[0], fy));
  const total = inWindow.length;
  const employees = new Set(inWindow.map(r => r[1]).filter(Boolean));
  let status = 'green', detail;
  if (total === 0) {
    status = 'amber';
    detail = `No payroll runs in this fiscal year. If you had no employees, leave as-is.`;
  } else {
    detail = `${total} pay run${total === 1 ? '' : 's'} for ${employees.size} employee${employees.size === 1 ? '' : 's'}. T4s due to CRA + employees by Feb 28.`;
  }
  return {
    key: 'payroll', label: 'Payroll runs + T4 readiness', status, detail,
    action: total > 0 ? { label: 'Generate T4s →', tab: 'taxfiling', subTab: 't4' } : null,
  };
}

function checkReceipts(receiptFiles, fy) {
  if (!receiptFiles.ok) return errorItem('receipts', 'Expense receipts', receiptFiles.error);
  const total = receiptFiles.files.length;
  // Heuristic: a small Ontario contracting business should have ~50-200+ receipts
  // per year between materials, fuel, tools, meals. <20 is suspicious.
  let status, detail;
  if (total === 0) {
    status = 'red';
    detail = `No receipts in Drive for this fiscal year. CRA wants receipts kept 6 years.`;
  } else if (total < 20) {
    status = 'amber';
    detail = `Only ${total} receipts archived. Make sure every business expense has a photo or PDF.`;
  } else {
    status = 'green';
    detail = `${total} receipts archived in Drive.`;
  }
  return {
    key: 'receipts', label: 'Expense receipts archived', status, detail,
    action: status !== 'green' ? { label: 'Open Receipt Scanner →', tab: 'receipt' } : null,
  };
}

// ── Helpers ────────────────────────────────────────────────────────

function errorItem(key, label, error) {
  return {
    key, label,
    status: 'amber',
    detail: `Could not check: ${error}`,
    action: null,
  };
}

async function safeReadRange(env, userId, range) {
  try {
    const r = await readRange(env, userId, range);
    if (!r.ok) return { ok: false, error: r.error };
    return { ok: true, values: r.values || [] };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// Synthetic result when a tab the checklist expects isn't in the sheet at
// all. Same shape as safeReadRange so the checks treat it as an empty range
// with a helpful explanation, rather than a hard failure.
function missingTabResult(logicalName) {
  return {
    ok: false,
    error: `'${logicalName}' tab not found in your sheet. Run Settings → Update sheet to latest schema, or check your tab names.`,
    values: [],
  };
}

async function listStatementFiles(accessToken, fy) {
  try {
    const parentId = await findOrCreateFolder(accessToken, YEAREND_PARENT, null);
    if (!parentId) return { ok: true, files: [] };
    const fyFolderId = await findOrCreateFolder(accessToken, fy.fyLabel, parentId);
    if (!fyFolderId) return { ok: true, files: [] };
    const stmtFolderId = await findOrCreateFolder(accessToken, 'Statements', fyFolderId);
    if (!stmtFolderId) return { ok: true, files: [] };
    const files = await listFolderFiles(accessToken, stmtFolderId);
    return { ok: true, files };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

async function listReceiptFiles(accessToken, fy) {
  // Receipts are stored under 'AI Bookkeeper Receipts/{calendar year}/'. A
  // fiscal year that crosses calendar boundaries (e.g. Apr 2025 → Mar 2026)
  // touches two calendar-year folders. We list both and filter by createdTime.
  try {
    const parentId = await findOrCreateFolder(accessToken, RECEIPTS_PARENT, null);
    if (!parentId) return { ok: true, files: [] };

    const startYear = fy.startDate.getUTCFullYear();
    const endYear = fy.endDate.getUTCFullYear();
    const calendarYears = [];
    for (let y = startYear; y <= endYear; y++) calendarYears.push(String(y));

    let allFiles = [];
    for (const yr of calendarYears) {
      const yearFolderId = await findOrCreateFolder(accessToken, yr, parentId);
      if (!yearFolderId) continue;
      const files = await listFolderFiles(accessToken, yearFolderId);
      allFiles = allFiles.concat(files);
    }

    // Filter to fiscal-year window by createdTime
    const startMs = fy.startDate.getTime();
    const endMs = fy.endDate.getTime();
    const filtered = allFiles.filter(f => {
      const t = f.createdTime ? Date.parse(f.createdTime) : NaN;
      return Number.isFinite(t) && t >= startMs && t <= endMs;
    });
    return { ok: true, files: filtered };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function isInWindow(dateStr, fy) {
  if (!dateStr) return false;
  const t = Date.parse(dateStr);
  if (!Number.isFinite(t)) return false;
  return t >= fy.startDate.getTime() && t <= fy.endDate.getTime();
}

function resolveFiscalYear(fyParam, fye) {
  if (fyParam && /^FY\d{4}$/.test(fyParam)) {
    const fyEndYear = parseInt(fyParam.slice(2), 10);
    const map = {january:1,february:2,march:3,april:4,may:5,june:6,july:7,august:8,september:9,october:10,november:11,december:12};
    const m = String(fye).match(/^([a-zA-Z]+)\s+(\d{1,2})$/);
    const month = m ? (map[m[1].toLowerCase()] || 12) : 12;
    const day = m ? (parseInt(m[2], 10) || 31) : 31;
    return fiscalYearOf(fye, new Date(Date.UTC(fyEndYear, month - 1, day)));
  }
  return fiscalYearOf(fye);
}
