// ════════════════════════════════════════════════════════════════════
// PAYROLL SHEET HELPERS — shared by every /api/payroll endpoint
//
// One copy of: the tab names, the column layouts, the number parser, and
// the sheet readers (employee record, Work Log range, YTD state, existing
// pay-run lookup). calculate.js and run.js used to carry private copies
// of these — the moment one changed and the other didn't, preview and
// commit would have disagreed. Now they cannot.
//
// WHY num() EXISTS: Sheets returns FORMATTED values by default. The Rate,
// Gross, CPP and tax columns are currency-formatted, so the API hands us
// "$30.00" or "$9,000.00". parseFloat("$30.00") is NaN → 0, which made a
// $9,000 pay run preview as $0 and every YTD total read as zero. Always
// go through num() for any money or hours cell read from the sheet.
// ════════════════════════════════════════════════════════════════════

import { readRange } from './_sheets.js';

/**
 * Thrown when a payroll read can't reach the sheet (expired Google grant,
 * missing sheet, API error). Callers MUST surface it — never fall back to
 * zeros. A silent zero made a Preview show "YTD $0 / 0 Work Log entries /
 * CPP $0" and would have let a wrong pay run be committed (Nov 16 2026).
 */
export class SheetReadError extends Error {
  constructor(what, result) {
    super(`Couldn't read ${what} from your Google Sheet: ${result?.error || 'unknown error'}`);
    this.name = 'SheetReadError';
    this.needsReauth = !!result?.needsReauth || /revoked|reconnect|invalid_grant/i.test(String(result?.error || ''));
  }
}

/** JSON body for a SheetReadError, shared by the endpoints. */
export function sheetReadErrorBody(err) {
  return {
    ok: false,
    error: err.needsReauth
      ? 'Google connection lost — the site can\'t read your sheet right now. Log out and sign in with Google again, then retry. Nothing was saved.'
      : err.message + ' Nothing was saved.',
    needsReauth: !!err.needsReauth,
  };
}

export const WORK_LOG_TAB = '📝 Work Log';
export const PAYROLL_TAB = '💼 Payroll';
export const TXN_TAB = '📒 Transactions';

/** Default wage category when the user's own list has no wages-like name. */
export const DEFAULT_WAGE_CATEGORY = 'Wages & Salaries';

// 💼 Payroll data columns B..U (row array indices 0..19)
export const PAY_COL = {
  payDate: 0, employee: 1, age: 2, business: 3, workDesc: 4,
  hours: 5, rate: 6, gross: 7, cpp: 8, ei: 9, fedTax: 10, onTax: 11,
  netPay: 12, ytdGross: 13, remitDue: 14, status: 15,
  // Added Oct 2026 (Migration 15). Blank on older rows — readers fall back.
  period: 16, employerCpp: 17, employerEi: 18, stub: 19,
};
export const PAY_READ_RANGE = `'${PAYROLL_TAB}'!B12:U`;

// 📝 Work Log data columns B..I (indices 0..7)
export const WL_COL = { date: 0, employee: 1, business: 2, task: 3, hours: 4, rate: 5, notes: 6, audit: 7 };
export const WL_READ_RANGE = `'${WORK_LOG_TAB}'!B12:I`;

// ─── Number parsing ─────────────────────────────────────────────────

/**
 * Parse a sheet cell into a number. Accepts raw numbers, "$1,234.56",
 * "(41.04)" (accounting negative), "-41.04", "  12 ", "". Anything
 * unparseable → 0.
 */
export function num(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  let s = String(v).trim();
  if (!s) return 0;
  let neg = false;
  if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
  s = s.replace(/[^0-9.\-]/g, '');          // drop $, commas, spaces, CA$, etc.
  if (s === '' || s === '-' || s === '.') return 0;
  const n = parseFloat(s);
  if (!isFinite(n)) return 0;
  return neg ? -n : n;
}

export function round2(n) { const x = Number(n) || 0; return Math.round(Number((x * 100).toFixed(6))) / 100; }

/** Text typed by a user is written with USER_ENTERED; never let it start a formula. */
export function safeText(v) {
  const s = String(v ?? '').trim();
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

/** True only for a real calendar date in YYYY-MM-DD form (rejects 2026-11-31, which JS rolls to Dec 1). */
export function isRealIsoDate(v) {
  const s = String(v || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  const d = new Date(s + 'T00:00:00Z');
  return !isNaN(d) && d.toISOString().slice(0, 10) === s;
}

/** Normalise a sheet date cell ("Sep 30, 2026", "2026-09-30") to ISO, or ''. */
export function isoDate(v) {
  if (!v) return '';
  const d = new Date(v);
  if (isNaN(d.getTime())) return '';
  return d.toISOString().slice(0, 10);
}

// ─── Employee record (D1 profiles.employees JSON) ───────────────────

export async function loadEmployees(env, userId) {
  try {
    const row = await env.DB.prepare('SELECT employees FROM profiles WHERE user_id = ?')
      .bind(userId).first();
    if (!row || !row.employees) return [];
    const list = JSON.parse(row.employees);
    return Array.isArray(list) ? list : [];
  } catch {
    return [];
  }
}

export async function loadEmployee(env, userId, employeeId) {
  const list = await loadEmployees(env, userId);
  return list.find(e => e.id === employeeId) || null;
}

/**
 * The category payroll rows are filed under. Prefers the user's own
 * expense-category name that looks like wages (Andrea's list says "Wages";
 * the default list says "Wages & Salaries"), so Accountant tabs file it
 * in the column she expects.
 */
export async function resolveWageCategory(env, userId) {
  try {
    const row = await env.DB.prepare('SELECT custom_expense_cats FROM profiles WHERE user_id = ?')
      .bind(userId).first();
    const cats = row?.custom_expense_cats ? JSON.parse(row.custom_expense_cats) : [];
    if (Array.isArray(cats)) {
      const hit = cats.find(c => /wage|salar|payroll/i.test(String(c)));
      if (hit) return String(hit);
    }
  } catch { /* fall through */ }
  return DEFAULT_WAGE_CATEGORY;
}

/** 'monthly' | 'quarterly' from the profile (CRA-assigned remitter type). */
export async function loadRemitter(env, userId) {
  try {
    const row = await env.DB.prepare('SELECT payroll_remitter FROM profiles WHERE user_id = ?').bind(userId).first();
    return String(row?.payroll_remitter || '').toLowerCase() === 'quarterly' ? 'quarterly' : 'monthly';
  } catch {
    return 'monthly';
  }
}

// ─── Work Log ───────────────────────────────────────────────────────

/**
 * Work Log rows for one employee with Date ∈ [startIso, endIso] (inclusive).
 * Returns [{date, business, task, hours, rate, notes, audit}].
 */
export async function loadWorkLogInRange(env, userId, employeeName, startIso, endIso) {
  const result = await readRange(env, userId, WL_READ_RANGE);
  if (!result.ok) throw new SheetReadError('the Work Log', result);
  const startT = Date.parse(startIso);
  const endT = Date.parse(endIso);
  const entries = [];
  for (let i = 0; i < result.values.length; i++) {
    const row = result.values[i];
    if (!row || !row[WL_COL.date]) continue;
    if (row[WL_COL.employee] !== employeeName) continue;
    if (isRemovedEntry(row)) continue;
    const t = Date.parse(row[WL_COL.date]);
    if (isNaN(t) || t < startT || t > endT) continue;
    entries.push({
      sheetRow: 12 + i,
      date: row[WL_COL.date],
      business: row[WL_COL.business] || '',
      task: row[WL_COL.task] || '',
      hours: num(row[WL_COL.hours]),
      rate: num(row[WL_COL.rate]),
      notes: row[WL_COL.notes] || '',
      audit: row[WL_COL.audit] || '',
    });
  }
  return entries;
}

/** A Work Log row marked removed (audit "|removed:") never counts toward pay. */
export function isRemovedEntry(row) {
  return /\|removed:/.test(String(row?.[WL_COL.audit] || ''));
}

/** Parse a Payroll "Period" cell ("2026-10-01 → 2026-10-31") into ISO dates. */
export function parsePeriod(text) {
  const m = /(\d{4}-\d{2}-\d{2})\s*(?:→|->|to|–|-)\s*(\d{4}-\d{2}-\d{2})/.exec(String(text || ''));
  return m ? { start: m[1], end: m[2] } : null;
}

/**
 * Periods already paid for one employee: every committed, non-cancelled
 * Payroll row with a readable Period. Hours dated inside one of these were
 * (or should have been) paid by that run — editing them after the fact
 * would make the Work Log disagree with what was actually paid.
 */
export function paidPeriods(payrollRows, employeeName) {
  const out = [];
  for (const r of payrollRows) {
    if (r.employee !== employeeName || r.status.toLowerCase() === 'cancelled') continue;
    const p = parsePeriod(r.period);
    if (!p) continue;
    out.push({ ...p, startT: Date.parse(p.start), endT: Date.parse(p.end), payDate: r.payDate, sheetRow: r.sheetRow, status: r.status });
  }
  return out;
}

/** The paid period (if any) that covers a date, for one employee. */
export function findPaidPeriod(periods, dateValue) {
  const t = Date.parse(isoDate(dateValue) || dateValue);
  if (isNaN(t)) return null;
  return periods.find(p => t >= p.startT && t <= p.endT + 86399999) || null;
}

/**
 * Work Log entries for one employee dated OUTSIDE [start, end] but within
 * `days` of it, and not covered by any committed (non-cancelled) pay run's
 * period. Shown on the Preview as "hours just outside this period" — the
 * exact trap Andrea hit (catch-up hours dated on the pay date, Oct 4,
 * instead of inside Apr 1–Sep 30).
 */
export async function loadWorkLogNearby(env, userId, employeeName, startIso, endIso, days = 62) {
  const result = await readRange(env, userId, WL_READ_RANGE);
  if (!result.ok) throw new SheetReadError('the Work Log', result);
  const pay = await loadPayrollRows(env, userId);
  if (!pay.ok) throw new SheetReadError('past pay runs (Payroll tab)', pay);
  const covered = paidPeriods(pay.rows, employeeName).map(p => [p.startT, p.endT]);
  const startT = Date.parse(startIso), endT = Date.parse(endIso), win = days * 86400000;
  const out = [];
  for (let i = 0; i < result.values.length; i++) {
    const row = result.values[i];
    if (!row || !row[WL_COL.date] || row[WL_COL.employee] !== employeeName || isRemovedEntry(row)) continue;
    const t = Date.parse(row[WL_COL.date]);
    if (isNaN(t) || (t >= startT && t <= endT)) continue;
    if (t < startT - win || t > endT + win) continue;
    if (covered.some(([a, b]) => t >= a && t <= b + 86399999)) continue;
    out.push({ sheetRow: 12 + i, date: row[WL_COL.date], isoDate: isoDate(row[WL_COL.date]), hours: num(row[WL_COL.hours]), rate: num(row[WL_COL.rate]), task: row[WL_COL.task] || '', after: t > endT });
  }
  return out;
}

/**
 * Period checks shared by Preview (calculate) and Commit (run):
 *   alreadyPaid — Work Log entries in this period that a committed,
 *                 non-cancelled run already covered (paying them again
 *                 would pay the same hours twice). Commit refuses these.
 *   overlaps    — committed runs whose period overlaps this one.
 *   nearby      — unpaid entries dated just outside this period (the
 *                 "hours dated on the pay day" trap). Preview shows them.
 * `excludeSheetRow`: the run being overwritten (same employee + pay date).
 */
export async function loadPeriodChecks(env, userId, employeeName, periodStart, periodEnd, wlEntries, excludeSheetRow = null) {
  const pay = await loadPayrollRows(env, userId);
  if (!pay.ok) throw new SheetReadError('past pay runs (Payroll tab)', pay);
  const periods = paidPeriods(pay.rows, employeeName).filter(p => p.sheetRow !== excludeSheetRow);
  const startT = Date.parse(periodStart), endT = Date.parse(periodEnd);
  const brief = p => ({ payDate: p.payDate, start: p.start, end: p.end, status: p.status });
  const alreadyPaid = [];
  for (const e of wlEntries) {
    const p = findPaidPeriod(periods, e.date);
    if (p) alreadyPaid.push({ ...e, paidBy: brief(p) });
  }
  const overlaps = periods.filter(p => p.startT <= endT && p.endT >= startT).map(brief);
  const nearby = await loadWorkLogNearby(env, userId, employeeName, periodStart, periodEnd);
  return { alreadyPaid, overlaps, nearby };
}

// ─── Payroll rows ───────────────────────────────────────────────────

/**
 * Parse one 💼 Payroll row array into a typed record. Employer share falls
 * back to the CRA identities when the S/T columns are blank (rows written
 * before Migration 15): employer CPP = employee CPP (base + CPP2), employer
 * EI = 1.4 × employee EI.
 */
export function parsePayrollRow(row, sheetRow) {
  const cpp = num(row[PAY_COL.cpp]);
  const ei = num(row[PAY_COL.ei]);
  const employerCppCell = row[PAY_COL.employerCpp];
  const employerEiCell = row[PAY_COL.employerEi];
  const employerCpp = employerCppCell == null || employerCppCell === '' ? round2(cpp) : num(employerCppCell);
  const employerEi = employerEiCell == null || employerEiCell === '' ? round2(ei * 1.4) : num(employerEiCell);
  return {
    sheetRow,
    payDate: isoDate(row[PAY_COL.payDate]) || String(row[PAY_COL.payDate] || ''),
    employee: row[PAY_COL.employee] || '',
    age: row[PAY_COL.age] ?? '',
    business: row[PAY_COL.business] || '',
    workDesc: row[PAY_COL.workDesc] || '',
    hours: num(row[PAY_COL.hours]),
    rate: num(row[PAY_COL.rate]),
    gross: num(row[PAY_COL.gross]),
    cpp, ei,
    fedTax: num(row[PAY_COL.fedTax]),
    onTax: num(row[PAY_COL.onTax]),
    netPay: num(row[PAY_COL.netPay]),
    ytdGross: num(row[PAY_COL.ytdGross]),
    remitDue: isoDate(row[PAY_COL.remitDue]),
    status: String(row[PAY_COL.status] || ''),
    period: String(row[PAY_COL.period] || ''),
    employerCpp, employerEi,
    stub: String(row[PAY_COL.stub] || ''),
    // What CRA actually receives for this run
    employeeDeductions: round2(cpp + ei + num(row[PAY_COL.fedTax]) + num(row[PAY_COL.onTax])),
    employerShare: round2(employerCpp + employerEi),
    totalRemittance: round2(cpp + ei + num(row[PAY_COL.fedTax]) + num(row[PAY_COL.onTax]) + employerCpp + employerEi),
  };
}

/** All Payroll rows, parsed, with 1-indexed sheet row numbers. */
export async function loadPayrollRows(env, userId) {
  const result = await readRange(env, userId, PAY_READ_RANGE);
  if (!result.ok) return { ok: false, error: result.error, needsReauth: !!result.needsReauth, rows: [] };
  const rows = [];
  for (let i = 0; i < result.values.length; i++) {
    const row = result.values[i];
    if (!row || !row[PAY_COL.payDate]) continue;
    rows.push(parsePayrollRow(row, 12 + i));
  }
  return { ok: true, rows };
}

/**
 * YTD state for one employee: calendar year of payDateIso, rows strictly
 * BEFORE payDateIso, Cancelled rows ignored. CPP column holds base + CPP2
 * combined, so cpp2 is always 0 here and the engine treats cppBase as the
 * combined figure (it caps against the combined annual max).
 */
export async function loadYtdState(env, userId, employeeName, payDateIso) {
  // cppBase      = employee CPP actually deducted (col J), incl. recoveries
  // cppRequired  = the CPP each run SHOULD have had (col S, the employer
  //                match — employer always remits the full correct share)
  // cppShortfall = cppRequired − cppBase: employee CPP still to recover
  const ytd = { gross: 0, cppBase: 0, cpp2: 0, fedTax: 0, onTax: 0, cppRequired: 0, cppShortfall: 0 };
  const res = await loadPayrollRows(env, userId);
  if (!res.ok) throw new SheetReadError('past pay runs (Payroll tab)', res);
  const payT = Date.parse(payDateIso);
  const payYear = new Date(payDateIso).getUTCFullYear();
  for (const r of res.rows) {
    if (r.employee !== employeeName) continue;
    if (r.status.toLowerCase() === 'cancelled') continue;
    const rowT = Date.parse(r.payDate);
    if (isNaN(rowT)) continue;
    if (new Date(r.payDate).getUTCFullYear() !== payYear) continue;
    if (rowT >= payT) continue;
    ytd.gross += r.gross;
    ytd.cppBase += r.cpp;
    ytd.cppRequired += r.employerCpp;
    ytd.fedTax += r.fedTax;
    ytd.onTax += r.onTax;
  }
  ytd.gross = round2(ytd.gross);
  ytd.cppBase = round2(ytd.cppBase);
  ytd.cppRequired = round2(ytd.cppRequired);
  ytd.cppShortfall = round2(Math.max(0, ytd.cppRequired - ytd.cppBase));
  ytd.fedTax = round2(ytd.fedTax);
  ytd.onTax = round2(ytd.onTax);
  return ytd;
}

/** Existing (employee, payDate) Payroll row, for the idempotency check. */
export async function findExistingPayrollRow(env, userId, employeeName, payDateIso) {
  const res = await loadPayrollRows(env, userId);
  if (!res.ok) throw new SheetReadError('past pay runs (Payroll tab)', res);
  // A Cancelled run on the same date is history, not a clash — redoing a
  // cancelled pay on the same pay date is the normal correction path.
  const hit = res.rows.find(r => r.employee === employeeName && r.payDate === payDateIso && r.status.toLowerCase() !== 'cancelled');
  return hit ? { sheetRow: hit.sheetRow, row: hit } : null;
}
