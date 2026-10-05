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

export function round2(n) { return Math.round((Number(n) || 0) * 100) / 100; }

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
  if (!result.ok) return [];
  const startT = Date.parse(startIso);
  const endT = Date.parse(endIso);
  const entries = [];
  for (const row of result.values) {
    if (!row || !row[WL_COL.date]) continue;
    if (row[WL_COL.employee] !== employeeName) continue;
    const t = Date.parse(row[WL_COL.date]);
    if (isNaN(t) || t < startT || t > endT) continue;
    entries.push({
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
  if (!result.ok) return { ok: false, error: result.error, rows: [] };
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
  const ytd = { gross: 0, cppBase: 0, cpp2: 0, fedTax: 0, onTax: 0 };
  const res = await loadPayrollRows(env, userId);
  if (!res.ok) return ytd;
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
    ytd.fedTax += r.fedTax;
    ytd.onTax += r.onTax;
  }
  ytd.gross = round2(ytd.gross);
  ytd.cppBase = round2(ytd.cppBase);
  ytd.fedTax = round2(ytd.fedTax);
  ytd.onTax = round2(ytd.onTax);
  return ytd;
}

/** Existing (employee, payDate) Payroll row, for the idempotency check. */
export async function findExistingPayrollRow(env, userId, employeeName, payDateIso) {
  const res = await loadPayrollRows(env, userId);
  if (!res.ok) return null;
  const hit = res.rows.find(r => r.employee === employeeName && r.payDate === payDateIso);
  return hit ? { sheetRow: hit.sheetRow, row: hit } : null;
}
