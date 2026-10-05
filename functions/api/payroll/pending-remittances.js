// ════════════════════════════════════════════════════════════════════
// GET /api/payroll/pending-remittances
//
// Returns pay runs where Status === 'Paid' and Remittance Due is populated,
// grouped by the remittance due date (usually one group per month for a
// monthly remitter). Also returns a short list of recently-completed
// remittances pulled from 📒 Transactions (ref prefix "CRA-REMIT-").
//
// Response shape:
// {
//   ok: true,
//   groups: [
//     {
//       remittanceDue: '2026-05-15',
//       label: 'Due May 15, 2026',
//       runs: [
//         { sheetRow, payDate, employee, cpp, ei, fedTax, onTax,
//           employerCpp, employerEi, employeePart, employerPart, total },
//         ...
//       ],
//       totalCpp, totalEi, totalFedTax, totalOnTax, totalEmployerCpp, totalEmployerEi,
//       totalEmployee, totalEmployer, totalAmount   (totalAmount = what CRA receives)
//     },
//   ],
//   recentRemittances: [
//     { date, amount, ref, description },
//     ...
//   ]
// }
// ════════════════════════════════════════════════════════════════════

import { authenticateRequest, json, options } from '../../_shared.js';
import { readRange, writeRange } from '../../_sheets.js';
import { PAYROLL_TAB, TXN_TAB, num, round2, loadPayrollRows, loadRemitter } from '../../_payroll_sheet.js';
import { remittanceDueDate } from '../../_payroll.js';

export const onRequestOptions = () => options();

export async function onRequestGet({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  const userId = auth.userId;

  const remitter = await loadRemitter(env, userId);
  const { groups, healed } = await loadPendingGroups(env, userId, remitter);
  const recentRemittances = await loadRecentRemittances(env, userId);

  return json({ ok: true, remitter, groups, recentRemittances, healed });
}

// ── Pending remittance groups ───────────────────────────────────────

async function loadPendingGroups(env, userId, remitter = 'monthly') {
  const res = await loadPayrollRows(env, userId);
  if (!res.ok) return { groups: [], healed: 0 };

  // Group by remittance due date. The due date is RECOMPUTED from the pay
  // date and the current remitter type, so switching monthly ↔ quarterly
  // in Settings re-buckets rows that were committed under the old type.
  // Rows still 'Paid' whose stored due date (col P) disagrees get rewritten
  // so the sheet matches what the app shows.
  const byDue = new Map();
  const heals = [];

  for (const r of res.rows) {
    if (!r.remitDue) continue;                                 // no deductions owed
    if (r.status.toLowerCase() !== 'paid') continue;           // already remitted or cancelled
    const dueNow = remittanceDueDate(r.payDate, remitter) || r.remitDue;
    if (dueNow !== r.remitDue) heals.push({ sheetRow: r.sheetRow, due: dueNow });
    r.remitDue = dueNow;

    const cppNum = r.cpp;
    const eiNum = r.ei;
    const fedNum = r.fedTax;
    const onNum = r.onTax;
    const employeePart = round2(cppNum + eiNum + fedNum + onNum);
    const employerPart = round2(r.employerCpp + r.employerEi);
    const total = round2(employeePart + employerPart);   // what CRA receives
    if (total <= 0) continue;

    const dueKey = r.remitDue;
    if (!byDue.has(dueKey)) {
      byDue.set(dueKey, {
        remittanceDue: dueKey,
        label: formatDueLabel(dueKey),
        runs: [],
        totalCpp: 0, totalEi: 0, totalFedTax: 0, totalOnTax: 0,
        totalEmployerCpp: 0, totalEmployerEi: 0,
        totalEmployee: 0, totalEmployer: 0, totalAmount: 0,
      });
    }
    const g = byDue.get(dueKey);
    g.runs.push({
      sheetRow: r.sheetRow,
      payDate: r.payDate,
      employee: r.employee,
      cpp: cppNum, ei: eiNum, fedTax: fedNum, onTax: onNum,
      employerCpp: r.employerCpp, employerEi: r.employerEi,
      employeePart, employerPart, total,
    });
    g.totalCpp += cppNum;
    g.totalEi += eiNum;
    g.totalFedTax += fedNum;
    g.totalOnTax += onNum;
    g.totalEmployerCpp += r.employerCpp;
    g.totalEmployerEi += r.employerEi;
    g.totalEmployee += employeePart;
    g.totalEmployer += employerPart;
    g.totalAmount += total;
  }

  // Round totals, sort by due date
  const groups = [...byDue.values()].map(g => ({
    ...g,
    totalCpp: round2(g.totalCpp),
    totalEi: round2(g.totalEi),
    totalFedTax: round2(g.totalFedTax),
    totalOnTax: round2(g.totalOnTax),
    totalEmployerCpp: round2(g.totalEmployerCpp),
    totalEmployerEi: round2(g.totalEmployerEi),
    totalEmployee: round2(g.totalEmployee),
    totalEmployer: round2(g.totalEmployer),
    totalAmount: round2(g.totalAmount),
  }));
  groups.sort((a, b) => a.remittanceDue.localeCompare(b.remittanceDue));

  for (const h of heals) {
    try { await writeRange(env, userId, `'${PAYROLL_TAB}'!P${h.sheetRow}`, [[h.due]]); } catch { /* cosmetic */ }
  }
  return { groups, healed: heals.length };
}

// ── Recent remittances (from 📒 Transactions) ────────────────────────

async function loadRecentRemittances(env, userId) {
  const result = await readRange(env, userId, `'${TXN_TAB}'!B12:M`);
  if (!result.ok) return [];

  // One remittance = one bank payment, but it is booked as two ledger rows
  // (…-EE employee deductions, …-ER employer share). Merge them by base ref.
  const byRef = new Map();
  for (const row of result.values) {
    if (!row || !row[0]) continue;
    // Columns 0 Date, 1 Party, 2 Description, 3 Amount, 4 Category,
    //         5 HST?, 6 HST$, 7 Account, 8 Source, 9 Ref, ...
    const [date, party, desc, amount, , , , , source, ref] = row;
    const refStr = String(ref || '');
    if (!refStr.startsWith('CRA-REMIT-')) continue;
    const base = refStr.replace(/-(EE|ER)$/, '');
    const part = refStr.endsWith('-ER') ? 'employer' : 'employee';
    if (!byRef.has(base)) {
      byRef.set(base, { date: normalizeDate(date), amount: 0, employee: 0, employer: 0, ref: base, description: desc || '', party: party || 'CRA', source: source || '' });
    }
    const it = byRef.get(base);
    const amt = Math.abs(num(amount));
    it.amount = round2(it.amount + amt);
    it[part] = round2(it[part] + amt);
    if (part === 'employee') it.description = String(desc || '').replace(/^CRA source deductions \(employee CPP\/EI\/tax\) — /, '');
  }
  const items = [...byRef.values()];
  items.sort((a, b) => b.date.localeCompare(a.date));
  return items.slice(0, 10);
}

// ── Helpers ──

function normalizeDate(v) {
  if (!v) return '';
  const d = new Date(v);
  if (isNaN(d.getTime())) return String(v);
  return d.toISOString().slice(0, 10);
}

function formatDueLabel(iso) {
  try {
    const d = new Date(iso);
    const m = d.toLocaleString('en-CA', { month: 'long', year: 'numeric' });
    return `Due ${d.toLocaleDateString('en-CA', { month: 'short', day: 'numeric', year: 'numeric' })}`;
  } catch {
    return `Due ${iso}`;
  }
}
