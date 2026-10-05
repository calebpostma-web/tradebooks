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
//         { sheetRow, payDate, employee, cpp, fedTax, onTax, total },
//         ...
//       ],
//       totalCpp, totalFedTax, totalOnTax, totalAmount,
//     },
//   ],
//   recentRemittances: [
//     { date, amount, ref, description },
//     ...
//   ]
// }
// ════════════════════════════════════════════════════════════════════

import { readRange } from '../../_sheets.js';
import { authenticateRequest, json, options } from '../../_shared.js';
import { TXN_TAB, num, round2, loadPayrollRows } from '../../_payroll_sheet.js';

export const onRequestOptions = () => options();

export async function onRequestGet({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  const userId = auth.userId;

  const groups = await loadPendingGroups(env, userId);
  const recentRemittances = await loadRecentRemittances(env, userId);

  return json({ ok: true, groups, recentRemittances });
}

// ── Pending remittance groups ───────────────────────────────────────

async function loadPendingGroups(env, userId) {
  const res = await loadPayrollRows(env, userId);
  if (!res.ok) return [];

  // Group by remittance due date
  const byDue = new Map();

  for (const r of res.rows) {
    if (!r.remitDue) continue;                                 // no deductions owed
    if (r.status.toLowerCase() !== 'paid') continue;           // already remitted or cancelled

    const cppNum = r.cpp;
    const fedNum = r.fedTax;
    const onNum = r.onTax;
    const total = round2(cppNum + fedNum + onNum);
    if (total <= 0) continue;

    const dueKey = r.remitDue;
    if (!byDue.has(dueKey)) {
      byDue.set(dueKey, {
        remittanceDue: dueKey,
        label: formatDueLabel(dueKey),
        runs: [],
        totalCpp: 0, totalFedTax: 0, totalOnTax: 0, totalAmount: 0,
      });
    }
    const g = byDue.get(dueKey);
    g.runs.push({
      sheetRow: r.sheetRow,
      payDate: r.payDate,
      employee: r.employee,
      cpp: cppNum, fedTax: fedNum, onTax: onNum, total,
    });
    g.totalCpp += cppNum;
    g.totalFedTax += fedNum;
    g.totalOnTax += onNum;
    g.totalAmount += total;
  }

  // Round totals, sort by due date
  const groups = [...byDue.values()].map(g => ({
    ...g,
    totalCpp: round2(g.totalCpp),
    totalFedTax: round2(g.totalFedTax),
    totalOnTax: round2(g.totalOnTax),
    totalAmount: round2(g.totalAmount),
  }));
  groups.sort((a, b) => a.remittanceDue.localeCompare(b.remittanceDue));
  return groups;
}

// ── Recent remittances (from 📒 Transactions) ────────────────────────

async function loadRecentRemittances(env, userId) {
  const result = await readRange(env, userId, `'${TXN_TAB}'!B12:M`);
  if (!result.ok) return [];

  const items = [];
  for (const row of result.values) {
    if (!row || !row[0]) continue;
    // Columns 0 Date, 1 Party, 2 Description, 3 Amount, 4 Category,
    //         5 HST?, 6 HST$, 7 Account, 8 Source, 9 Ref, ...
    const [date, party, desc, amount, , , , , source, ref] = row;
    const refStr = String(ref || '');
    if (!refStr.startsWith('CRA-REMIT-')) continue;

    items.push({
      date: normalizeDate(date),
      amount: Math.abs(num(amount)),
      ref: refStr,
      description: desc || '',
      party: party || 'CRA',
      source: source || '',
    });
  }
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
