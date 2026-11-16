// ════════════════════════════════════════════════════════════════════
// /api/payroll/cpp-review  — "Check CPP against CRA's method"
//
// GET  ?year=2026            → per run: CPP deducted vs CPP CRA's method
//                              requires, and the employee shortfall still
//                              to recover (or over-deduction).
// POST { year, apply: true } → writes the CORRECT contribution into each
//                              run's Employer CPP (col S). Employee
//                              deductions already taken (col J) are left
//                              alone — that's what actually happened to
//                              the paycheque. The difference is recovered
//                              on later pays (CRA cap: up to that pay's
//                              normal CPP), see applyCppRecovery().
//
// Why col S: the employer always remits its FULL correct share for each
// pay, so S = what the pay should have carried. Then
//   Σ(S − J) = employee CPP still owed by the employee,
//   Σ(J + S) = what goes to CRA,
//   Σ J      = T4 box 16 at year end.
//
// Remitted rows are reported but not changed: if one was short, the
// difference has to go to CRA as a separate payment — we flag it.
// ════════════════════════════════════════════════════════════════════

import { authenticateRequest, json, options } from '../../_shared.js';
import { writeRange } from '../../_sheets.js';
import { calculateCpp, normalizePayFrequency, PAY_FREQUENCIES, periodsCovered, cppApplies } from '../../_payroll.js';
import { PAYROLL_TAB, loadPayrollRows, loadEmployees, round2, SheetReadError, sheetReadErrorBody } from '../../_payroll_sheet.js';

export const onRequestOptions = () => options();

function periodOf(run) {
  const m = /(\d{4}-\d{2}-\d{2})\s*(?:→|->|to|–|-)\s*(\d{4}-\d{2}-\d{2})/.exec(run.period || '');
  return m ? { start: m[1], end: m[2] } : null;
}

async function review(env, userId, year) {
  const res = await loadPayrollRows(env, userId);
  if (!res.ok) throw new SheetReadError('past pay runs (Payroll tab)', res);
  const employees = await loadEmployees(env, userId);
  // Same name can appear more than once (re-saved records) — prefer the active one.
  const byName = {};
  for (const e of employees) if (!byName[e.name] || e.active !== false) byName[e.name] = e;

  const runs = res.rows
    .filter(r => r.status.toLowerCase() !== 'cancelled' && String(r.payDate).startsWith(String(year)))
    .sort((a, b) => a.payDate.localeCompare(b.payDate) || a.sheetRow - b.sheetRow);

  const perEmp = {};
  const out = [];
  for (const r of runs) {
    const emp = byName[r.employee] || {};
    const P = PAY_FREQUENCIES[normalizePayFrequency(emp.payFrequency)];
    const per = periodOf(r);
    const N = per ? periodsCovered(per.start, per.end, P) : 1;
    const st = perEmp[r.employee] || (perEmp[r.employee] = { required: 0, deducted: 0, gross: 0 });
    const { cppBase, cpp2, math } = calculateCpp(r.gross, st.gross, st.required, 0, !cppApplies(emp.dob, r.payDate), P, N);
    const required = round2(cppBase + cpp2);
    st.required = round2(st.required + required);
    st.deducted = round2(st.deducted + r.cpp);
    st.gross = round2(st.gross + r.gross);
    out.push({
      sheetRow: r.sheetRow, payDate: r.payDate, employee: r.employee, period: r.period, status: r.status,
      gross: r.gross, periodsCovered: N, exemption: math.exemption,
      deducted: r.cpp, employerRecorded: r.employerCpp, required,
      employerNeedsUpdate: Math.abs(r.employerCpp - required) > 0.005,
      remitted: r.status.toLowerCase() === 'remitted',
    });
  }
  const employeesOut = Object.entries(perEmp).map(([name, st]) => ({
    name, gross: st.gross, required: st.required, deducted: st.deducted,
    shortfall: round2(st.required - st.deducted),   // + = still to recover from the employee
  }));
  return { runs: out, employees: employeesOut };
}

export async function onRequestGet({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  const year = new URL(request.url).searchParams.get('year') || String(new Date().getUTCFullYear());
  try {
    return json({ ok: true, year, ...(await review(env, auth.userId, year)) });
  } catch (err) {
    if (err instanceof SheetReadError) return json(sheetReadErrorBody(err), 503);
    throw err;
  }
}

export async function onRequestPost({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  let body = {}; try { body = await request.json(); } catch { /* empty */ }
  const year = String(body.year || new Date().getUTCFullYear());
  if (!body.apply) return json({ ok: false, error: 'Pass apply:true to write the corrected employer CPP' }, 400);
  try {
    const r = await review(env, auth.userId, year);
    const updated = [], skipped = [];
    for (const run of r.runs) {
      if (!run.employerNeedsUpdate) continue;
      if (run.remitted) { skipped.push({ ...run, reason: 'Already remitted — send the difference to CRA separately' }); continue; }
      const w = await writeRange(env, auth.userId, `'${PAYROLL_TAB}'!S${run.sheetRow}`, [[run.required]]);
      if (!w.ok) return json({ ok: false, error: `Row ${run.sheetRow}: ${w.error}`, updated }, 502);
      updated.push({ sheetRow: run.sheetRow, payDate: run.payDate, from: run.employerRecorded, to: run.required });
    }
    return json({ ok: true, year, updated, skipped, employees: r.employees });
  } catch (err) {
    if (err instanceof SheetReadError) return json(sheetReadErrorBody(err), 503);
    throw err;
  }
}
