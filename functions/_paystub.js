// ════════════════════════════════════════════════════════════════════
// PAY STUB — data builder + HTML renderer (shared by print and email)
//
// Ontario ESA s.12 wage statement: pay period, wage rate, gross, each
// deduction (amount + purpose), net, and how it was paid. We add the
// employer's own CPP/EI cost and YTD so the employee can reconcile her T4.
// SIN is deliberately NOT printed — it isn't required on a stub and the
// stub gets emailed.
// ════════════════════════════════════════════════════════════════════

import { loadPayrollRows, loadEmployees, loadWorkLogInRange, round2 } from './_payroll_sheet.js';

/** Collect everything the stub needs for one committed Payroll row. */
export async function buildStubData(env, userId, sheetRow) {
  const res = await loadPayrollRows(env, userId);
  if (!res.ok) return { ok: false, error: 'Failed to read Payroll: ' + res.error };
  const run = res.rows.find(r => r.sheetRow === sheetRow);
  if (!run) return { ok: false, error: `No pay run at Payroll row ${sheetRow}` };

  // Employer identity
  let profile = null;
  try { profile = await env.DB.prepare('SELECT business_name, trading_name, owner_name, email, city, province, hst_number FROM profiles WHERE user_id = ?').bind(userId).first(); } catch { /* ignore */ }
  const bnDigits = String(profile?.hst_number || '').replace(/\D/g, '').slice(0, 9);
  const employer = {
    businessName: profile?.business_name || '',
    tradingName: profile?.trading_name || '',
    ownerName: profile?.owner_name || '',
    email: profile?.email || '',
    city: profile?.city || '',
    province: profile?.province || 'ON',
    payrollAccount: bnDigits ? `${bnDigits} RP0001` : '',
  };

  // Employee record (email, start date)
  const employees = await loadEmployees(env, userId);
  const emp = employees.find(e => e.name === run.employee) || {};
  const employee = {
    id: emp.id || '',
    name: run.employee,
    email: emp.email || '',
    startDate: emp.startDate || '',
    relationship: emp.relationship || '',
    td1Fed: emp.td1FedClaim ?? 1,
    td1On: emp.td1OnClaim ?? 1,
  };

  // YTD through this run (same employee, calendar year of payDate, payDate <= this run, not cancelled)
  const year = run.payDate.slice(0, 4);
  const ytd = { gross: 0, cpp: 0, ei: 0, fedTax: 0, onTax: 0, netPay: 0, employerCpp: 0, employerEi: 0, runs: 0 };
  for (const r of res.rows) {
    if (r.employee !== run.employee) continue;
    if (r.status.toLowerCase() === 'cancelled') continue;
    if (!r.payDate.startsWith(year)) continue;
    if (r.payDate > run.payDate) continue;
    if (r.payDate === run.payDate && r.sheetRow > run.sheetRow) continue;
    ytd.gross += r.gross; ytd.cpp += r.cpp; ytd.ei += r.ei; ytd.fedTax += r.fedTax; ytd.onTax += r.onTax;
    ytd.netPay += r.netPay; ytd.employerCpp += r.employerCpp; ytd.employerEi += r.employerEi; ytd.runs++;
  }
  for (const k of Object.keys(ytd)) if (k !== 'runs') ytd[k] = round2(ytd[k]);

  // Period + hours detail
  let periodStart = '', periodEnd = '';
  const pm = /(\d{4}-\d{2}-\d{2})\s*(?:→|->|to|–|-)\s*(\d{4}-\d{2}-\d{2})/.exec(run.period || '');
  if (pm) { periodStart = pm[1]; periodEnd = pm[2]; }
  let workLog = [];
  if (periodStart && periodEnd) {
    workLog = await loadWorkLogInRange(env, userId, run.employee, periodStart, periodEnd);
  }

  return {
    ok: true,
    employer, employee, run, ytd,
    period: { start: periodStart, end: periodEnd, label: run.period || '' },
    workLog,
    stubId: `STUB-${run.payDate.replace(/-/g, '')}-${run.sheetRow}`,
    generatedAt: new Date().toISOString(),
  };
}

// ─── Renderer ───────────────────────────────────────────────────────

const esc = (s) => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const money = (n) => {
  const v = Number(n) || 0;
  const s = Math.abs(v).toLocaleString('en-CA', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return v < 0 ? `($${s})` : `$${s}`;
};
const longDate = (iso) => {
  if (!iso) return '—';
  const d = new Date(iso + (iso.length === 10 ? 'T00:00:00Z' : ''));
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleDateString('en-CA', { year: 'numeric', month: 'long', day: 'numeric', timeZone: 'UTC' });
};

/** Full standalone HTML document (print-ready, email-safe inline styles). */
export function renderPayStub(d) {
  const { employer, employee, run, ytd, period, workLog } = d;
  const legal = employer.businessName || 'Employer';
  const trading = employer.tradingName && employer.tradingName !== employer.businessName ? employer.tradingName : '';
  const payMethod = /cash/i.test(run.workDesc) ? 'Cash' : 'Direct deposit / e-Transfer';

  const dedRows = [
    ['Canada Pension Plan (CPP)', run.cpp, ytd.cpp],
    ['Employment Insurance (EI)', run.ei, ytd.ei],
    ['Federal income tax', run.fedTax, ytd.fedTax],
    ['Ontario income tax', run.onTax, ytd.onTax],
  ].map(([label, amt, y]) => `
      <tr><td>${label}</td><td class="n">${money(amt)}</td><td class="n muted">${money(y)}</td></tr>`).join('');

  const totalDed = round2(run.cpp + run.ei + run.fedTax + run.onTax);
  const ytdDed = round2(ytd.cpp + ytd.ei + ytd.fedTax + ytd.onTax);

  const hoursRows = workLog.length
    ? workLog.map(e => `<tr><td>${esc(e.date)}</td><td>${esc(e.task)}${e.business ? ` <span class="muted">· ${esc(e.business)}</span>` : ''}</td><td class="n">${e.hours.toFixed(2)}</td><td class="n">${money(e.rate)}</td><td class="n">${money(e.hours * e.rate)}</td></tr>`).join('')
    : `<tr><td colspan="5" class="muted">${esc(run.workDesc || '—')}</td></tr>`;

  return `<!DOCTYPE html>
<html lang="en-CA"><head><meta charset="utf-8">
<title>Pay statement — ${esc(employee.name)} — ${esc(run.payDate)}</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
  body{font-family:-apple-system,Segoe UI,Helvetica,Arial,sans-serif;color:#1a1a1a;margin:0;padding:28px;max-width:760px;background:#fff}
  h1{font-size:20px;margin:0 0 2px}
  .sub{color:#666;font-size:13px}
  .hdr{display:flex;justify-content:space-between;gap:24px;border-bottom:2px solid #1a1a1a;padding-bottom:14px;margin-bottom:18px}
  .hdr .right{text-align:right;font-size:13px;line-height:1.5}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:14px 28px;font-size:13px;margin-bottom:18px}
  .lbl{color:#666;font-size:11px;text-transform:uppercase;letter-spacing:.06em}
  table{width:100%;border-collapse:collapse;font-size:13px;margin-bottom:18px}
  th{text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#666;border-bottom:1px solid #999;padding:6px 8px}
  td{padding:6px 8px;border-bottom:1px solid #eee;vertical-align:top}
  td.n,th.n{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
  .muted{color:#777}
  tr.total td{border-top:1px solid #1a1a1a;border-bottom:none;font-weight:600}
  .net{background:#f3f7f4;border:1px solid #cfe3d6;border-radius:6px;padding:12px 16px;display:flex;justify-content:space-between;align-items:baseline;margin:8px 0 20px}
  .net .amt{font-size:24px;font-weight:700}
  .foot{font-size:11px;color:#777;border-top:1px solid #ddd;padding-top:10px;line-height:1.5}
  @media print{body{padding:0} .noprint{display:none}}
</style></head><body>
<div class="hdr">
  <div>
    <h1>${esc(legal)}</h1>
    ${trading ? `<div class="sub">o/a ${esc(trading)}</div>` : ''}
    <div class="sub">${esc([employer.city, employer.province].filter(Boolean).join(', '))}${employer.payrollAccount ? ` · Payroll account ${esc(employer.payrollAccount)}` : ''}</div>
  </div>
  <div class="right">
    <div style="font-size:15px;font-weight:600">Statement of Earnings</div>
    <div>Pay date <strong>${esc(longDate(run.payDate))}</strong></div>
    <div class="muted">${esc(d.stubId)}</div>
  </div>
</div>

<div class="grid">
  <div><div class="lbl">Employee</div><div style="font-size:15px;font-weight:600">${esc(employee.name)}</div>${employee.startDate ? `<div class="muted">Employed since ${esc(longDate(employee.startDate))}</div>` : ''}</div>
  <div><div class="lbl">Pay period</div><div>${period.start ? `${esc(longDate(period.start))} – ${esc(longDate(period.end))}` : esc(period.label || '—')}</div><div class="muted">Paid by ${payMethod}</div></div>
  <div><div class="lbl">Rate</div><div>${run.rate ? `${money(run.rate)} / hour` : '—'} ${run.hours ? `· ${run.hours.toFixed(2)} hours` : ''}</div></div>
  <div><div class="lbl">TD1 claim codes</div><div>Federal ${esc(employee.td1Fed)} · Ontario ${esc(employee.td1On)}</div></div>
</div>

<table>
  <thead><tr><th>Date</th><th>Work</th><th class="n">Hours</th><th class="n">Rate</th><th class="n">Amount</th></tr></thead>
  <tbody>${hoursRows}
  <tr class="total"><td colspan="4">Gross earnings</td><td class="n">${money(run.gross)}</td></tr></tbody>
</table>

<table>
  <thead><tr><th>Deductions</th><th class="n">This pay</th><th class="n">Year to date</th></tr></thead>
  <tbody>${dedRows}
  <tr class="total"><td>Total deductions</td><td class="n">${money(totalDed)}</td><td class="n muted">${money(ytdDed)}</td></tr></tbody>
</table>

<div class="net"><div><div class="lbl">Net pay</div><div class="muted" style="font-size:12px">Gross ${money(run.gross)} − deductions ${money(totalDed)}</div></div><div class="amt">${money(run.netPay)}</div></div>

<table>
  <thead><tr><th>Year to date (${esc(run.payDate.slice(0, 4))})</th><th class="n">Gross</th><th class="n">CPP</th><th class="n">EI</th><th class="n">Income tax</th><th class="n">Net</th></tr></thead>
  <tbody><tr><td>${ytd.runs} pay${ytd.runs === 1 ? '' : 's'} to date</td><td class="n">${money(ytd.gross)}</td><td class="n">${money(ytd.cpp)}</td><td class="n">${money(ytd.ei)}</td><td class="n">${money(ytd.fedTax + ytd.onTax)}</td><td class="n">${money(ytd.netPay)}</td></tr></tbody>
</table>

<table>
  <thead><tr><th>Employer contributions (not deducted from you)</th><th class="n">This pay</th><th class="n">Year to date</th></tr></thead>
  <tbody>
    <tr><td>Employer CPP</td><td class="n">${money(run.employerCpp)}</td><td class="n muted">${money(ytd.employerCpp)}</td></tr>
    <tr><td>Employer EI</td><td class="n">${money(run.employerEi)}</td><td class="n muted">${money(ytd.employerEi)}</td></tr>
  </tbody>
</table>

<div class="foot">
  This statement is issued under the Ontario <em>Employment Standards Act, 2000</em>, s.12. Deductions are remitted to the Canada Revenue Agency under the employer's payroll account and reported on your T4 slip by the last day of February. Keep this statement with your tax records. Questions: ${esc(employer.email || employer.ownerName || legal)}.
</div>
</body></html>`;
}
