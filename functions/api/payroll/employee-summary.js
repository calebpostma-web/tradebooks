// ════════════════════════════════════════════════════════════════════
// GET /api/payroll/employee-summary?employeeId=X&year=YYYY
//
// Per-employee annual summary. Designed for hand-over at CRA audit —
// answers "what did this kid actually do, how many hours, at what rate,
// for whom, and when?" from the contemporaneous Work Log.
//
// Response:
// {
//   ok: true, year, employee: {...},
//   totals: { hours, gross, netPayFromPayroll, deductions, entryCount, payRunCount },
//   tasks: [ { task, hours, gross, count }, ... ],   // top tasks by hours
//   rates: [ { rate, firstSeen, lastSeen, hours, gross } ],
//   businesses: [ { business, hours, gross } ],
//   byMonth: [ { month:'2025-01', hours, gross } ] × 12,
// }
// ════════════════════════════════════════════════════════════════════

import { readRange } from '../../_sheets.js';
import { authenticateRequest, json, options } from '../../_shared.js';

const WORK_LOG_TAB = '📝 Work Log';
const PAYROLL_TAB = '💼 Payroll';

export const onRequestOptions = () => options();

export async function onRequestGet({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  const userId = auth.userId;

  const url = new URL(request.url);
  const employeeId = url.searchParams.get('employeeId');
  let year = parseInt(url.searchParams.get('year'), 10);
  if (!year) year = new Date().getUTCFullYear() - 1;
  if (!employeeId) return json({ ok: false, error: 'employeeId is required' }, 400);

  // Load employee
  const profile = await loadProfile(env, userId);
  const employees = safeParseArray(profile?.employees);
  const employee = employees.find(e => e.id === employeeId);
  if (!employee) return json({ ok: false, error: `Employee ${employeeId} not found` }, 404);

  const yearStart = Date.parse(`${year}-01-01T00:00:00Z`);
  const yearEnd = Date.parse(`${year}-12-31T23:59:59Z`);

  // Work Log entries for this employee in this year
  const wlResult = await readRange(env, userId, `'${WORK_LOG_TAB}'!B12:I1000`);
  const wlEntries = [];
  if (wlResult.ok) {
    for (const row of wlResult.values) {
      if (!row || !row[0]) continue;
      const [date, name, business, task, hours, rate, notes, audit] = row;
      if (name !== employee.name) continue;
      const t = Date.parse(date);
      if (isNaN(t) || t < yearStart || t > yearEnd) continue;
      wlEntries.push({
        date, business: business || 'Postma', task: task || '',
        hours: parseFloat(hours) || 0, rate: parseFloat(rate) || 0,
        notes: notes || '', audit: audit || '',
      });
    }
  }

  // Payroll rows for this employee in this year
  const payResult = await readRange(env, userId, `'${PAYROLL_TAB}'!B12:Q500`);
  const payRuns = [];
  if (payResult.ok) {
    for (const row of payResult.values) {
      if (!row || !row[0]) continue;
      const [payDate, name, , , , , , gross, cpp, ei, fedTax, onTax, netPay, ytdGross, , status] = row;
      if (name !== employee.name) continue;
      if (status && String(status).toLowerCase() === 'cancelled') continue;
      const t = Date.parse(payDate);
      if (isNaN(t) || t < yearStart || t > yearEnd) continue;
      payRuns.push({
        payDate,
        gross: parseFloat(gross) || 0,
        cpp: parseFloat(cpp) || 0,
        ei: parseFloat(ei) || 0,
        fedTax: parseFloat(fedTax) || 0,
        onTax: parseFloat(onTax) || 0,
        netPay: parseFloat(netPay) || 0,
        status: status || '',
      });
    }
  }

  // ── Totals
  const totalHours = round2(wlEntries.reduce((s, e) => s + e.hours, 0));
  const totalGrossFromLog = round2(wlEntries.reduce((s, e) => s + e.hours * e.rate, 0));
  const totalGrossFromPayroll = round2(payRuns.reduce((s, r) => s + r.gross, 0));
  const totalNet = round2(payRuns.reduce((s, r) => s + r.netPay, 0));
  const totalDeductions = round2(payRuns.reduce((s, r) => s + r.cpp + r.ei + r.fedTax + r.onTax, 0));

  // ── Task breakdown — aggregate by task label (or first few words)
  const taskMap = new Map();
  for (const e of wlEntries) {
    const key = (e.task || 'Unlabelled').trim();
    if (!taskMap.has(key)) taskMap.set(key, { task: key, hours: 0, gross: 0, count: 0 });
    const t = taskMap.get(key);
    t.hours += e.hours;
    t.gross += e.hours * e.rate;
    t.count += 1;
  }
  const tasks = [...taskMap.values()]
    .map(t => ({ ...t, hours: round2(t.hours), gross: round2(t.gross) }))
    .sort((a, b) => b.hours - a.hours);

  // ── Rate history — distinct rates with first/last date and usage
  const rateMap = new Map();
  for (const e of wlEntries) {
    const key = e.rate.toFixed(2);
    if (!rateMap.has(key)) rateMap.set(key, { rate: e.rate, firstSeen: e.date, lastSeen: e.date, hours: 0, gross: 0 });
    const r = rateMap.get(key);
    if (e.date < r.firstSeen) r.firstSeen = e.date;
    if (e.date > r.lastSeen) r.lastSeen = e.date;
    r.hours += e.hours;
    r.gross += e.hours * e.rate;
  }
  const rates = [...rateMap.values()]
    .map(r => ({ ...r, hours: round2(r.hours), gross: round2(r.gross) }))
    .sort((a, b) => a.rate - b.rate);

  // ── Business breakdown
  const bizMap = new Map();
  for (const e of wlEntries) {
    const key = e.business || 'Postma';
    if (!bizMap.has(key)) bizMap.set(key, { business: key, hours: 0, gross: 0 });
    const b = bizMap.get(key);
    b.hours += e.hours;
    b.gross += e.hours * e.rate;
  }
  const businesses = [...bizMap.values()]
    .map(b => ({ ...b, hours: round2(b.hours), gross: round2(b.gross) }))
    .sort((a, b) => b.hours - a.hours);

  // ── Monthly breakdown (always 12 rows, zeros for empty months)
  const byMonth = Array.from({ length: 12 }, (_, i) => ({
    month: `${year}-${String(i + 1).padStart(2, '0')}`,
    hours: 0, gross: 0, payRunCount: 0, payRunGross: 0,
  }));
  for (const e of wlEntries) {
    const d = new Date(e.date);
    if (isNaN(d.getTime())) continue;
    const m = d.getUTCMonth();
    byMonth[m].hours += e.hours;
    byMonth[m].gross += e.hours * e.rate;
  }
  for (const r of payRuns) {
    const d = new Date(r.payDate);
    if (isNaN(d.getTime())) continue;
    const m = d.getUTCMonth();
    byMonth[m].payRunCount += 1;
    byMonth[m].payRunGross += r.gross;
  }
  for (const m of byMonth) {
    m.hours = round2(m.hours);
    m.gross = round2(m.gross);
    m.payRunGross = round2(m.payRunGross);
  }

  return json({
    ok: true,
    year,
    employee: {
      id: employee.id,
      name: employee.name,
      dob: employee.dob,
      relationship: employee.relationship,
      startDate: employee.startDate,
      sin: employee.sin || '',
    },
    totals: {
      hours: totalHours,
      grossFromLog: totalGrossFromLog,
      grossFromPayroll: totalGrossFromPayroll,
      net: totalNet,
      deductions: totalDeductions,
      entryCount: wlEntries.length,
      payRunCount: payRuns.length,
    },
    tasks, rates, businesses, byMonth,
    employer: {
      businessName: profile?.business_name || '',
      ownerName: profile?.owner_name || '',
      city: profile?.city || '',
      province: profile?.province || 'ON',
    },
  });
}

// ── Helpers ──

function round2(n) { return Math.round((Number(n) || 0) * 100) / 100; }

function safeParseArray(s) {
  try { const v = JSON.parse(s || '[]'); return Array.isArray(v) ? v : []; } catch { return []; }
}

async function loadProfile(env, userId) {
  try { return await env.DB.prepare('SELECT * FROM profiles WHERE user_id = ?').bind(userId).first(); }
  catch { return null; }
}
