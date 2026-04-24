// ════════════════════════════════════════════════════════════════════
// GET /api/payroll/t4?year=YYYY
//
// Year-end T4 slip generation. Reads all Payroll rows for the given
// calendar year, groups by employee name, sums them into T4 box values.
// Returns per-employee slip data + T4 Summary totals + a reference CSV
// the user can paste into CRA My Business Account.
//
// Box mapping (2026 T4):
//   Box 14  Employment income (Gross)
//   Box 16  CPP contributions (employee)
//   Box 17  CPP2 contributions (merged into Box 16 here — Caleb's kids
//           never cross YMPE so CPP2 is always $0)
//   Box 18  EI premiums (zero for family-EI-exempt)
//   Box 22  Income tax deducted (Fed + Ontario combined, per T4 conventions)
//   Box 24  EI insurable earnings ($0 for family-EI-exempt)
//   Box 26  CPP pensionable earnings
//   Box 28  CPP / EI / PPIP exemption flags
//   Box 29  Employment code (blank for standard)
//
// Response:
// {
//   ok: true, year, employer: {...},
//   slips: [
//     {
//       employee: { id, name, dob, sin, relationship },
//       ageAtYearStart, cppExemptAllYear, familyEiExempt,
//       box14, box16, box18, box22, box24, box26, box28Cpp, box28Ei,
//       fedTaxWithheld, onTaxWithheld,       // split for reference
//       runs: [ { payDate, gross, cpp, fedTax, onTax } ],
//     },
//   ],
//   summary: {
//     year, totalSlips, totalBox14, totalBox16, totalBox22, totalBox24, totalBox26,
//   },
//   csv: "Year,Employer,Employee,SIN,DOB,Province,..."
// }
// ════════════════════════════════════════════════════════════════════

import { readRange } from '../../_sheets.js';
import { authenticateRequest, json, options } from '../../_shared.js';
import { ageOnDate } from '../../_payroll.js';

const PAYROLL_TAB = '💼 Payroll';

export const onRequestOptions = () => options();

export async function onRequestGet({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  const userId = auth.userId;

  const url = new URL(request.url);
  let year = parseInt(url.searchParams.get('year'), 10);
  if (!year) year = new Date().getUTCFullYear() - 1;  // default: previous calendar year
  if (year < 2020 || year > 2100) return json({ ok: false, error: 'Invalid year' }, 400);

  // Load employer + employees
  const profile = await loadProfile(env, userId);
  const employees = safeParseArray(profile?.employees);
  const employer = {
    businessName: profile?.business_name || '',
    ownerName: profile?.owner_name || '',
    city: profile?.city || '',
    province: profile?.province || 'ON',
    bn: profile?.hst_number || '',   // 9-digit BN is embedded in HST #
  };

  // Read all Payroll rows for the year
  const result = await readRange(env, userId, `'${PAYROLL_TAB}'!B12:Q500`);
  if (!result.ok) return json({ ok: false, error: 'Failed to read Payroll: ' + result.error });

  // Group rows by employee name
  const byName = new Map();
  for (const row of result.values) {
    if (!row || !row[0]) continue;
    const [payDate, empName, , , , , , gross, cpp, ei, fedTax, onTax, netPay, , , status] = row;
    if (!empName) continue;
    if (status && String(status).toLowerCase() === 'cancelled') continue;
    const d = new Date(payDate);
    if (isNaN(d.getTime()) || d.getUTCFullYear() !== year) continue;

    if (!byName.has(empName)) byName.set(empName, []);
    byName.get(empName).push({
      payDate: d.toISOString().slice(0, 10),
      gross: parseFloat(gross) || 0,
      cpp: parseFloat(cpp) || 0,
      ei: parseFloat(ei) || 0,
      fedTax: parseFloat(fedTax) || 0,
      onTax: parseFloat(onTax) || 0,
      netPay: parseFloat(netPay) || 0,
      status: status || '',
    });
  }

  // Build slips
  const slips = [];
  const yearStartIso = `${year}-01-01`;
  const yearEndIso = `${year}-12-31`;

  for (const [empName, runs] of byName.entries()) {
    const emp = employees.find(e => e.name === empName) || {};
    const ageStart = emp.dob ? ageOnDate(emp.dob, yearStartIso) : null;
    const ageEnd = emp.dob ? ageOnDate(emp.dob, yearEndIso) : null;
    const cppExemptAllYear = ageEnd != null && ageEnd < 18;   // under 18 through whole year
    const familyEiExempt = emp.familyEiExempt !== false;

    const box14 = round2(runs.reduce((s, r) => s + r.gross, 0));
    const box16 = round2(runs.reduce((s, r) => s + r.cpp, 0));
    const box18 = round2(runs.reduce((s, r) => s + r.ei, 0));
    const fedTaxTotal = round2(runs.reduce((s, r) => s + r.fedTax, 0));
    const onTaxTotal = round2(runs.reduce((s, r) => s + r.onTax, 0));
    const box22 = round2(fedTaxTotal + onTaxTotal);
    const box24 = familyEiExempt ? 0 : box14;
    // Pensionable = gross, capped at YMPE for CPP-earning employees; zero if CPP-exempt
    const box26 = cppExemptAllYear ? 0 : Math.min(box14, 74600);

    slips.push({
      employee: {
        id: emp.id || '',
        name: empName,
        dob: emp.dob || '',
        sin: emp.sin || '',
        relationship: emp.relationship || '',
      },
      ageAtYearStart: ageStart,
      ageAtYearEnd: ageEnd,
      cppExemptAllYear,
      familyEiExempt,
      box14,
      box16,
      box18,
      box22,
      box24,
      box26,
      box28Cpp: cppExemptAllYear ? 'X' : '',
      box28Ei: familyEiExempt ? 'X' : '',
      fedTaxWithheld: fedTaxTotal,
      onTaxWithheld: onTaxTotal,
      province: employer.province,
      runs,
    });
  }

  slips.sort((a, b) => a.employee.name.localeCompare(b.employee.name));

  const summary = {
    year,
    totalSlips: slips.length,
    totalBox14: round2(slips.reduce((s, t) => s + t.box14, 0)),
    totalBox16: round2(slips.reduce((s, t) => s + t.box16, 0)),
    totalBox18: round2(slips.reduce((s, t) => s + t.box18, 0)),
    totalBox22: round2(slips.reduce((s, t) => s + t.box22, 0)),
    totalBox24: round2(slips.reduce((s, t) => s + t.box24, 0)),
    totalBox26: round2(slips.reduce((s, t) => s + t.box26, 0)),
  };

  const csv = buildCsv(year, employer, slips);

  return json({ ok: true, year, employer, slips, summary, csv });
}

// ── Helpers ──

function round2(n) { return Math.round((Number(n) || 0) * 100) / 100; }

function safeParseArray(s) {
  try { const v = JSON.parse(s || '[]'); return Array.isArray(v) ? v : []; } catch { return []; }
}

async function loadProfile(env, userId) {
  try {
    return await env.DB.prepare('SELECT * FROM profiles WHERE user_id = ?').bind(userId).first();
  } catch {
    return null;
  }
}

function buildCsv(year, employer, slips) {
  const header = [
    'Year', 'Employer', 'Employer BN', 'Employee Name', 'SIN', 'DOB', 'Province',
    'Box 14 Employment Income',
    'Box 16 CPP',
    'Box 18 EI',
    'Box 22 Income Tax (Fed+ON)',
    'Box 24 EI Insurable Earnings',
    'Box 26 CPP Pensionable Earnings',
    'Box 28 CPP Exempt', 'Box 28 EI Exempt',
    'Fed Tax (reference)', 'ON Tax (reference)',
  ].join(',');

  const rows = slips.map(s => [
    year,
    quote(employer.businessName),
    quote(employer.bn),
    quote(s.employee.name),
    quote(s.employee.sin || ''),
    quote(s.employee.dob || ''),
    quote(s.province || 'ON'),
    s.box14.toFixed(2),
    s.box16.toFixed(2),
    s.box18.toFixed(2),
    s.box22.toFixed(2),
    s.box24.toFixed(2),
    s.box26.toFixed(2),
    s.box28Cpp,
    s.box28Ei,
    s.fedTaxWithheld.toFixed(2),
    s.onTaxWithheld.toFixed(2),
  ].join(','));

  // Summary line at bottom
  const sum = (fn) => slips.reduce((t, s) => t + fn(s), 0).toFixed(2);
  const summaryRow = [
    `"SUMMARY (${slips.length} slip${slips.length === 1 ? '' : 's'})"`,
    '', '', '', '', '', '',
    sum(s => s.box14),
    sum(s => s.box16),
    sum(s => s.box18),
    sum(s => s.box22),
    sum(s => s.box24),
    sum(s => s.box26),
    '', '', sum(s => s.fedTaxWithheld), sum(s => s.onTaxWithheld),
  ].join(',');

  return [header, ...rows, summaryRow].join('\n');
}

function quote(s) {
  const str = String(s || '');
  return `"${str.replace(/"/g, '""')}"`;
}
