// ════════════════════════════════════════════════════════════════════
// POST /api/payroll/log-work
//
// Appends one row to the 📝 Work Log tab with a server-side timestamp
// in the Entry Audit column. The timestamp cannot be set client-side —
// this is the contemporaneous-records audit-defence layer.
//
// Request body:
//   {
//     employeeId: string,   // id of the employee from the profile
//     employeeName?: string, // resolved server-side if not provided
//     workDate:   'YYYY-MM-DD',
//     business:   'Postma' | 'HVAC' | 'PCTires',
//     task:       string,
//     hours:      number,
//     rate:       number,
//     notes?:     string,
//   }
//
// Columns written (B–I): Date, Employee, Business, Task, Hours, Rate, Notes, Entry Audit
// Entry Audit format: `entered:<ISO timestamp>|by:<userId>`
// Edits go through /api/payroll/worklog, which appends
// `|corrected:<ISO>|by:|was:|reason:` (or `|removed:…`) rather than overwriting.
//
// Returns 409 { confirm:'duplicate'|'paidPeriod', error } when the entry looks
// like a duplicate or lands in an already-paid period; the client asks, then
// resends with allowDuplicate / allowPaidPeriod set.
// ════════════════════════════════════════════════════════════════════

import { appendRows, readRange } from '../../_sheets.js';
import { authenticateRequest, json, options } from '../../_shared.js';
import {
  WORK_LOG_TAB, WL_COL, WL_READ_RANGE, num, isoDate, isRemovedEntry,
  loadPayrollRows, paidPeriods, findPaidPeriod, SheetReadError, sheetReadErrorBody, safeText, isRealIsoDate,
} from '../../_payroll_sheet.js';

export const onRequestOptions = () => options();

export async function onRequestPost(ctx) {
  try { return await handlePost(ctx); }
  catch (err) {
    if (err instanceof SheetReadError) return json(sheetReadErrorBody(err), 503);
    throw err;
  }
}

function fmtLong(iso) {
  const d = new Date(iso + 'T00:00:00Z');
  return isNaN(d) ? iso : d.toLocaleDateString('en-CA', { month: 'short', day: 'numeric', year: 'numeric', timeZone: 'UTC' });
}

async function handlePost({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  const userId = auth.userId;

  let body;
  try {
    body = await request.json();
  } catch {
    return json({ ok: false, error: 'Invalid JSON' }, 400);
  }

  if (!body.workDate || !body.task || body.hours == null || body.rate == null) {
    return json({ ok: false, error: 'workDate, task, hours, rate are required' }, 400);
  }

  // Resolve employee name if only id given.
  let employeeName = body.employeeName || '';
  if (!employeeName && body.employeeId) {
    try {
      const row = await env.DB.prepare('SELECT employees FROM profiles WHERE user_id = ?')
        .bind(userId).first();
      if (row && row.employees) {
        const emps = JSON.parse(row.employees);
        const match = Array.isArray(emps) && emps.find(e => e.id === body.employeeId);
        if (match) employeeName = match.name;
      }
    } catch { /* fall through with empty name */ }
  }

  const hours = parseFloat(body.hours) || 0;
  const rate = parseFloat(body.rate) || 0;
  const business = safeText(body.business || 'Postma');
  const task = safeText(body.task);
  const notes = body.notes ? safeText(body.notes) : '';
  const workDate = String(body.workDate).trim();
  if (!isRealIsoDate(workDate)) {
    return json({ ok: false, error: 'Work date must be a real date (YYYY-MM-DD).' }, 400);
  }
  if (!(hours > 0)) return json({ ok: false, error: 'Hours must be more than 0.' }, 400);
  if (!employeeName) return json({ ok: false, error: 'Pick an employee.' }, 400);

  // Two checks before writing, each answered with a question the screen
  // shows ("Log it anyway?") rather than a hard stop:
  //  1. Same person, same hours and rate, dated within a week of an existing
  //     live entry → probably entered twice (it happened with the 300 h entry).
  //  2. Date inside a period that's already been paid → these hours won't be
  //     on any pay unless that run is cancelled and redone.
  if (!body.allowDuplicate || !body.allowPaidPeriod) {
    const wl = await readRange(env, userId, WL_READ_RANGE);
    if (!wl.ok) throw new SheetReadError('the Work Log', wl);
    const t = Date.parse(workDate);
    if (!body.allowDuplicate) {
      const dup = wl.values.find(r => r && r[WL_COL.date] && !isRemovedEntry(r) &&
        r[WL_COL.employee] === employeeName &&
        Math.abs(num(r[WL_COL.hours]) - hours) < 0.001 && Math.abs(num(r[WL_COL.rate]) - rate) < 0.001 &&
        Math.abs(Date.parse(isoDate(r[WL_COL.date])) - t) <= 7 * 86400000);
      if (dup) {
        const d = isoDate(dup[WL_COL.date]);
        return json({ ok: false, confirm: 'duplicate',
          existing: { isoDate: d, hours: num(dup[WL_COL.hours]), rate: num(dup[WL_COL.rate]), task: dup[WL_COL.task] || '' },
          error: `There's already an entry for ${employeeName}: ${num(dup[WL_COL.hours])} h at $${num(dup[WL_COL.rate])} on ${fmtLong(d)}` +
            (dup[WL_COL.task] ? ` ("${dup[WL_COL.task]}")` : '') + `. If this is the same work, don't log it again — it would be paid twice.` }, 409);
      }
    }
    if (!body.allowPaidPeriod) {
      const pay = await loadPayrollRows(env, userId);
      if (!pay.ok) throw new SheetReadError('past pay runs (Payroll tab)', pay);
      const p = findPaidPeriod(paidPeriods(pay.rows, employeeName), workDate);
      if (p) {
        return json({ ok: false, confirm: 'paidPeriod', paid: { payDate: p.payDate, start: p.start, end: p.end },
          error: `${fmtLong(workDate)} is inside a pay that's already been run (the ${fmtLong(p.payDate)} pay covered ` +
            `${fmtLong(p.start)} to ${fmtLong(p.end)}). Hours dated there won't be on any pay unless that run is cancelled and redone. ` +
            `If this is new work, date it ${fmtLong(new Date(Date.parse(p.end) + 86400000).toISOString().slice(0, 10))} or later.` }, 409);
      }
    }
  }

  // Server-stamped audit string — client cannot forge this.
  const entryAudit = `entered:${new Date().toISOString()}|by:${userId}`;

  const row = [[
    workDate,
    employeeName,
    business,
    task,
    hours,
    rate,
    notes,
    entryAudit,
  ]];

  const result = await appendRows(env, userId, `'${WORK_LOG_TAB}'!B12:I`, row);
  if (!result.ok) return json({ ok: false, error: 'Work Log write failed: ' + result.error });

  return json({
    ok: true,
    workDate, employeeName, business, task, hours, rate,
    grossForRow: Math.round(hours * rate * 100) / 100,
    entryAudit,
  });
}
