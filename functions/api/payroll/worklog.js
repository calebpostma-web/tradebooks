// ════════════════════════════════════════════════════════════════════
// /api/payroll/worklog — list, correct and remove Work Log entries
//
// GET  ?employeeId=<id>&limit=25
//   → { ok, entries:[{ sheetRow, date, isoDate, employee, business, task,
//        hours, rate, notes, removed, corrected, paidBy:{payDate,start,end}|null }] }
//   Newest entry first. `paidBy` is set when a committed (non-cancelled)
//   pay run's period covers the entry's date — those can't be changed here.
//
// POST { action:'update'|'remove', sheetRow,
//        expected:{ employee, isoDate, hours },        // what the screen showed
//        workDate?, business?, task?, hours?, rate?, notes?,   // update only
//        reason }                                       // required, for the trail
//
// Nothing is ever deleted. The Work Log is the contemporaneous record CRA
// asks for when a family member is on payroll, so every change is visible:
//   update → row rewritten, Entry Audit gets
//            "|corrected:<ISO>|by:<user>|was:<date>;<h>h@<rate>|reason:<…>"
//   remove → Hours set to 0 (so sheet totals stay right), Notes prefixed
//            "[REMOVED]", Entry Audit gets "|removed:<ISO>|by:<user>|was:…|reason:…".
//            Every payroll reader skips rows whose audit has "|removed:".
//
// Moving an entry INTO an already-paid period needs allowPaidPeriod (the
// screen asks first) — for hours that pay really covered but were mis-dated.
//
// Inside a period that's already been paid, only changes that leave the paid
// hours alone go through: rewording task/notes, or removing an exact duplicate.
//
// Entries inside a period that's already been paid are refused with the
// fix spelled out: cancel that pay run first, correct the hours, redo the pay.
// ════════════════════════════════════════════════════════════════════

import { readRange, writeRange } from '../../_sheets.js';
import { authenticateRequest, json, options } from '../../_shared.js';
import {
  WORK_LOG_TAB, WL_COL, WL_READ_RANGE, num, round2, isoDate,
  loadPayrollRows, paidPeriods, findPaidPeriod, isRemovedEntry,
  SheetReadError, sheetReadErrorBody, safeText, isRealIsoDate,
} from '../../_payroll_sheet.js';

export const onRequestOptions = () => options();

export async function onRequestGet(ctx) {
  try { return await handleGet(ctx); }
  catch (err) {
    if (err instanceof SheetReadError) return json(sheetReadErrorBody(err), 503);
    throw err;
  }
}

export async function onRequestPost(ctx) {
  try { return await handlePost(ctx); }
  catch (err) {
    if (err instanceof SheetReadError) return json(sheetReadErrorBody(err), 503);
    throw err;
  }
}

async function loadEmployeesList(env, userId) {
  try {
    const row = await env.DB.prepare('SELECT employees FROM profiles WHERE user_id = ?').bind(userId).first();
    const emps = row?.employees ? JSON.parse(row.employees) : [];
    return Array.isArray(emps) ? emps : [];
  } catch { return []; }
}

function fmtPaid(p) {
  return p ? { payDate: p.payDate, start: p.start, end: p.end } : null;
}

async function handleGet({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  const userId = auth.userId;
  const url = new URL(request.url);
  const employeeId = url.searchParams.get('employeeId') || '';
  const limit = Math.min(200, Math.max(1, parseInt(url.searchParams.get('limit'), 10) || 25));

  let employeeName = '';
  if (employeeId) {
    const emp = (await loadEmployeesList(env, userId)).find(e => e.id === employeeId);
    if (!emp) return json({ ok: false, error: `Employee ${employeeId} not found` }, 404);
    employeeName = emp.name;
  }

  const wl = await readRange(env, userId, WL_READ_RANGE);
  if (!wl.ok) throw new SheetReadError('the Work Log', wl);
  const pay = await loadPayrollRows(env, userId);
  if (!pay.ok) throw new SheetReadError('past pay runs (Payroll tab)', pay);

  const periodsByName = new Map();
  const entries = [];
  for (let i = wl.values.length - 1; i >= 0 && entries.length < limit; i--) {
    const row = wl.values[i];
    if (!row || !row[WL_COL.date]) continue;
    const name = row[WL_COL.employee] || '';
    if (employeeName && name !== employeeName) continue;
    if (!periodsByName.has(name)) periodsByName.set(name, paidPeriods(pay.rows, name));
    const audit = String(row[WL_COL.audit] || '');
    const removed = isRemovedEntry(row);
    entries.push({
      sheetRow: 12 + i,
      date: row[WL_COL.date],
      isoDate: isoDate(row[WL_COL.date]),
      employee: name,
      business: row[WL_COL.business] || '',
      task: row[WL_COL.task] || '',
      hours: num(row[WL_COL.hours]),
      rate: num(row[WL_COL.rate]),
      notes: row[WL_COL.notes] || '',
      removed,
      corrected: /\|corrected:/.test(audit),
      // last "was:<date>;<h>h@<rate>" — what the row said before the latest change
      was: (() => { const m = [...audit.matchAll(/\|was:([^;|]*);([\d.]+)h@([\d.]+)/g)].pop(); return m ? { isoDate: m[1], hours: +m[2], rate: +m[3] } : null; })(),
      paidBy: removed ? null : fmtPaid(findPaidPeriod(periodsByName.get(name), row[WL_COL.date])),
    });
  }
  return json({ ok: true, entries });
}

function cleanReason(v) {
  return String(v ?? '').replace(/[|\r\n]+/g, ' ').trim().slice(0, 200);
}
function fmtLong(iso) {
  const d = new Date(iso + 'T00:00:00Z');
  return isNaN(d) ? iso : d.toLocaleDateString('en-CA', { month: 'short', day: 'numeric', year: 'numeric', timeZone: 'UTC' });
}
function paidMessage(p, what) {
  return `${what} inside a pay that's already been run — the ${fmtLong(p.payDate)} pay, which covered ` +
    `${fmtLong(p.start)} to ${fmtLong(p.end)}. Changing hours after they've been paid would make the ` +
    `Work Log disagree with the pay stub. If that pay hasn't been remitted yet: Pay Runs → Pay stubs → ` +
    `✕ Cancel on that run, then fix the hours here, then run the pay again.`;
}

async function handlePost({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  const userId = auth.userId;

  let body;
  try { body = await request.json(); } catch { return json({ ok: false, error: 'Invalid JSON' }, 400); }
  const action = body.action;
  const sheetRow = parseInt(body.sheetRow, 10);
  if (!['update', 'remove'].includes(action)) return json({ ok: false, error: "action must be 'update' or 'remove'" }, 400);
  if (!(sheetRow >= 12)) return json({ ok: false, error: 'sheetRow is required' }, 400);
  const reason = cleanReason(body.reason);
  if (reason.length < 3) return json({ ok: false, error: 'Please say briefly why (e.g. "wrong date" or "entered twice"). It goes in the audit trail.' }, 400);

  const range = `'${WORK_LOG_TAB}'!B${sheetRow}:I${sheetRow}`;
  const cur = await readRange(env, userId, range);
  if (!cur.ok) throw new SheetReadError('the Work Log', cur);
  const row = (cur.values && cur.values[0]) || [];
  if (!row[WL_COL.date]) return json({ ok: false, error: 'That Work Log row is empty now. Refresh the list and try again.' }, 409);

  // Make sure the row is still the one the screen showed (someone may have
  // sorted or inserted rows in the sheet since).
  const exp = body.expected || {};
  const rowIso = isoDate(row[WL_COL.date]);
  if ((exp.employee && exp.employee !== row[WL_COL.employee]) ||
      (exp.isoDate && exp.isoDate !== rowIso) ||
      (exp.hours != null && Math.abs(num(exp.hours) - num(row[WL_COL.hours])) > 0.001)) {
    return json({ ok: false, stale: true, error: 'That row in the Google Sheet has changed since this list was loaded. Refresh the list and try again.' }, 409);
  }
  if (isRemovedEntry(row)) return json({ ok: false, error: 'That entry was already removed.' }, 409);

  const employee = row[WL_COL.employee] || '';
  const pay = await loadPayrollRows(env, userId);
  if (!pay.ok) throw new SheetReadError('past pay runs (Payroll tab)', pay);
  const periods = paidPeriods(pay.rows, employee);

  const oldHours = num(row[WL_COL.hours]);
  const oldRate = num(row[WL_COL.rate]);

  // Inside an already-paid period, only changes that leave the paid hours
  // alone are allowed: wording (task/notes/business), or removing an exact
  // duplicate (same employee, date, hours and rate on another live row).
  const paidOld = findPaidPeriod(periods, row[WL_COL.date]);
  if (paidOld) {
    let allowed = false;
    if (action === 'update') {
      const sameDate = body.workDate == null || String(body.workDate).trim() === rowIso;
      const sameHours = body.hours == null || Math.abs(parseFloat(body.hours) - oldHours) < 0.001;
      const sameRate = body.rate == null || Math.abs(parseFloat(body.rate) - oldRate) < 0.001;
      allowed = sameDate && sameHours && sameRate;
    } else {
      const all = await readRange(env, userId, WL_READ_RANGE);
      if (!all.ok) throw new SheetReadError('the Work Log', all);
      allowed = all.values.some((r, i) => 12 + i !== sheetRow && r && !isRemovedEntry(r) &&
        r[WL_COL.employee] === employee && isoDate(r[WL_COL.date]) === rowIso &&
        Math.abs(num(r[WL_COL.hours]) - oldHours) < 0.001 && Math.abs(num(r[WL_COL.rate]) - oldRate) < 0.001);
    }
    if (!allowed) {
      return json({ ok: false, paid: fmtPaid(paidOld), error: paidMessage(paidOld, 'These hours are') +
        (action === 'update' ? ' (You can still change the task wording or notes without cancelling.)' : '') }, 409);
    }
  }
  const now = new Date().toISOString();
  const was = `was:${rowIso || row[WL_COL.date]};${oldHours}h@${oldRate}`;
  const audit = String(row[WL_COL.audit] || '');

  let newRow;
  if (action === 'remove') {
    newRow = [
      rowIso || row[WL_COL.date], employee, row[WL_COL.business] || '', row[WL_COL.task] || '',
      0, oldRate,
      '[REMOVED] ' + String(row[WL_COL.notes] || '').replace(/^\[REMOVED\]\s*/, ''),
      `${audit}|removed:${now}|by:${userId}|${was}|reason:${reason}`,
    ];
  } else {
    const workDate = String(body.workDate ?? rowIso).trim();
    if (!isRealIsoDate(workDate)) {
      return json({ ok: false, error: 'Work date must be a real date (YYYY-MM-DD).' }, 400);
    }
    const hours = body.hours == null ? oldHours : parseFloat(body.hours);
    const rate = body.rate == null ? oldRate : parseFloat(body.rate);
    if (!(hours > 0) || hours > 744) return json({ ok: false, error: 'Hours must be more than 0. To take an entry out completely, use Remove.' }, 400);
    if (!(rate >= 0)) return json({ ok: false, error: 'Rate must be 0 or more.' }, 400);
    const task = body.task == null ? (row[WL_COL.task] || '') : safeText(body.task);
    if (!String(task).trim()) return json({ ok: false, error: 'Task description is required.' }, 400);
    // Moving hours INTO a pay that's already been run is legitimate when that
    // pay really covered them (Andrea's 300 h were paid in the Sept catch-up
    // but dated on the pay day, Oct 4). Ask first; the screen resends with
    // allowPaidPeriod after the person confirms.
    const paidNew = findPaidPeriod(periods, workDate);
    if (paidNew && !paidOld && !body.allowPaidPeriod) {
      return json({ ok: false, confirm: 'paidPeriod', paid: fmtPaid(paidNew),
        error: `${fmtLong(workDate)} is inside the ${fmtLong(paidNew.payDate)} pay, which covered ${fmtLong(paidNew.start)} to ${fmtLong(paidNew.end)}. ` +
          `Only do this if that pay already paid for these hours (for example, they were dated on the pay day by mistake). ` +
          `If these hours haven't been paid yet, pick a date after ${fmtLong(paidNew.end)} instead.` }, 409);
    }
    newRow = [
      workDate, employee,
      body.business == null ? (row[WL_COL.business] || '') : safeText(body.business),
      task, hours, rate,
      body.notes == null ? (row[WL_COL.notes] || '') : safeText(body.notes),
      `${audit}|corrected:${now}|by:${userId}|${was}|reason:${reason}`,
    ];
  }

  const w = await writeRange(env, userId, range, [newRow]);
  if (!w.ok) return json({ ok: false, error: 'Work Log write failed: ' + w.error });
  return json({
    ok: true, action, sheetRow,
    entry: { isoDate: String(newRow[0]), employee, business: newRow[2], task: newRow[3], hours: newRow[4], rate: newRow[5], notes: newRow[6], gross: round2(num(newRow[4]) * num(newRow[5])) },
  });
}
