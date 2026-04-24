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
// Future edits should append `|corrected:<ISO timestamp>|reason:<...>` rather
// than overwriting — that's how we keep the audit trail intact.
// ════════════════════════════════════════════════════════════════════

import { appendRows } from '../../_sheets.js';
import { authenticateRequest, json, options } from '../../_shared.js';

const WORK_LOG_TAB = '📝 Work Log';

export const onRequestOptions = () => options();

export async function onRequestPost({ request, env }) {
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
  const business = body.business || 'Postma';
  const task = String(body.task).trim();
  const notes = body.notes ? String(body.notes).trim() : '';
  const workDate = String(body.workDate).trim();

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
