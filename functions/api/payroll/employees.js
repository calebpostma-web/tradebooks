// ════════════════════════════════════════════════════════════════════
// /api/payroll/employees
//
// GET  → returns the employees array from the user's D1 profile.
// POST → upsert one employee. If body.id matches an existing entry,
//        that entry is updated; otherwise a new entry with a generated
//        UUID is appended. Soft-delete via { id, active: false }.
//
// D1 profile.employees is the source of truth. On every write, the
// ⚙️ Config tab's Employees section (B77:I86) is mirrored so the sheet
// reflects the current state for user visibility.
//
// Per-employee schema:
//   {
//     id:              uuid,
//     name:            string,
//     dob:             'YYYY-MM-DD',
//     sin:             string   (optional — required at T4 generation),
//     relationship:    'child' | 'spouse' | 'parent' | 'other',
//     familyEiExempt:  boolean  (defaults true for 'child'|'spouse'),
//     startDate:       'YYYY-MM-DD',
//     defaultRate:     number,
//     td1FedClaim:     number   (1 = basic only),
//     td1OnClaim:      number,
//     active:          boolean  (soft-delete flag),
//     notes:           string   (optional),
//   }
// ════════════════════════════════════════════════════════════════════

import { writeRange } from '../../_sheets.js';
import { authenticateRequest, json, options } from '../../_shared.js';

const CONFIG_TAB = '⚙️ Config';
const EMP_SHEET_RANGE = `'${CONFIG_TAB}'!B77:I86`;
const EMP_SHEET_ROWS = 10;

export const onRequestOptions = () => options();

// ── GET /api/payroll/employees ──────────────────────────────────────
export async function onRequestGet({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);

  const employees = await loadEmployees(env, auth.userId);
  return json({ ok: true, employees });
}

// ── POST /api/payroll/employees — upsert one ────────────────────────
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

  if (!body.name || !body.dob) {
    return json({ ok: false, error: 'name and dob are required' }, 400);
  }

  const relationship = body.relationship || 'child';
  const familyEiExempt = body.familyEiExempt !== undefined
    ? !!body.familyEiExempt
    : (relationship === 'child' || relationship === 'spouse' || relationship === 'parent');

  const incoming = {
    id: body.id || (crypto.randomUUID ? crypto.randomUUID() : generateId()),
    name: String(body.name).trim(),
    dob: String(body.dob).trim(),
    sin: body.sin ? String(body.sin).trim() : '',
    relationship,
    familyEiExempt,
    startDate: body.startDate ? String(body.startDate).trim() : new Date().toISOString().slice(0, 10),
    defaultRate: parseFloat(body.defaultRate) || 0,
    td1FedClaim: parseInt(body.td1FedClaim, 10) || 1,
    td1OnClaim: parseInt(body.td1OnClaim, 10) || 1,
    active: body.active !== false,
    notes: body.notes ? String(body.notes).trim() : '',
    updatedAt: new Date().toISOString(),
  };

  const existing = await loadEmployees(env, userId);
  const idx = existing.findIndex(e => e.id === incoming.id);
  if (idx >= 0) {
    // Preserve createdAt and merge; overwrite with new fields.
    existing[idx] = { ...existing[idx], ...incoming };
  } else {
    incoming.createdAt = incoming.updatedAt;
    existing.push(incoming);
  }

  await saveEmployees(env, userId, existing);

  // Mirror to sheet (best-effort — non-blocking). Only active employees shown.
  try {
    await mirrorToSheet(env, userId, existing);
  } catch (e) {
    console.warn('Sheet mirror failed:', e.message);
  }

  return json({ ok: true, employee: incoming, employees: existing });
}

// ── D1 helpers ──────────────────────────────────────────────────────

async function loadEmployees(env, userId) {
  try {
    const row = await env.DB.prepare('SELECT employees FROM profiles WHERE user_id = ?')
      .bind(userId).first();
    if (!row || !row.employees) return [];
    const parsed = JSON.parse(row.employees);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

async function saveEmployees(env, userId, employees) {
  await env.DB.prepare(
    'UPDATE profiles SET employees = ?, updated_at = datetime(\'now\') WHERE user_id = ?'
  ).bind(JSON.stringify(employees), userId).run();
}

// ── Sheet mirror ────────────────────────────────────────────────────

async function mirrorToSheet(env, userId, employees) {
  const active = employees.filter(e => e.active !== false);
  const rows = [];
  for (let i = 0; i < EMP_SHEET_ROWS; i++) {
    const e = active[i];
    if (e) {
      rows.push([
        e.name || '',
        e.dob || '',
        e.sin || '',
        e.relationship || '',
        e.startDate || '',
        e.defaultRate || '',
        e.td1FedClaim || 1,
        e.td1OnClaim || 1,
      ]);
    } else {
      rows.push(['', '', '', '', '', '', '', '']);
    }
  }
  await writeRange(env, userId, EMP_SHEET_RANGE, rows);
}

// ── Fallback UUID (unused in modern Workers but here for safety) ────

function generateId() {
  return 'emp_' + Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
}
