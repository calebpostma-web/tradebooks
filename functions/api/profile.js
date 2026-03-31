// ════════════════════════════════════════════════════════════════════
// /functions/api/profile.js
// Handles: GET /api/profile (load), PUT /api/profile (save)
//
// Requires JWT auth token in Authorization header
// Required D1 binding: DB
// Required env var: JWT_SECRET
// ════════════════════════════════════════════════════════════════════

import { json, options, authenticateRequest } from '../_shared.js';

export async function onRequestOptions() { return options(); }

// ── GET /api/profile — load the user's business profile ──
export async function onRequestGet(context) {
  const { request, env } = context;

  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Not authenticated' }, 401);

  try {
    const row = await env.DB.prepare(
      'SELECT * FROM profiles WHERE user_id = ?'
    ).bind(auth.userId).first();

    if (!row || !row.business_name) {
      return json({ ok: true, profile: null }); // Onboarding not done
    }

    return json({
      ok: true,
      profile: rowToProfile(row)
    });
  } catch (err) {
    return json({ ok: false, error: err.message }, 500);
  }
}

// ── PUT /api/profile — save/update the user's business profile ──
export async function onRequestPut(context) {
  const { request, env } = context;

  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Not authenticated' }, 401);

  try {
    const profile = await request.json();

    // Validate required fields
    if (!profile.businessName) {
      return json({ ok: false, error: 'Business name is required' }, 400);
    }

    // Upsert profile (INSERT OR REPLACE based on user_id)
    await env.DB.prepare(`
      INSERT INTO profiles (
        user_id, business_name, trading_name, owner_name, email, city, province,
        hst_number, business_type, fiscal_year_end, primary_bank, credit_card,
        invoice_start, home_office_percent, clients, employees, structure,
        activities, sheet_id, script_url, updated_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
      ON CONFLICT(user_id) DO UPDATE SET
        business_name = excluded.business_name,
        trading_name = excluded.trading_name,
        owner_name = excluded.owner_name,
        email = excluded.email,
        city = excluded.city,
        province = excluded.province,
        hst_number = excluded.hst_number,
        business_type = excluded.business_type,
        fiscal_year_end = excluded.fiscal_year_end,
        primary_bank = excluded.primary_bank,
        credit_card = excluded.credit_card,
        invoice_start = excluded.invoice_start,
        home_office_percent = excluded.home_office_percent,
        clients = excluded.clients,
        employees = excluded.employees,
        structure = excluded.structure,
        activities = excluded.activities,
        sheet_id = excluded.sheet_id,
        script_url = excluded.script_url,
        updated_at = datetime('now')
    `).bind(
      auth.userId,
      profile.businessName || '',
      profile.tradingName || '',
      profile.ownerName || '',
      profile.email || '',
      profile.city || '',
      profile.province || 'ON',
      profile.hstNumber || '',
      profile.businessType || 'sole_prop',
      profile.fiscalYearEnd || 'December 31',
      profile.primaryBank || 'BMO',
      profile.creditCard || 'AMEX',
      profile.invoiceStart || 1001,
      profile.homeOfficePercent || 0,
      JSON.stringify(profile.clients || []),
      JSON.stringify(profile.employees || []),
      profile.structure || '',
      profile.activities || '',
      profile.sheetId || '',
      profile.scriptUrl || ''
    ).run();

    return json({ ok: true, message: 'Profile saved' });
  } catch (err) {
    return json({ ok: false, error: err.message }, 500);
  }
}

function rowToProfile(row) {
  return {
    businessName: row.business_name || '',
    tradingName: row.trading_name || '',
    ownerName: row.owner_name || '',
    email: row.email || '',
    city: row.city || '',
    province: row.province || 'ON',
    hstNumber: row.hst_number || '',
    businessType: row.business_type || 'sole_prop',
    fiscalYearEnd: row.fiscal_year_end || 'December 31',
    primaryBank: row.primary_bank || 'BMO',
    creditCard: row.credit_card || 'AMEX',
    invoiceStart: row.invoice_start || 1001,
    homeOfficePercent: row.home_office_percent || 0,
    clients: safeJSON(row.clients, []),
    employees: safeJSON(row.employees, []),
    structure: row.structure || '',
    activities: row.activities || '',
    sheetId: row.sheet_id || '',
    scriptUrl: row.script_url || '',
  };
}

function safeJSON(str, fallback) {
  try { return JSON.parse(str); } catch { return fallback; }
}
