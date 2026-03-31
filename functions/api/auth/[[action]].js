// ════════════════════════════════════════════════════════════════════
// /functions/api/auth/[[action]].js
// Handles: POST /api/auth/register, POST /api/auth/login
//
// Cloudflare Pages Function — catches /api/auth/register and /api/auth/login
// 
// Required env vars:
//   JWT_SECRET — a random string for signing tokens (generate with: openssl rand -hex 32)
//
// Required D1 binding:
//   DB — bound to your TradeBooks D1 database
// ════════════════════════════════════════════════════════════════════

import {
  json, options, CORS,
  hashPassword, verifyPassword, createToken
} from '../../_shared.js';

export async function onRequestOptions() { return options(); }

export async function onRequestPost(context) {
  const { request, env, params } = context;
  const action = params.action?.[0]; // 'register' or 'login'

  if (!env.DB) return json({ ok: false, error: 'Database not configured' }, 500);
  if (!env.JWT_SECRET) return json({ ok: false, error: 'JWT secret not configured' }, 500);

  try {
    const body = await request.json();

    if (action === 'register') {
      return await handleRegister(body, env);
    } else if (action === 'login') {
      return await handleLogin(body, env);
    } else {
      return json({ ok: false, error: `Unknown action: ${action}` }, 400);
    }
  } catch (err) {
    return json({ ok: false, error: err.message }, 500);
  }
}

// ── REGISTER ──
async function handleRegister(body, env) {
  const { name, email, password } = body;

  if (!name || !email || !password) {
    return json({ ok: false, error: 'Name, email, and password are required' }, 400);
  }
  if (password.length < 6) {
    return json({ ok: false, error: 'Password must be at least 6 characters' }, 400);
  }
  if (!email.includes('@') || !email.includes('.')) {
    return json({ ok: false, error: 'Please enter a valid email address' }, 400);
  }

  // Check if user already exists
  const existing = await env.DB.prepare('SELECT id FROM users WHERE email = ?').bind(email.toLowerCase()).first();
  if (existing) {
    return json({ ok: false, error: 'An account with this email already exists. Try logging in.' }, 409);
  }

  // Hash password
  const passwordHash = await hashPassword(password);

  // Calculate trial end (14 days from now)
  const trialEnds = new Date(Date.now() + 14 * 24 * 60 * 60 * 1000).toISOString();

  // Insert user
  const result = await env.DB.prepare(
    `INSERT INTO users (email, name, password_hash, subscription_status, trial_ends_at) 
     VALUES (?, ?, ?, 'trial', ?)`
  ).bind(email.toLowerCase(), name, passwordHash, trialEnds).run();

  const userId = result.meta.last_row_id;

  // Create empty profile
  await env.DB.prepare(
    `INSERT INTO profiles (user_id, owner_name, email) VALUES (?, ?, ?)`
  ).bind(userId, name, email.toLowerCase()).run();

  // Generate JWT
  const token = await createToken({ userId, email: email.toLowerCase(), name }, env.JWT_SECRET);

  return json({
    ok: true,
    token,
    profile: null, // No business profile yet — triggers onboarding
    user: { id: userId, name, email: email.toLowerCase(), subscription: 'trial', trialEnds }
  });
}

// ── LOGIN ──
async function handleLogin(body, env) {
  const { email, password } = body;

  if (!email || !password) {
    return json({ ok: false, error: 'Email and password are required' }, 400);
  }

  // Find user
  const user = await env.DB.prepare(
    'SELECT id, email, name, password_hash, subscription_status, trial_ends_at FROM users WHERE email = ?'
  ).bind(email.toLowerCase()).first();

  if (!user) {
    return json({ ok: false, error: 'No account found with that email' }, 401);
  }

  // Verify password
  const valid = await verifyPassword(password, user.password_hash);
  if (!valid) {
    return json({ ok: false, error: 'Incorrect password' }, 401);
  }

  // Check subscription status
  const subStatus = checkSubscription(user);

  // Load profile
  const profile = await loadProfile(user.id, env);

  // Generate JWT
  const token = await createToken(
    { userId: user.id, email: user.email, name: user.name },
    env.JWT_SECRET
  );

  return json({
    ok: true,
    token,
    profile, // null if onboarding not completed
    user: {
      id: user.id,
      name: user.name,
      email: user.email,
      subscription: subStatus,
      trialEnds: user.trial_ends_at
    }
  });
}

// ── Subscription check ──
function checkSubscription(user) {
  if (user.subscription_status === 'active') return 'active';
  if (user.subscription_status === 'trial') {
    if (new Date(user.trial_ends_at) > new Date()) return 'trial';
    return 'expired';
  }
  return user.subscription_status || 'expired';
}

// ── Load profile (returns null if not configured) ──
async function loadProfile(userId, env) {
  const row = await env.DB.prepare(
    'SELECT * FROM profiles WHERE user_id = ?'
  ).bind(userId).first();

  if (!row || !row.business_name) return null; // Onboarding not done

  return {
    businessName: row.business_name,
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
