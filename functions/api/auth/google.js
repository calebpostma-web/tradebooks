// ════════════════════════════════════════════════════════════════════
// POST /api/auth/google
//
// Sign-in-with-Google for TradeBooks. Replaces email/password as the
// primary auth path — the user signs in with their Google account and
// their Google identity (sub + email) is the TradeBooks account.
//
// Request body:
//   { code, redirectUri }
//
// Flow:
//   1. Exchange the OAuth code for tokens (access + refresh + id_token).
//   2. Decode id_token to extract the Google `sub` (stable user ID),
//      email, name, picture.
//   3. Find user by google_sub. If not found, fall back to email-match
//      (so existing email/password users can sign in with Google and
//      automatically link their account).
//   4. If still not found, create a new user — Google identity only,
//      no password.
//   5. Ensure the profiles row exists; save the refresh token against
//      it for future sheet operations.
//   6. Mint a TradeBooks JWT and return it alongside the access token
//      (so the frontend can create the sheet without a second OAuth).
//
// Required D1 schema: users.google_sub column (migration 0001).
// Required env vars: JWT_SECRET, GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET.
// ════════════════════════════════════════════════════════════════════

import { json, options, createToken } from '../../_shared.js';
import { saveGoogleRefreshToken } from '../../_google.js';

export const onRequestOptions = () => options();

export async function onRequestPost({ request, env }) {
  if (!env.DB) return json({ ok: false, error: 'Database not configured' }, 500);
  if (!env.JWT_SECRET) return json({ ok: false, error: 'JWT secret not configured' }, 500);
  if (!env.GOOGLE_CLIENT_ID || !env.GOOGLE_CLIENT_SECRET) {
    return json({ ok: false, error: 'Google OAuth not configured' }, 500);
  }

  let body;
  try { body = await request.json(); }
  catch { return json({ ok: false, error: 'Invalid JSON' }, 400); }

  const { code, redirectUri } = body;
  if (!code) return json({ ok: false, error: 'code is required' }, 400);
  if (!redirectUri) return json({ ok: false, error: 'redirectUri is required' }, 400);

  // 1. Exchange code for tokens
  const tokenResp = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      code,
      client_id: env.GOOGLE_CLIENT_ID,
      client_secret: env.GOOGLE_CLIENT_SECRET,
      redirect_uri: redirectUri,
      grant_type: 'authorization_code',
    }),
  });
  const tokenData = await tokenResp.json();
  if (tokenData.error) {
    return json({ ok: false, error: tokenData.error_description || tokenData.error }, 400);
  }

  // 2. Decode id_token payload (base64-url JSON). We don't need full
  // signature verification here because we just did the code exchange over
  // TLS to Google directly — if Google's response included this id_token,
  // it's trusted.
  const idToken = tokenData.id_token;
  if (!idToken) return json({ ok: false, error: 'No id_token returned from Google' }, 400);
  const idPayload = decodeJwtPayload(idToken);
  if (!idPayload || !idPayload.sub) {
    return json({ ok: false, error: 'Could not decode Google id_token' }, 400);
  }

  const googleSub = String(idPayload.sub);
  const email = (idPayload.email || '').toLowerCase();
  const name = idPayload.name || idPayload.given_name || email.split('@')[0] || 'User';

  // 3. Find user by google_sub; fall back to email-match (links existing email/password account)
  let user = await env.DB.prepare(
    'SELECT id, email, name, subscription_status, trial_ends_at, google_sub FROM users WHERE google_sub = ? LIMIT 1'
  ).bind(googleSub).first();

  if (!user && email) {
    user = await env.DB.prepare(
      'SELECT id, email, name, subscription_status, trial_ends_at, google_sub FROM users WHERE email = ? LIMIT 1'
    ).bind(email).first();
    if (user) {
      // Link the existing email account to this Google sub.
      await env.DB.prepare('UPDATE users SET google_sub = ? WHERE id = ?').bind(googleSub, user.id).run();
    }
  }

  // 4. Create new user if still not found
  if (!user) {
    const trialEnds = new Date(Date.now() + 14 * 24 * 60 * 60 * 1000).toISOString();
    const ins = await env.DB.prepare(
      `INSERT INTO users (email, name, google_sub, subscription_status, trial_ends_at)
       VALUES (?, ?, ?, 'trial', ?)`
    ).bind(email, name, googleSub, trialEnds).run();
    const userId = ins.meta.last_row_id;
    // Seed profile row so saveGoogleRefreshToken has something to update
    await env.DB.prepare(
      `INSERT INTO profiles (user_id, owner_name, email) VALUES (?, ?, ?)`
    ).bind(userId, name, email).run();
    user = { id: userId, email, name, subscription_status: 'trial', trial_ends_at: trialEnds, google_sub: googleSub };
  }

  // 5. Ensure profile row exists (safety — the email/password register path also does this)
  const hasProfile = await env.DB.prepare('SELECT user_id FROM profiles WHERE user_id = ?').bind(user.id).first();
  if (!hasProfile) {
    await env.DB.prepare(
      `INSERT INTO profiles (user_id, owner_name, email) VALUES (?, ?, ?)`
    ).bind(user.id, user.name || name, user.email || email).run();
  }

  // Save refresh token so future sheet ops can re-auth without the user
  if (tokenData.refresh_token) {
    try {
      await saveGoogleRefreshToken(env, user.id, tokenData.refresh_token, tokenData.access_token, tokenData.expires_in);
    } catch (e) {
      console.warn('saveGoogleRefreshToken failed:', e.message);
    }
  }

  // 6. Load profile for the response so the frontend can route correctly
  const profile = await loadProfile(env, user.id);

  const jwt = await createToken(
    { userId: user.id, email: user.email, name: user.name },
    env.JWT_SECRET
  );

  return json({
    ok: true,
    token: jwt,
    profile,
    user: {
      id: user.id,
      name: user.name,
      email: user.email,
      subscription: checkSubscription(user),
      trialEnds: user.trial_ends_at,
    },
    accessToken: tokenData.access_token,    // short-lived, used for immediate sheet build
    expiresIn: tokenData.expires_in,
    googleProfile: { sub: googleSub, email, name, picture: idPayload.picture || '' },
  });
}

// ─── Helpers ─────────────────────────────────────────────────────────

function decodeJwtPayload(token) {
  try {
    const parts = token.split('.');
    if (parts.length !== 3) return null;
    const b64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
    const padded = b64 + '='.repeat((4 - b64.length % 4) % 4);
    const json = atob(padded);
    return JSON.parse(new TextDecoder().decode(Uint8Array.from(json, c => c.charCodeAt(0))));
  } catch {
    return null;
  }
}

function checkSubscription(user) {
  if (user.subscription_status === 'active') return 'active';
  if (user.subscription_status === 'trial') {
    if (new Date(user.trial_ends_at) > new Date()) return 'trial';
    return 'expired';
  }
  return user.subscription_status || 'expired';
}

async function loadProfile(env, userId) {
  const row = await env.DB.prepare('SELECT * FROM profiles WHERE user_id = ?').bind(userId).first();
  if (!row || !row.business_name) return null;  // triggers onboarding
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
