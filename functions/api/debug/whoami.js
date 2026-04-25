// ════════════════════════════════════════════════════════════════════
// GET /api/debug/whoami
//
// Diagnostic endpoint — returns the server-side state for the authenticated
// user. Used to figure out why specific accounts can't establish a server-side
// Google connection (the "no google connection" loop).
//
// SAFETY: read-only, no writes. Returns booleans / counts for sensitive
// fields (we never echo back the actual refresh_token value).
// ════════════════════════════════════════════════════════════════════

import { authenticateRequest, json, options } from '../../_shared.js';

export const onRequestOptions = () => options();

export async function onRequestGet({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);

  // Pull the user row + the profile row separately. Either can be missing in
  // edge cases (user without a profile = onboarding not done).
  const userRow = await env.DB.prepare(
    'SELECT id, email, name, google_sub, subscription_status, trial_ends_at, created_at FROM users WHERE id = ?'
  ).bind(auth.userId).first();

  const profileRow = await env.DB.prepare(
    `SELECT business_name, owner_name, email, sheet_id,
            google_refresh_token, google_access_token, google_token_expires_at,
            updated_at
       FROM profiles WHERE user_id = ?`
  ).bind(auth.userId).first();

  const now = Math.floor(Date.now() / 1000);

  return json({
    ok: true,
    jwt: {
      userId: auth.userId,
      email: auth.email || null,
    },
    user: userRow ? {
      id: userRow.id,
      email: userRow.email,
      name: userRow.name,
      googleSubLinked: !!userRow.google_sub,
      // Only first 6 chars of google_sub so we can spot-check matches without
      // exposing the full ID
      googleSubPrefix: userRow.google_sub ? String(userRow.google_sub).slice(0, 6) + '…' : null,
      subscriptionStatus: userRow.subscription_status,
      trialEndsAt: userRow.trial_ends_at,
      createdAt: userRow.created_at,
    } : { found: false, note: 'No row in users table for this JWT user_id' },
    profile: profileRow ? {
      businessName: profileRow.business_name || null,
      ownerName: profileRow.owner_name || null,
      email: profileRow.email || null,
      sheetId: profileRow.sheet_id || null,
      sheetIdLooksValid: !!(profileRow.sheet_id && profileRow.sheet_id.length > 20),
      // Boolean — never expose the actual token. THIS is the field the
      // "no google connection" debugging hinges on.
      hasRefreshToken: !!profileRow.google_refresh_token,
      hasAccessToken: !!profileRow.google_access_token,
      accessTokenExpiresAt: profileRow.google_token_expires_at || null,
      accessTokenSecondsRemaining: profileRow.google_token_expires_at
        ? (profileRow.google_token_expires_at - now)
        : null,
      profileUpdatedAt: profileRow.updated_at,
    } : { found: false, note: 'No row in profiles table for this user_id — onboarding not done' },
    diagnosis: makeDiagnosis(userRow, profileRow, now),
  });
}

// Plain-language interpretation so the user doesn't have to read JSON.
// Each rule covers one concrete failure mode the user might be hitting.
function makeDiagnosis(userRow, profileRow, now) {
  const notes = [];
  if (!userRow) notes.push('❌ No user row — JWT references a user_id that no longer exists. Sign out and back in.');
  if (!profileRow) notes.push('❌ No profile row — onboarding never completed. Run setup.');
  if (profileRow && !profileRow.sheet_id) notes.push('⚠ No Google Sheet connected. Build one or paste a sheet ID into Settings.');
  if (profileRow && profileRow.sheet_id && !profileRow.google_refresh_token) {
    notes.push("❌ THIS is the 'No Google connection' bug — sheet_id is set but no refresh_token saved. Either Google didn't return one during OAuth, or save failed silently. Check tradebooks OAuth verification status with Google.");
  }
  if (profileRow && profileRow.google_refresh_token && profileRow.google_access_token && profileRow.google_token_expires_at && profileRow.google_token_expires_at < now) {
    notes.push('⚠ Access token expired — server should auto-refresh on next call. If it doesn\'t, refresh_token may be invalid.');
  }
  if (notes.length === 0) notes.push('✓ Server-side state looks healthy.');
  return notes;
}
