// ════════════════════════════════════════════════════════════════════
// GOOGLE AUTH HELPER
// Manages refresh tokens and cached access tokens for Sheets API calls
// ════════════════════════════════════════════════════════════════════

/**
 * Get a valid Google access token for a user. Uses cached token if still valid,
 * otherwise refreshes via the refresh_token.
 * Returns: { ok: true, accessToken } | { ok: false, error, needsReauth }
 */
export async function getGoogleAccessToken(env, userId) {
    const profile = await env.DB.prepare(
          'SELECT google_refresh_token, google_access_token, google_token_expires_at FROM profiles WHERE user_id = ?'
        ).bind(userId).first();

  if (!profile || !profile.google_refresh_token) {
        return { ok: false, error: 'No Google connection', needsReauth: true };
  }

  const now = Math.floor(Date.now() / 1000);

  // Return cached access token if still valid (with 60s buffer)
  if (profile.google_access_token && profile.google_token_expires_at > now + 60) {
        return { ok: true, accessToken: profile.google_access_token };
  }

  // Refresh the access token
  const tokenResp = await fetch('https://oauth2.googleapis.com/token', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
                client_id: env.GOOGLE_CLIENT_ID,
                client_secret: env.GOOGLE_CLIENT_SECRET,
                refresh_token: profile.google_refresh_token,
                grant_type: 'refresh_token',
        }),
  });

  const tokenData = await tokenResp.json();

  if (tokenData.error) {
        // If refresh token is invalid, flag for re-auth
      if (tokenData.error === 'invalid_grant') {
              return { ok: false, error: 'Google access revoked. Please reconnect.', needsReauth: true };
      }
        return { ok: false, error: tokenData.error_description || tokenData.error };
  }

  // Cache the new access token
  const expiresAt = now + (tokenData.expires_in || 3600);
    await env.DB.prepare(
          'UPDATE profiles SET google_access_token = ?, google_token_expires_at = ? WHERE user_id = ?'
        ).bind(tokenData.access_token, expiresAt, userId).run();

  return { ok: true, accessToken: tokenData.access_token };
}

/**
 * Save refresh token for a user (called during OAuth exchange).
 */
export async function saveGoogleRefreshToken(env, userId, refreshToken, accessToken, expiresIn) {
    const expiresAt = Math.floor(Date.now() / 1000) + (expiresIn || 3600);
    await env.DB.prepare(
          'UPDATE profiles SET google_refresh_token = ?, google_access_token = ?, google_token_expires_at = ? WHERE user_id = ?'
        ).bind(refreshToken, accessToken, expiresAt, userId).run();
}

/**
 * Get user's sheet_id from profile.
 */
export async function getUserSheetId(env, userId) {
    const profile = await env.DB.prepare(
          'SELECT sheet_id FROM profiles WHERE user_id = ?'
        ).bind(userId).first();
    return profile?.sheet_id || null;
}
