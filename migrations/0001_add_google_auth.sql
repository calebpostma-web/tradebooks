-- ══════════════════════════════════════════════════════════════════
-- Migration 0001 — Add Google OAuth identity to users
--
-- Adds google_sub (Google's stable user ID from the id_token) so users
-- can sign in with Google without an email/password pair.
--
-- Password is now optional: a row with google_sub set but password_hash
-- NULL represents a Google-only account.
--
-- Run with:
--   npx wrangler d1 execute tradebooks --file=./migrations/0001_add_google_auth.sql
-- Or for local dev:
--   npx wrangler d1 execute tradebooks --local --file=./migrations/0001_add_google_auth.sql
-- ══════════════════════════════════════════════════════════════════

ALTER TABLE users ADD COLUMN google_sub TEXT;
CREATE INDEX IF NOT EXISTS idx_users_google_sub ON users(google_sub);
