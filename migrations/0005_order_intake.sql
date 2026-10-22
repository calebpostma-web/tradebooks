-- ══════════════════════════════════════════════════════════════════
-- Migration 0005 — email order intake
--
-- Each account gets a private intake address (orders-<token>@aibookkeeper.ca).
-- Cloudflare Email Routing (catch-all) hands every message on the domain to
-- the order-intake Worker, which maps the token → user and books the order.
-- intake_events is the audit trail + "needs a look" list shown in the app
-- (parse failures, duplicates, Gmail forwarding-confirmation codes).
--
-- Run with:
--   npx wrangler d1 execute tradebooks-db --remote --file=./migrations/0005_order_intake.sql
-- ══════════════════════════════════════════════════════════════════

ALTER TABLE profiles ADD COLUMN intake_token TEXT;
CREATE UNIQUE INDEX IF NOT EXISTS idx_profiles_intake_token ON profiles(intake_token);

CREATE TABLE IF NOT EXISTS intake_events (
  id            INTEGER PRIMARY KEY AUTOINCREMENT,
  user_id       INTEGER,                      -- NULL when the token didn't resolve
  received_at   TEXT NOT NULL DEFAULT (datetime('now')),
  to_addr       TEXT,
  from_addr     TEXT,
  subject       TEXT,
  kind          TEXT NOT NULL,                -- 'order' | 'gmail_verify' | 'ignored' | 'error'
  order_no      TEXT,                         -- PCT-XXXXXX when parsed
  status        TEXT NOT NULL,                -- 'booked' | 'duplicate' | 'needs_review' | 'error' | 'info'
  detail        TEXT,                         -- human-readable note / error / verification code
  payload       TEXT                          -- parsed order JSON (no raw email body)
);
CREATE INDEX IF NOT EXISTS idx_intake_events_user ON intake_events(user_id, received_at DESC);
