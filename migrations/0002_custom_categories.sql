-- ══════════════════════════════════════════════════════════════════
-- Migration 0002 — Persist user custom categories
--
-- The Settings screen and the importer's "+ Add new category…" both
-- collected customExpenseCats / customIncomeCats, but profile.js never
-- stored them: the profile row had no column, so every custom category
-- vanished on the next page load. These two JSON-array columns fix that.
--
-- Run with:
--   npx wrangler d1 execute tradebooks-db --remote --file=./migrations/0002_custom_categories.sql
-- ══════════════════════════════════════════════════════════════════

ALTER TABLE profiles ADD COLUMN custom_expense_cats TEXT DEFAULT '[]';
ALTER TABLE profiles ADD COLUMN custom_income_cats  TEXT DEFAULT '[]';
