-- ══════════════════════════════════════════════════════════════════
-- Migration 0004 — payroll remitter type
--
-- CRA assigns each payroll (RP) account a remitter type independent of the
-- HST filing frequency. 'monthly' (regular) = due the 15th of the month
-- after the pay date. 'quarterly' = due the 15th of the month after the
-- calendar quarter in which the pay date falls (Apr/Jul/Oct/Jan 15).
-- Postma Contracting files nil remittances quarterly → 'quarterly'.
--
-- Run with:
--   npx wrangler d1 execute tradebooks-db --remote --file=./migrations/0004_payroll_remitter.sql
-- ══════════════════════════════════════════════════════════════════

ALTER TABLE profiles ADD COLUMN payroll_remitter TEXT DEFAULT 'monthly';
