-- ══════════════════════════════════════════════════════════════════
-- Migration 0003 — "Pocket" categories (transfer accounts by the user's own names)
--
-- Money moved between the business's own pockets (credit-card payment,
-- CRA remittance, shareholder loan, owner draw) is not income or expense.
-- Until now every such row had to be categorized "Internal Transfer" and
-- the user's own name for the pocket was lost. pocket_cats holds the user's
-- names; the sheet's Type column (Transactions!P) classifies rows as
-- Transfer / P&L from it, and every P&L / HST formula keys on that column.
--
-- Run with:
--   npx wrangler d1 execute tradebooks-db --remote --file=./migrations/0003_pocket_cats.sql
-- ══════════════════════════════════════════════════════════════════

ALTER TABLE profiles ADD COLUMN pocket_cats TEXT DEFAULT '[]';
