// ════════════════════════════════════════════════════════════════════
// POST /api/import/statement
// Import parsed bank statement rows into the 📒 Transactions tab.
//
// Phase 2A cash-basis behaviour:
// - All rows written to 📒 Transactions with SIGNED amounts.
// - Amex bill payments from BMO → category "Internal Transfer" (excluded from P&L / HST).
// - BMO deposits matched against open invoices (amount exact, date within 14d)
//   and proposed matches returned. Nothing is auto-confirmed — user approves
//   via /api/import/confirm-match.
// ════════════════════════════════════════════════════════════════════

import { appendRows, writeRange, readRange, readExistingRefs, generateRef } from '../../_sheets.js';
import { authenticateRequest, json, options } from '../../_shared.js';

const TXN_TAB = '📒 Transactions';
const INVOICES_TAB = '🧾 Invoices';

const DEFAULT_INCOME_CATS = new Set([
  'Consulting Revenue', 'Service Revenue', 'Sales Revenue',
  'Sales Revenue — Materials', 'Rental Income', 'Interest Income',
  'Income Received', 'Other Income',
]);

// Categories that map directly to Internal Transfer on import.
const TRANSFER_CATS = new Set([
  'Bill Payment / Transfer',   // legacy category name from parsers
  'Internal Transfer',
]);

// Categories to skip entirely (never write to ledger).
const SKIP_CATS = new Set([
  'SKIP — not a business expense',
  'Owner Draw / Distribution',
]);

// Match window: deposit date must be within this many days of the invoice date.
const MATCH_DATE_WINDOW_DAYS = 14;

export const onRequestOptions = () => options();

export async function onRequestPost({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  const userId = auth.userId;

  let body;
  try {
    body = await request.json();
  } catch {
    return json({ ok: false, error: 'Invalid JSON' }, 400);
  }

  const rows = body.rows || [];
  const bank = body.bank || 'AMEX';
  const source = body.source || 'Import';

  if (!rows.length) return json({ ok: false, error: 'No rows' }, 400);

  // Dedup: read existing refs from the Ref column (K) of Transactions.
  const existingRefs = await readExistingRefs(env, userId, TXN_TAB, 'K');

  // Load open invoices up-front so we can propose matches without per-row reads.
  const openInvoices = await loadOpenInvoices(env, userId);

  const txnRows = [];
  const batchCounter = {};
  const proposedMatches = [];
  const duplicateDetails = [];
  let skipped = 0, duplicates = 0;

  for (const row of rows) {
    const cat = row.category || '';
    if (SKIP_CATS.has(cat)) {
      skipped++;
      continue;
    }

    // Ref for dedup — batch counter so N same-vendor-same-amount on same day all survive
    const base = generateRef(bank, row.date, row.amount || row.net, row.vendor);
    const baseLC = base.toLowerCase();
    batchCounter[baseLC] = (batchCounter[baseLC] || 0) + 1;
    const ref = batchCounter[baseLC] > 1 ? `${base}-${batchCounter[baseLC]}` : base;
    const refLC = ref.toLowerCase();

    if (existingRefs.has(refLC)) {
      duplicates++;
      duplicateDetails.push({ date: row.date, vendor: row.vendor, amount: row.amount });
      continue;
    }
    existingRefs.add(refLC);

    // Sign comes from the bank statement (positive = money in, negative = money out)
    // — NOT from the category. The front-end has already extracted HST and signed
    // `net` correctly based on the parsed direction. Pre-fix, the server overrode
    // the sign based on a hardcoded DEFAULT_INCOME_CATS set, which broke for any
    // user-defined income category (e.g. "Customer Payment of Invoice"): those
    // landed as expenses with negative amounts. Trusting the bank-direction sign
    // makes user-custom categories Just Work and is more robust generally.
    const signedNet = parseFloat(row.net) || parseFloat(row.amount) || 0;
    const rawAmount = Math.abs(signedNet);
    const hstAmount = Math.abs(parseFloat(row.hst) || 0);

    let signedAmount = signedNet;
    let finalCategory = cat;
    let matchStatus = 'N/A';
    let relatedInvoice = '';

    if (TRANSFER_CATS.has(cat)) {
      // Internal transfers keep their bank-direction sign (BMO->AMEX is -, AMEX
      // receiving the same is +). The matching pair offsets in P&L because both
      // sides are excluded by the Internal Transfer filter.
      finalCategory = 'Internal Transfer';
      // signedAmount already = signedNet from bank
    } else if (signedNet > 0) {
      // Money in. Try to propose an invoice match regardless of which specific
      // income category the user picked (custom categories like "Customer
      // Payment of Invoice" should match too).
      const candidate = findInvoiceMatch(openInvoices, row.date, rawAmount + hstAmount, rawAmount);
      if (candidate) {
        proposedMatches.push({
          batchIndex: txnRows.length,
          date: row.date,
          amount: rawAmount,
          vendor: row.vendor,
          invNum: candidate.invNum,
          client: candidate.client,
          invAmount: candidate.expected,
          invTotal: candidate.total,
          invDate: candidate.dateIssued,
          confidence: candidate.confidence,
          nextLeg: candidate.nextLeg,
        });
        matchStatus = 'Unmatched';  // pending user confirmation
      } else {
        matchStatus = 'Unmatched';
      }
    }
    // Else: expense, signedAmount already negative from signedNet

    // Row layout (B-M): Date | Party | Description | Amount | Category | HST? | HST | Account | Source | Ref | Related Invoice | Match Status
    txnRows.push([
      row.date || '',
      row.vendor || '',
      row.description || '',
      signedAmount,
      finalCategory,
      hstAmount > 0 ? 'Yes' : 'No',
      hstAmount,
      bank,
      source,
      ref,
      relatedInvoice,
      matchStatus,
    ]);
  }

  let firstAppendedRow = null;
  if (txnRows.length) {
    const appendResult = await appendRows(env, userId, `'${TXN_TAB}'!B12:M`, txnRows);
    if (!appendResult.ok) return json({ ok: false, error: 'Transactions write failed: ' + appendResult.error });

    const match = /!B(\d+):M(\d+)/.exec(appendResult.updates.updatedRange);
    if (match) firstAppendedRow = parseInt(match[1]);
  }

  // Resolve batch indexes to real sheet rows for the proposed-match payload.
  const resolvedMatches = (firstAppendedRow != null)
    ? proposedMatches.map(m => ({ ...m, txnRow: firstAppendedRow + m.batchIndex, batchIndex: undefined }))
    : [];

  return json({
    ok: true,
    written: txnRows.length,
    skipped,
    duplicates,
    duplicatesBlocked: duplicates,
    duplicateDetails,
    total: rows.length,
    proposedMatches: resolvedMatches,
  });
}

// ── Invoice matching ─────────────────────────────────────────────────

/**
 * Load open invoices from the Invoices tab — anything that isn't fully Paid.
 * Columns B-Q: InvNum, Date, Client, Description, Subtotal, HST, Total, HSTFlag,
 *              Due, Status, DatePaid, Notes, RevenueCat, DepositAmount,
 *              DepositDate, BalanceDue
 *
 * For each open invoice we expose `expected` — the dollar amount we're hoping
 * to match against an incoming deposit:
 *   - 'Awaiting Deposit'   → expected = deposit amount        (nextLeg = 'deposit')
 *   - 'Deposit Received'   → expected = balance due           (nextLeg = 'final')
 *   - 'Unpaid' (no dep)    → expected = total                 (nextLeg = 'final')
 *   - 'Unpaid' (deposit set but date blank — legacy edge case) → treat as Awaiting Deposit
 */
async function loadOpenInvoices(env, userId) {
  // Read out to col Q so the deposit columns are included. Sheets returns ragged
  // rows (trailing empty cells stripped), so destructure with defaults.
  const result = await readRange(env, userId, `'${INVOICES_TAB}'!B12:Q500`);
  if (!result.ok) return [];
  const invoices = [];
  for (const row of result.values) {
    const invNum = row[0];
    const dateIssued = row[1] || '';
    const client = row[2] || '';
    const sub = row[4];
    const hst = row[5];
    const total = parseFloat(row[6]) || 0;
    const status = (row[9] || '').toString().trim();
    const depositAmount = parseFloat(row[13]) || 0;
    const depositDate = (row[14] || '').toString().trim();
    const balanceDue = parseFloat(row[15]) || 0;

    if (!invNum) continue;
    const sLower = status.toLowerCase();
    if (sLower === 'paid' || sLower === 'cancelled') continue;

    let expected, nextLeg;
    if (sLower === 'deposit received') {
      // Deposit already in the books; we're now looking for the final-balance payment.
      expected = balanceDue > 0 ? balanceDue : Math.max(0, total - depositAmount);
      nextLeg = 'final';
    } else if (sLower === 'awaiting deposit' || (depositAmount > 0 && !depositDate)) {
      // Deposit hasn't landed yet — match against the deposit amount.
      expected = depositAmount;
      nextLeg = 'deposit';
    } else {
      // Plain unpaid invoice (legacy or no deposit configured).
      expected = total;
      nextLeg = 'final';
    }

    invoices.push({
      invNum: String(invNum),
      dateIssued,
      client,
      subtotal: parseFloat(sub) || 0,
      hst: parseFloat(hst) || 0,
      total,
      depositAmount,
      depositDate,
      balanceDue,
      status,
      expected,
      nextLeg, // 'deposit' or 'final' — passed through to confirm-match
    });
  }
  return invoices;
}

/**
 * Find the best-matching open invoice for a deposit.
 * The "expected amount" depends on the invoice's current state — for invoices
 * with status 'Deposit Received' we're looking for the balance, not the total.
 * Date window is anchored to the issue date, which works for both legs since
 * deposits and final payments both typically arrive within a few weeks.
 *
 * Priority:
 *   1. Exact match on invoice expected amount (incl HST), within date window.
 *   2. Exact match on net (HST-stripped) version of expected, within date window.
 *   3. Exact match on subtotal (legacy fallback).
 */
function findInvoiceMatch(openInvoices, depositDateStr, depositTotal, depositNet) {
  if (!openInvoices.length) return null;
  const depositDate = parseDate(depositDateStr);
  if (!depositDate) return null;

  const candidates = [];
  for (const inv of openInvoices) {
    const invDate = parseDate(inv.dateIssued);
    if (!invDate) continue;
    const daysApart = Math.abs((depositDate - invDate) / (1000 * 60 * 60 * 24));
    if (daysApart > MATCH_DATE_WINDOW_DAYS) continue;

    if (approxEqual(depositTotal, inv.expected) || approxEqual(depositNet, inv.expected)) {
      candidates.push({ ...inv, confidence: 'high', daysApart });
    } else if (approxEqual(depositTotal, inv.subtotal) || approxEqual(depositNet, inv.subtotal)) {
      candidates.push({ ...inv, confidence: 'medium', daysApart });
    }
  }
  if (!candidates.length) return null;
  // Prefer highest confidence, then closest date.
  candidates.sort((a, b) => {
    const rank = { high: 0, medium: 1, low: 2 };
    if (rank[a.confidence] !== rank[b.confidence]) return rank[a.confidence] - rank[b.confidence];
    return a.daysApart - b.daysApart;
  });
  return candidates[0];
}

function approxEqual(a, b, tolerance = 0.01) {
  return Math.abs(a - b) < tolerance;
}

function parseDate(s) {
  if (!s) return null;
  // Accept ISO "yyyy-mm-dd", "mm/dd/yyyy", or Sheets-rendered "mmm d, yyyy".
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  return null;
}
