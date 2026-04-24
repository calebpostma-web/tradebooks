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

    // Determine sign. Parser gives us raw amount (positive) and category tells us direction.
    const rawAmount = Math.abs(parseFloat(row.net) || parseFloat(row.amount) || 0);
    const hstAmount = Math.abs(parseFloat(row.hst) || 0);

    let signedAmount;       // col E
    let finalCategory = cat;
    let matchStatus = 'N/A';
    let relatedInvoice = '';

    if (TRANSFER_CATS.has(cat)) {
      // Amex bill payment from BMO → negative, Internal Transfer, not P&L.
      finalCategory = 'Internal Transfer';
      // Sign based on parser hint: if the row debited (amount negative on source),
      // keep it negative. For transfers between own accounts we use negative by convention
      // (money leaving the observed account).
      signedAmount = -rawAmount;
    } else if (DEFAULT_INCOME_CATS.has(cat)) {
      signedAmount = rawAmount;
      // For deposits, try to propose a match against an open invoice.
      if (cat === 'Income Received' || cat.includes('Revenue') || cat.includes('Income')) {
        const candidate = findInvoiceMatch(openInvoices, row.date, rawAmount + hstAmount, rawAmount);
        if (candidate) {
          proposedMatches.push({
            // Row position in the append batch — resolved to sheet row below after append.
            batchIndex: txnRows.length,
            date: row.date,
            amount: rawAmount,
            vendor: row.vendor,
            invNum: candidate.invNum,
            client: candidate.client,
            invAmount: candidate.total,
            invDate: candidate.dateIssued,
            confidence: candidate.confidence,
          });
          matchStatus = 'Unmatched';  // pending user confirmation
        } else {
          matchStatus = 'Unmatched';
        }
      }
    } else {
      // Expense category
      signedAmount = -rawAmount;
    }

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
 * Load open (Unpaid) invoices from the Invoices tab.
 * Columns B-K: InvNum, Date, Client, Description, Subtotal, HST, Total, HSTFlag, Due, Status
 */
async function loadOpenInvoices(env, userId) {
  const result = await readRange(env, userId, `'${INVOICES_TAB}'!B12:K500`);
  if (!result.ok) return [];
  const invoices = [];
  for (const row of result.values) {
    const [invNum, dateIssued, client, desc, sub, hst, total, hstFlag, due, status] = row;
    if (!invNum || (status && String(status).toLowerCase() !== 'unpaid')) continue;
    invoices.push({
      invNum: String(invNum),
      dateIssued: dateIssued || '',
      client: client || '',
      subtotal: parseFloat(sub) || 0,
      hst: parseFloat(hst) || 0,
      total: parseFloat(total) || 0,
    });
  }
  return invoices;
}

/**
 * Find the best-matching open invoice for a deposit.
 * Priority:
 *   1. Exact match on invoice total (incl HST), within date window.
 *   2. Exact match on invoice subtotal (client paid net of HST? rare but possible), within date window.
 * Returns the candidate or null.
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

    if (approxEqual(depositTotal, inv.total)) {
      candidates.push({ ...inv, confidence: 'high', daysApart });
    } else if (approxEqual(depositNet, inv.total)) {
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
