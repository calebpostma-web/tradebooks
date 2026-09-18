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

import { appendRows, writeRange, readRange, batchUpdate, generateRef } from '../../_sheets.js';
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

  // The user's own pocket names (Settings → Pockets) are transfer categories
  // too: kept verbatim in the Category column, excluded from P&L by the
  // sheet's Type column. Legacy names still collapse to "Internal Transfer".
  let pockets = new Set();
  try {
    const prow = await env.DB.prepare('SELECT pocket_cats FROM profiles WHERE user_id = ?').bind(userId).first();
    (JSON.parse(prow?.pocket_cats || '[]') || []).forEach(p => { if (p) pockets.add(String(p).trim().toLowerCase()); });
  } catch (e) { /* no pockets configured */ }
  const isTransfer = c => TRANSFER_CATS.has(c) || pockets.has(String(c || '').trim().toLowerCase());

  if (!rows.length) return json({ ok: false, error: 'No rows' }, 400);

  // One read of the ledger (B:O) serves both dedup (Source Ref, col K) and
  // receipt ↔ statement matching (1b-4). Values come back FORMATTED, so money
  // is "($41.04)" and dates are "Aug 31, 2026" — parsed below.
  const ledger = await readRange(env, userId, `'${TXN_TAB}'!B12:O`);
  if (!ledger.ok) return json({ ok: false, error: 'Could not read Transactions: ' + ledger.error }, 500);
  const existingRefs = new Set();
  const ledgerRows = [];
  ledger.values.forEach((v, i) => {
    const ref = String(v[9] || '').trim();
    if (ref) existingRefs.add(ref.toLowerCase());
    const amount = parseMoney(v[3]);
    if (!v[0] || amount == null) return;
    ledgerRows.push({
      row: 12 + i,
      date: parseDate(v[0]),
      party: String(v[1] || ''),
      amount,
      hst: Math.abs(parseMoney(v[6]) || 0),
      category: String(v[4] || '').trim(),
      account: String(v[7] || '').trim(),
      source: String(v[8] || '').trim(),
      ref, refLC: ref.toLowerCase(),
      status: String(v[11] || '').trim(),
      receipt: String(v[13] || '').trim(),
    });
  });
  const claimed = new Set();     // ledger rows already matched in this batch
  const updates = [];            // in-place cell writes (values:batchUpdate)
  const merged = [];             // { date, vendor, amount, row, how }
  const cell = (col, row, val) => updates.push({ range: `'${TXN_TAB}'!${col}${row}`, values: [[val]] });

  // Load open invoices up-front so we can propose matches without per-row reads.
  const openInvoices = await loadOpenInvoices(env, userId);

  const txnRows = [];
  const batchCounter = {};
  const proposedMatches = [];
  const duplicateDetails = [];
  const receiptUrls = [];   // parallel to txnRows → column O (Receipt link)
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
      // Rescanned receipt: still attach the photo if the existing row has none.
      const rcptUrl = String(row.receiptUrl || '').trim();
      if (rcptUrl) {
        const existing = ledgerRows.find(r => r.refLC === refLC && !r.receipt);
        if (existing) { cell('O', existing.row, rcptUrl); existing.receipt = rcptUrl; }
      }
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
    const receiptUrl = String(row.receiptUrl || '').trim();
    // Receipt Scanner rows: the card statement hasn't arrived yet. 'Awaiting
    // statement' lets the later import find and merge them (1b-4) and the
    // Checks tab flag the ones that never show up. Cash never gets a statement.
    const fromReceipt = source === 'Receipt Scanner';

    if (isTransfer(cat)) {
      // Internal transfers keep their bank-direction sign (BMO->AMEX is -, AMEX
      // receiving the same is +). The matching pair offsets in P&L because both
      // sides are excluded by the Type column. A pocket name stays as typed.
      finalCategory = TRANSFER_CATS.has(cat) ? 'Internal Transfer' : cat;
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

    if (fromReceipt && !isTransfer(cat)) {
      matchStatus = /^cash$/i.test(bank) ? 'N/A' : 'Awaiting statement';
    }

    // ── 1b-4: merge receipt and statement rows instead of writing both ──
    // Same account, same total (±2¢), dates within MATCH_DAYS. Vendor text is
    // deliberately ignored — the AI and the bank never spell it the same way.
    if (signedNet < 0 && !isTransfer(cat)) {
      const total = rawAmount + hstAmount;
      if (fromReceipt) {
        // Same receipt scanned twice with the AI spelling the vendor differently
        // ("barBURRITO" vs "barBURRITO Chatham Grand") → different ref → slipped
        // past dedup. Same account + same total + same day + shared vendor stem
        // = the same piece of paper.
        const twin = findReceiptTwin(ledgerRows, { account: bank, total, date: row.date, vendor: row.vendor });
        if (twin) {
          duplicates++;
          duplicateDetails.push({ date: row.date, vendor: row.vendor, amount: row.amount });
          if (receiptUrl && !twin.receipt) { cell('O', twin.row, receiptUrl); twin.receipt = receiptUrl; }
          continue;
        }
        // Receipt arriving after the statement: attach to the statement row.
        const hit = findLedgerMatch(ledgerRows, claimed, { account: bank, total, date: row.date, wantReceiptRow: false });
        if (hit) {
          claimed.add(hit.row);
          if (finalCategory) cell('F', hit.row, finalCategory);   // she saw the receipt — her category wins
          cell('M', hit.row, 'Receipt ✓');
          if (receiptUrl) cell('O', hit.row, receiptUrl);
          merged.push({ date: row.date, vendor: row.vendor, amount: total, row: hit.row, how: 'receipt→statement', party: hit.party });
          continue;
        }
      } else {
        // Statement arriving after the receipt: claim the receipt row.
        const hit = findLedgerMatch(ledgerRows, claimed, { account: bank, total, date: row.date, wantReceiptRow: true });
        if (hit) {
          claimed.add(hit.row);
          cell('K', hit.row, ref);            // statement ref becomes the dedup key
          cell('M', hit.row, 'Receipt ✓');
          merged.push({ date: row.date, vendor: row.vendor, amount: total, row: hit.row, how: 'statement→receipt', party: hit.party });
          continue;
        }
      }
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
    receiptUrls.push(receiptUrl);
  }

  let firstAppendedRow = null;
  if (txnRows.length) {
    const appendResult = await appendRows(env, userId, `'${TXN_TAB}'!B12:M`, txnRows);
    if (!appendResult.ok) return json({ ok: false, error: 'Transactions write failed: ' + appendResult.error });

    const match = /!B(\d+):M(\d+)/.exec(appendResult.updates.updatedRange);
    if (match) firstAppendedRow = parseInt(match[1]);

    // Column O = link to the receipt photo in Drive. Written separately because
    // N is an ARRAYFORMULA and must never be touched by an append.
    if (firstAppendedRow != null && receiptUrls.some(Boolean)) {
      const oRange = `'${TXN_TAB}'!O${firstAppendedRow}:O${firstAppendedRow + receiptUrls.length - 1}`;
      const oRes = await writeRange(env, userId, oRange, receiptUrls.map(u => [u]));
      if (!oRes.ok) console.warn('Receipt link column write failed:', oRes.error);
    }
  }

  if (updates.length) {
    const upd = await batchUpdate(env, userId, updates);
    if (!upd.ok) return json({ ok: false, error: 'Ledger update failed: ' + upd.error, written: txnRows.length });
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
    merged: merged.length,
    mergedDetails: merged,
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
  const result = await readRange(env, userId, `'${INVOICES_TAB}'!B12:Q`);
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

const MATCH_DAYS = 5;

// "($41.04)" → -41.04 ; "$1,234.50" → 1234.5 ; "" → null
function parseMoney(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return v;
  const str = String(v).trim();
  const neg = /^\(.*\)$/.test(str) || /^-/.test(str);
  const n = parseFloat(str.replace(/[^0-9.]/g, ''));
  if (isNaN(n)) return null;
  return neg ? -n : n;
}

// Find the ledger row a receipt/statement line should merge into.
// wantReceiptRow=true  → looking for a scanned receipt still "Awaiting statement"
// wantReceiptRow=false → looking for a statement row that has no receipt yet
function findLedgerMatch(ledgerRows, claimed, { account, total, date, wantReceiptRow }) {
  const d = parseDate(date);
  if (!d || !(total > 0)) return null;
  const acct = String(account || '').toLowerCase();
  let best = null;
  for (const r of ledgerRows) {
    if (claimed.has(r.row) || !r.date || r.amount >= 0) continue;
    if (r.account.toLowerCase() !== acct) continue;
    const isOpenReceipt = r.source === 'Receipt Scanner' && r.status === 'Awaiting statement';
    if (wantReceiptRow) { if (!isOpenReceipt) continue; }
    else { if (r.source === 'Receipt Scanner' || r.status === 'Receipt ✓' || r.receipt) continue; }
    if (Math.abs((Math.abs(r.amount) + r.hst) - total) > 0.02) continue;
    const days = Math.abs((r.date - d) / 86400000);
    if (days > MATCH_DAYS) continue;
    if (!best || days < best.days) best = { ...r, days };
  }
  return best;
}

// Another scan of the same receipt already in the ledger (any status).
function findReceiptTwin(ledgerRows, { account, total, date, vendor }) {
  const d = parseDate(date);
  if (!d || !(total > 0)) return null;
  const acct = String(account || '').toLowerCase();
  const stem = String(vendor || '').replace(/[^a-z0-9]/gi, '').toUpperCase().slice(0, 6);
  for (const r of ledgerRows) {
    if (r.source !== 'Receipt Scanner' || r.amount >= 0 || !r.date) continue;
    if (r.account.toLowerCase() !== acct) continue;
    if (Math.abs((Math.abs(r.amount) + r.hst) - total) > 0.02) continue;
    if (Math.abs((r.date - d) / 86400000) > 0.5) continue;
    const rStem = r.party.replace(/[^a-z0-9]/gi, '').toUpperCase().slice(0, 6);
    if (stem && rStem && stem !== rStem) continue;
    return r;
  }
  return null;
}

function parseDate(s) {
  if (!s) return null;
  // Accept ISO "yyyy-mm-dd", "mm/dd/yyyy", or Sheets-rendered "mmm d, yyyy".
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  return null;
}
