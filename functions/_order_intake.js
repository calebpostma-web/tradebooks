// ════════════════════════════════════════════════════════════════════
// ORDER INTAKE — book a parsed pctires.ca order into the user's sheet
//
// Shared by the email Worker (workers/order-intake) and the paste endpoint
// (/api/intake/paste). One call does everything the order email tells us:
//
//   🧾 Invoices   one row: PCT-# as the invoice number, customer, lines,
//                 HST, total, status, Stripe payment (if any) in Payments.
//   📒 Transactions
//     • TDG cost  → expense row, Account = the card TDG charges (AMEX),
//                   Match Status "Awaiting statement": the card import
//                   matches it by amount/date like a scanned receipt.
//     • Stripe    → when the site took a deposit / full payment:
//                   revenue row in Account "Stripe" (net + HST share) and
//                   the Stripe fee as an expense in Account "Stripe".
//                   Stripe → BMO payouts are then transfers, not income.
//
// Idempotent: the invoice is keyed on PCT-#, every ledger row on a Source
// Ref, so the same email twice (the site sometimes sends two) books once.
// ════════════════════════════════════════════════════════════════════

import { appendRows, readRange } from './_sheets.js';
import { findInvoice, appendInvoiceRow, deriveStatus, TXN_TAB } from './_invoices.js';
import { round2 } from './_payroll_sheet.js';

export const STRIPE_ACCOUNT = 'Stripe';

function safeJSON(v, fb) { try { const x = JSON.parse(v); return x == null ? fb : x; } catch { return fb; } }

/** Pick the user's own category names for revenue / tire COGS / Stripe fees. */
export async function resolveOrderCategories(env, userId) {
  let row = null;
  try { row = await env.DB.prepare('SELECT custom_expense_cats, custom_income_cats, credit_card, primary_bank FROM profiles WHERE user_id = ?').bind(userId).first(); } catch { /* defaults */ }
  const exp = safeJSON(row?.custom_expense_cats, []);
  const inc = safeJSON(row?.custom_income_cats, []);
  const find = (list, re, fb) => (Array.isArray(list) && list.find(c => re.test(String(c)))) || fb;
  return {
    revenue: find(inc, /tire.*(sale|install)|sales and install/i, find(inc, /sale/i, 'Sales')),
    cogs: find(exp, /purchased.*(customer|inventory)|cogs|cost of goods/i, 'Cost of Goods Sold'),
    stripeFee: find(exp, /stripe/i, find(exp, /bank (charges|fees)|interest & bank/i, 'Bank Charges')),
    card: row?.credit_card || 'AMEX',
    bank: row?.primary_bank || 'BMO',
  };
}

async function existingRefs(env, userId) {
  const r = await readRange(env, userId, `'${TXN_TAB}'!K12:K`);
  return new Set(r.ok ? r.values.map(v => String(v?.[0] || '')).filter(Boolean) : []);
}

/**
 * Book one parsed order. Returns { ok, duplicate?, invoice:{sheetRow,status}, ledger:[{ref, kind}], skipped:[...] }.
 */
export async function bookOrder(env, userId, parsed) {
  if (!parsed?.ok) return { ok: false, error: parsed?.error || 'Unparsed order' };
  const inv = parsed.invoice;
  const found = await findInvoice(env, userId, { invNum: inv.invNum });
  if (!found.ok) return { ok: false, error: found.error };
  if (found.invoice) return { ok: true, duplicate: true, invoice: { sheetRow: found.invoice.sheetRow, status: found.invoice.status }, ledger: [] };

  const cats = await resolveOrderCategories(env, userId);
  const refs = await existingRefs(env, userId);
  const ledger = [];
  const rows = [];
  const dateIso = inv.dateVal || new Date().toISOString().slice(0, 10);

  // ── Stripe money the site already collected ──
  const payments = (inv.payments || []).map(p => ({ ...p }));
  for (const p of payments) {
    if (p.via !== 'stripe' && p.via !== 'affirm') continue;
    const gross = round2(p.amount);
    const hstShare = inv.total > 0 && inv.hstAmt > 0 ? round2(gross * (inv.hstAmt / inv.total)) : 0;
    const net = round2(gross - hstShare);
    const ref = `INV-${inv.invNum}-${p.kind.toUpperCase()}-STRIPE`;
    if (!refs.has(ref)) {
      rows.push([dateIso, inv.client, `Invoice ${inv.invNum} — ${p.kind} via ${p.via === 'affirm' ? 'Affirm/Stripe' : 'Stripe'}`, net, cats.revenue, hstShare > 0 ? 'Yes' : 'No', hstShare, STRIPE_ACCOUNT, 'Order import', ref, inv.invNum, 'N/A']);
      ledger.push({ ref, kind: 'stripe-revenue', net, hst: hstShare });
    }
    if (p.fee > 0) {
      const fref = `INV-${inv.invNum}-${p.kind.toUpperCase()}-STRIPEFEE`;
      if (!refs.has(fref)) {
        rows.push([dateIso, 'Stripe', `Stripe fee — invoice ${inv.invNum}${p.feeEstimated ? ' (estimate — confirm in Stripe)' : ''}`, -round2(p.fee), cats.stripeFee, 'No', 0, STRIPE_ACCOUNT, 'Order import', fref, inv.invNum, 'N/A']);
        ledger.push({ ref: fref, kind: 'stripe-fee', amount: -round2(p.fee), estimated: !!p.feeEstimated });
      }
    }
    p.pending = false; p.ref = ref; p.method = 'card';
  }

  // ── TDG cost, waiting for the card statement ──
  const bill = parsed.expectedBill;
  if (bill && bill.total) {
    const ref = bill.ref;
    if (!refs.has(ref)) {
      rows.push([dateIso, bill.party, bill.description, round2(bill.amount), cats.cogs, bill.hst > 0 ? 'Yes' : 'No', round2(bill.hst), cats.card, 'Order import', ref, inv.invNum, 'Awaiting statement']);
      ledger.push({ ref, kind: 'tdg-cost', amount: round2(bill.amount), hst: round2(bill.hst), account: cats.card });
    }
  }

  // ── The invoice itself ──
  const { status, balanceDue } = deriveStatus({ total: inv.total, payments });
  const a = await appendInvoiceRow(env, userId, {
    invNum: inv.invNum, date: dateIso, client: inv.client, desc: inv.desc,
    sub: inv.sub, hst: inv.hstAmt, total: inv.total, hstFlag: inv.hst, due: parsed.installDate || '',
    status, datePaid: status === 'Paid' ? dateIso : '', notes: inv.notes, category: cats.revenue,
    depositAmount: parsed.charged?.depositPaid || 0, depositDate: parsed.charged?.depositPaid ? dateIso : '',
    balanceDue, lines: inv.lines, payments,
  });
  if (!a.ok) return { ok: false, error: 'Invoice write failed: ' + a.error };

  if (rows.length) {
    const t = await appendRows(env, userId, `'${TXN_TAB}'!B12:M`, rows);
    if (!t.ok) return { ok: false, error: 'Invoice written (row ' + a.sheetRow + ') but ledger write failed: ' + t.error, invoice: { sheetRow: a.sheetRow, status } };
  }

  return { ok: true, invoice: { sheetRow: a.sheetRow, status, balanceDue }, ledger, categories: cats };
}

/** Persist an intake event (never the raw email). */
export async function logIntake(env, { userId = null, to = '', from = '', subject = '', kind, orderNo = null, status, detail = '', payload = null }) {
  try {
    await env.DB.prepare(
      'INSERT INTO intake_events (user_id, to_addr, from_addr, subject, kind, order_no, status, detail, payload) VALUES (?,?,?,?,?,?,?,?,?)'
    ).bind(userId, to, from, String(subject || '').slice(0, 300), kind, orderNo, status, String(detail || '').slice(0, 1000), payload ? JSON.stringify(payload).slice(0, 20000) : null).run();
  } catch (e) { console.warn('intake log failed', e.message); }
}

/** Resolve orders-<token>@domain → userId (or null). */
export async function userForIntakeAddress(env, toAddr) {
  const m = /orders-([a-z0-9]{6,16})@/i.exec(String(toAddr || ''));
  if (!m) return null;
  try {
    const row = await env.DB.prepare('SELECT user_id FROM profiles WHERE intake_token = ?').bind(m[1].toLowerCase()).first();
    return row?.user_id ?? null;
  } catch { return null; }
}
