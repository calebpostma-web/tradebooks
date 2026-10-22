// ════════════════════════════════════════════════════════════════════
// Email Worker: *@aibookkeeper.ca → book the order
//
// Flow per message:
//   1. Who is it for?  orders-<token>@aibookkeeper.ca → profiles.intake_token
//   2. What is it?
//        • Gmail's forwarding-confirmation email → store the code so the
//          user can read it in Settings and confirm the filter
//        • a pctires.ca "New Order" email (direct, or forwarded by Gmail —
//          Gmail keeps the original From and adds X-Forwarded-For)
//        • anything else → logged as ignored
//   3. Parse → book (invoice + ledger rows) → log the outcome.
//
// Nothing is ever bounced or replied to. The raw email is not stored.
// ════════════════════════════════════════════════════════════════════

import PostalMime from 'postal-mime';
import { parseOrderEmail } from '../../../functions/_pctires_order.js';
import { bookOrder, logIntake, userForIntakeAddress } from '../../../functions/_order_intake.js';

const ORDER_SENDERS = /@pctires\.ca$/i;

export default {
  async email(message, env, ctx) {
    const to = message.to || '';
    const envelopeFrom = message.from || '';
    let parsedMail;
    try {
      parsedMail = await PostalMime.parse(message.raw);
    } catch (e) {
      await logIntake(env, { to, from: envelopeFrom, subject: '', kind: 'error', status: 'error', detail: 'MIME parse failed: ' + e.message });
      return;
    }
    const subject = parsedMail.subject || '';
    const headerFrom = parsedMail.from?.address || envelopeFrom;
    // Gmail forwarding: original sender stays in From; the forwarding mailbox is in X-Forwarded-For
    const fwdFor = (parsedMail.headers || []).find(h => /^x-forwarded-for$/i.test(h.key))?.value || '';
    const userId = await userForIntakeAddress(env, to);

    // ── Gmail forwarding confirmation ──
    if (/forwarding confirmation/i.test(subject) || /forwarding-noreply@google\.com/i.test(headerFrom)) {
      // Gmail puts the code in the subject: "(#123456789) Gmail Forwarding Confirmation - Receive Mail from …"
      // and the body is often HTML-only, so look in subject, text, then de-tagged HTML.
      const text = parsedMail.text || '';
      const htmlText = String(parsedMail.html || '').replace(/<[^>]+>/g, ' ');
      const code = (/\(#?(\d{6,12})\)/.exec(subject) || /confirmation code:?\s*(\d{6,12})/i.exec(text + ' ' + htmlText) || /\b(\d{9})\b/.exec(text + ' ' + htmlText) || [])[1] || '';
      const link = (/https:\/\/mail-settings\.google\.com\/mail\/[^\s)>"']+/i.exec(text + ' ' + htmlText) || [])[0] || '';
      await logIntake(env, { userId, to, from: headerFrom, subject, kind: 'gmail_verify', status: 'info', detail: code || '(no code found)', payload: { link } });
      return;
    }

    if (!userId) {
      await logIntake(env, { to, from: headerFrom, subject, kind: 'ignored', status: 'needs_review', detail: 'No account matches this intake address' });
      return;
    }

    // ── Order email? ──
    const looksLikeOrder = /new order/i.test(subject) || /PCT-[A-Z0-9]{4,}/.test(subject) || ORDER_SENDERS.test(headerFrom);
    if (!looksLikeOrder) {
      await logIntake(env, { userId, to, from: headerFrom, subject, kind: 'ignored', status: 'info', detail: 'Not an order email' + (fwdFor ? ` (forwarded by ${fwdFor})` : '') });
      return;
    }
    // Customer-facing "Order Confirmed" copies carry no cost data — skip them
    if (/order confirmed/i.test(subject) && !/new order/i.test(subject)) {
      await logIntake(env, { userId, to, from: headerFrom, subject, kind: 'ignored', status: 'info', detail: 'Customer confirmation copy — the New Order email is the one we book' });
      return;
    }

    const parsed = parseOrderEmail({ html: parsedMail.html || '', text: parsedMail.text || '', subject });
    if (!parsed.ok) {
      await logIntake(env, { userId, to, from: headerFrom, subject, kind: 'order', status: 'needs_review', detail: 'Could not parse: ' + parsed.error });
      return;
    }
    if (!(parsed.charged?.total > 0) || !parsed.customer?.name) {
      await logIntake(env, { userId, to, from: headerFrom, subject, kind: 'order', orderNo: parsed.orderNo, status: 'needs_review', detail: 'Parsed but incomplete (missing total or customer) — paste it into the app to book manually', payload: parsed.charged });
      return;
    }

    const result = await bookOrder(env, userId, parsed);
    await logIntake(env, {
      userId, to, from: headerFrom, subject, kind: 'order', orderNo: parsed.orderNo,
      status: !result.ok ? 'error' : result.duplicate ? 'duplicate' : 'booked',
      detail: result.ok
        ? (result.duplicate ? `Already in Invoices (row ${result.invoice.sheetRow})` : `Invoices row ${result.invoice.sheetRow} · ${result.ledger.map(l => l.kind).join(', ') || 'no ledger rows'}`)
        : result.error,
      payload: { orderNo: parsed.orderNo, client: parsed.customer.name, total: parsed.charged.total, status: parsed.status, tdg: parsed.tdg.totalCost, stripe: parsed.charged.depositPaid || (parsed.charged.paidInFull ? parsed.charged.total : 0) },
    });
  },

  // Health check so the Worker URL answers something sensible
  async fetch() {
    return new Response('aibookkeeper order intake — email only', { status: 200 });
  },
};
