// ════════════════════════════════════════════════════════════════════
// POST /api/payroll/stub-email
//
// Body: { row: number, to?: string }
//
// 1. Builds + renders the stub for Payroll row N (same HTML as Print).
// 2. Saves a copy to the user's Drive: "AI Bookkeeper Payroll / <year> /
//    <Employee> — <payDate>.html" and writes the link into Payroll col U.
// 3. Emails it (HTML body + .html attachment) via Resend to the employee's
//    address on file (or `to`), cc the business email.
//
// Needs the RESEND_API_KEY secret and a verified sending domain. Sender
// defaults to payroll@aibookkeeper.ca; override with PAYROLL_FROM.
// ════════════════════════════════════════════════════════════════════

import { authenticateRequest, json, options } from '../../_shared.js';
import { getGoogleAccessToken } from '../../_google.js';
import { writeRange } from '../../_sheets.js';
import { PAYROLL_TAB } from '../../_payroll_sheet.js';
import { buildStubData, renderPayStub } from '../../_paystub.js';
import { uploadToDrive } from '../../_drive.js';

export const onRequestOptions = () => options();

export async function onRequestPost({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  const userId = auth.userId;

  let body;
  try { body = await request.json(); } catch { return json({ ok: false, error: 'Invalid JSON' }, 400); }
  const row = parseInt(body.row, 10);
  if (!row || row < 12) return json({ ok: false, error: 'row is required' }, 400);

  const data = await buildStubData(env, userId, row);
  if (!data.ok) return json(data, 404);

  const to = String(body.to || data.employee.email || '').trim();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(to)) {
    return json({ ok: false, error: `No valid email for ${data.employee.name}. Add one on the Employees tab, or pass "to".` }, 400);
  }

  const html = renderPayStub(data);
  const safeName = data.employee.name.replace(/[^\w\- ]+/g, '').trim();
  const filename = `${safeName} — ${data.run.payDate}.html`;

  // ── Drive copy (best-effort; the email still goes if Drive fails) ──
  let driveUrl = data.run.stub || '';
  let driveError = '';
  if (!driveUrl) {
    const tok = await getGoogleAccessToken(env, userId);
    if (tok.ok) {
      const up = await uploadToDrive(tok.accessToken, {
        folderPath: ['AI Bookkeeper Payroll', data.run.payDate.slice(0, 4)],
        filename, mimeType: 'text/html', bytes: new TextEncoder().encode(html),
      });
      if (up.ok) {
        driveUrl = up.url;
        const w = await writeRange(env, userId, `'${PAYROLL_TAB}'!U${row}`, [[driveUrl]]);
        if (!w.ok) driveError = 'Saved to Drive but could not write the link to the Payroll row: ' + w.error;
      } else {
        driveError = up.error;
      }
    } else {
      driveError = tok.error || 'Google token unavailable';
    }
  }

  // ── Email via Resend ──
  if (!env.RESEND_API_KEY) {
    return json({ ok: false, error: 'Email is not configured yet (RESEND_API_KEY missing). The stub was saved to Drive.', driveUrl, driveError }, 501);
  }
  const from = env.PAYROLL_FROM || 'AI Bookkeeper Payroll <payroll@aibookkeeper.ca>';
  const employerName = data.employer.tradingName || data.employer.businessName || 'your employer';
  const subject = `Pay statement — ${data.run.payDate} — ${employerName}`;
  const cc = data.employer.email && data.employer.email.toLowerCase() !== to.toLowerCase() ? [data.employer.email] : [];
  const b64 = btoa(unescape(encodeURIComponent(html)));

  const r = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: { Authorization: `Bearer ${env.RESEND_API_KEY}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      from, to: [to], cc, subject,
      html,
      text: `Pay statement for ${data.employee.name}, pay date ${data.run.payDate}. Gross ${data.run.gross.toFixed(2)}, deductions ${(data.run.cpp + data.run.ei + data.run.fedTax + data.run.onTax).toFixed(2)}, net ${data.run.netPay.toFixed(2)}. Open the attached statement for the full breakdown.`,
      attachments: [{ filename, content: b64 }],
    }),
  });
  const rd = await r.json().catch(() => ({}));
  if (!r.ok) {
    return json({ ok: false, error: 'Email failed: ' + (rd?.message || rd?.error || r.status), driveUrl, driveError }, 502);
  }

  return json({ ok: true, to, cc, subject, messageId: rd.id || null, driveUrl, driveError: driveError || null });
}
