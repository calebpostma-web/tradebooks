// ════════════════════════════════════════════════════════════════════
// POST /api/yearend/upload-statement
//
// Stores a single monthly bank or AMEX statement PDF in Drive at:
//   AI Bookkeeper Year-End / FY{YYYY} / Statements / {BANK}_{YYYY-MM}.pdf
//
// The fiscal year folder is computed from the statement month + the user's
// fiscalYearEnd profile field — e.g. a Mar 2026 statement for a Mar 31 FY
// lands in FY2026; an Apr 2026 statement lands in FY2027.
//
// Filename uses the bank slug + YYYY-MM so the year-end checklist can scan
// the folder and quickly tell which months are still missing.
//
// Request body:
//   {
//     bank: 'BMO' | 'AMEX' | 'Other',  // free text allowed for "Other"
//     month: 'YYYY-MM',                 // statement period
//     pdf: { base64, filename?, mimeType? },
//     fiscalYearEnd?: 'March 31',       // optional override; defaults to profile field
//   }
// ════════════════════════════════════════════════════════════════════

import { getGoogleAccessToken } from '../../_google.js';
import { authenticateRequest, json, options, fiscalYearOf } from '../../_shared.js';
import { findOrCreateFolder, uploadFile, makeShareable } from '../../_drive.js';

const PARENT_FOLDER = 'AI Bookkeeper Year-End';

export const onRequestOptions = () => options();

export async function onRequestPost({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  const userId = auth.userId;

  let body;
  try { body = await request.json(); }
  catch { return json({ ok: false, error: 'Invalid JSON' }, 400); }

  const bank = (body.bank || '').trim();
  const month = (body.month || '').trim();   // 'YYYY-MM'
  const fye = body.fiscalYearEnd || 'December 31';

  if (!bank) return json({ ok: false, error: 'bank is required' }, 400);
  if (!/^\d{4}-\d{2}$/.test(month)) return json({ ok: false, error: 'month must be YYYY-MM' }, 400);
  if (!body.pdf || !body.pdf.base64) return json({ ok: false, error: 'pdf base64 is required' }, 400);

  // Compute which fiscal year this month belongs to. We anchor to the LAST
  // day of the month to avoid edge cases (statement for March → fiscal year
  // determined by Mar 31, not Mar 1).
  const [y, m] = month.split('-').map(Number);
  const lastDayOfMonth = new Date(Date.UTC(y, m, 0));  // 0th day of next month = last day of this month
  const fy = fiscalYearOf(fye, lastDayOfMonth);

  const tok = await getGoogleAccessToken(env, userId);
  if (!tok.ok) return json({ ok: false, error: tok.error || 'Drive token unavailable' }, 401);

  try {
    // ── Folder hierarchy: AI Bookkeeper Year-End / FY2026 / Statements ──
    const parentId = await findOrCreateFolder(tok.accessToken, PARENT_FOLDER, null);
    if (!parentId) throw new Error('Could not create year-end parent folder');
    const fyFolderId = await findOrCreateFolder(tok.accessToken, fy.fyLabel, parentId);
    if (!fyFolderId) throw new Error('Could not create fiscal year folder');
    const stmtFolderId = await findOrCreateFolder(tok.accessToken, 'Statements', fyFolderId);
    if (!stmtFolderId) throw new Error('Could not create Statements folder');

    // ── Filename + decode + upload ──
    const safeBank = bank.replace(/[^A-Za-z0-9]+/g, '');
    const ext = (body.pdf.mimeType === 'image/jpeg') ? 'jpg'
              : (body.pdf.mimeType === 'image/png')  ? 'png'
              : 'pdf';
    const filename = body.pdf.filename || `${safeBank}_${month}.${ext}`;
    const mimeType = body.pdf.mimeType || 'application/pdf';

    const base64Data = body.pdf.base64.includes(',') ? body.pdf.base64.split(',').pop() : body.pdf.base64;
    const binary = Uint8Array.from(atob(base64Data), c => c.charCodeAt(0));

    const uploaded = await uploadFile(tok.accessToken, stmtFolderId, filename, mimeType, binary);
    await makeShareable(tok.accessToken, uploaded.id);

    return json({
      ok: true,
      bank, month, filename,
      fiscalYear: fy.fyLabel,
      driveUrl: uploaded.viewUrl,
      fileId: uploaded.id,
    });
  } catch (err) {
    return json({ ok: false, error: 'Upload failed: ' + err.message }, 500);
  }
}
