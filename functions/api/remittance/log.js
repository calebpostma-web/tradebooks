// ════════════════════════════════════════════════════════════════════
// POST /api/remittance/log
//
// Logs a single payment to CRA (HST, payroll source deductions, corp tax
// instalment, etc). Three side effects on success:
//
//   1. Append a row to the 📑 CRA Remittances tab (full structured record).
//   2. Append a matching negative row to 📒 Transactions, category
//      'Internal Transfer' so it leaves the bank but doesn't pollute P&L
//      or HST math. Source column flags the payment type.
//   3. (Optional) Upload the receipt PDF to Drive at
//      'AI Bookkeeper CRA Remittances/YYYY/' and store the share link.
//
// PDF upload errors are non-fatal — the row still gets logged so the user
// isn't blocked. They'll see the missing-receipt count in the tab summary.
//
// Request body:
//   {
//     type: 'HST' | 'Payroll (PD7A)' | 'Corporate Tax Instalment' |
//           'Corporate Tax Final' | 'Other',
//     period: string,          // e.g. 'Q4 2025-2026', 'Mar 2026'
//     amount: number,          // CAD, positive
//     datePaid: 'YYYY-MM-DD',
//     confirmation: string,    // CRA confirmation # or '0' if none
//     account: string,         // 'BMO' / 'Cash' / etc, defaults 'BMO'
//     notes?: string,
//     pdf?: { base64: string, filename?: string, mimeType?: string },
//   }
// ════════════════════════════════════════════════════════════════════

import { appendRows } from '../../_sheets.js';
import { getGoogleAccessToken } from '../../_google.js';
import { authenticateRequest, json, options } from '../../_shared.js';

const REMIT_TAB = '📑 CRA Remittances';
const TXN_TAB = '📒 Transactions';
const TRANSFER_CATEGORY = 'Internal Transfer';
const DRIVE_API = 'https://www.googleapis.com/drive/v3/files';
const UPLOAD_API = 'https://www.googleapis.com/upload/drive/v3/files';
const PARENT_FOLDER = 'AI Bookkeeper CRA Remittances';

const VALID_TYPES = new Set([
  'HST',
  'Payroll (PD7A)',
  'Corporate Tax Instalment',
  'Corporate Tax Final',
  'Other',
]);

// Map the user-facing Type to a short Source tag for the Transactions row.
// Keeps the Transaction grep-able by remittance kind without overloading
// the Category column (which stays 'Internal Transfer' for math reasons).
const SOURCE_BY_TYPE = {
  'HST':                       'CRA HST Remit',
  'Payroll (PD7A)':            'CRA Payroll Remit',
  'Corporate Tax Instalment':  'CRA Corp Tax Inst',
  'Corporate Tax Final':       'CRA Corp Tax Final',
  'Other':                     'CRA Other',
};

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

  // ── Validate ──
  const type = (body.type || '').trim();
  const period = (body.period || '').trim();
  const amount = Math.round((parseFloat(body.amount) || 0) * 100) / 100;
  const datePaid = (body.datePaid || '').trim();
  const confirmation = (body.confirmation || '').toString().trim();
  const account = (body.account || 'BMO').trim();
  const notes = (body.notes || '').trim();

  if (!VALID_TYPES.has(type)) return json({ ok: false, error: `Invalid type: ${type}` }, 400);
  if (!period) return json({ ok: false, error: 'period is required' }, 400);
  if (amount <= 0) return json({ ok: false, error: 'amount must be > 0' }, 400);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(datePaid)) return json({ ok: false, error: 'datePaid must be YYYY-MM-DD' }, 400);

  // Build the Transactions ref now so we can write it into both rows.
  // Pattern: CRA-REMIT-{TYPESLUG}-{YYYYMMDD} — keeps it deduplicable and
  // human-readable in bank-import dedup checks.
  const typeSlug = SOURCE_BY_TYPE[type].replace(/[^A-Z0-9]/gi, '').toUpperCase();
  const ref = `CRA-REMIT-${typeSlug}-${datePaid.replace(/-/g, '')}`;

  // ── Optional PDF upload ──
  // We upload first (if requested) so we can write the Drive URL into both
  // rows in one pass. Failure is non-fatal: log a warning and proceed.
  let pdfDriveUrl = '';
  let pdfWarning = '';
  if (body.pdf && body.pdf.base64) {
    try {
      const tok = await getGoogleAccessToken(env, userId);
      if (!tok.ok) throw new Error(tok.error || 'Drive token unavailable');

      const year = datePaid.slice(0, 4);
      const parentId = await findOrCreateFolder(tok.accessToken, PARENT_FOLDER, null);
      if (!parentId) throw new Error('Could not create remittance folder');
      const yearFolderId = await findOrCreateFolder(tok.accessToken, year, parentId);
      if (!yearFolderId) throw new Error('Could not create year folder');

      // Filename: HST_Q4-2026_2026-04-28.pdf
      const safeType = type.replace(/[^A-Za-z0-9]+/g, '');
      const safePeriod = period.replace(/[^A-Za-z0-9]+/g, '-').replace(/^-+|-+$/g, '') || 'Period';
      const ext = (body.pdf.mimeType === 'image/jpeg') ? 'jpg'
                 : (body.pdf.mimeType === 'image/png')  ? 'png'
                 : 'pdf';
      const filename = body.pdf.filename || `${safeType}_${safePeriod}_${datePaid}.${ext}`;
      const mimeType = body.pdf.mimeType || 'application/pdf';

      const base64Data = body.pdf.base64.includes(',') ? body.pdf.base64.split(',').pop() : body.pdf.base64;
      const binary = Uint8Array.from(atob(base64Data), c => c.charCodeAt(0));

      pdfDriveUrl = await uploadToDrive(tok.accessToken, yearFolderId, filename, mimeType, binary);
      if (!pdfDriveUrl) throw new Error('Upload returned no URL');
    } catch (err) {
      pdfWarning = `PDF upload failed: ${err.message}. Row was still logged — re-upload the receipt directly to the Drive folder if needed.`;
    }
  }

  // ── Write CRA Remittances row (cols B-J: 9 cols) ──
  const remitRow = [[
    datePaid,                      // B Date Paid
    type,                          // C Type
    period,                        // D Period Covered
    amount,                        // E Amount
    confirmation || '',            // F Confirmation #
    account,                       // G Account
    pdfDriveUrl,                   // H PDF Receipt (Drive)
    notes,                         // I Notes
    ref,                           // J Linked Txn Ref
  ]];
  const remResult = await appendRows(env, userId, `'${REMIT_TAB}'!B12:J`, remitRow);
  if (!remResult.ok) return json({ ok: false, error: 'Remittance log write failed: ' + remResult.error });

  // ── Write Transactions row (cols B-M: 12 cols) ──
  // Negative amount (cash leaving), Internal Transfer category so HST and P&L
  // formulas exclude it. Source carries the type so it's filterable.
  const description = `${type} remittance — ${period}` + (confirmation ? ` (conf #${confirmation})` : '');
  const txnRow = [[
    datePaid,                      // B Date
    'CRA',                         // C Party
    description,                   // D Description
    -amount,                       // E Amount (signed; negative = cash out)
    TRANSFER_CATEGORY,             // F Category — excluded from P&L + HST math
    'No',                          // G HST?
    0,                             // H HST Amount
    account,                       // I Account
    SOURCE_BY_TYPE[type],          // J Source (e.g. 'CRA HST Remit')
    ref,                           // K Ref — matches the Remittances row
    '',                            // L Related Invoice
    'N/A',                         // M Match Status
  ]];
  const txnResult = await appendRows(env, userId, `'${TXN_TAB}'!B12:M`, txnRow);
  // Non-fatal — if Transactions write fails the Remittances row still exists,
  // and the user can manually add the bank entry.
  let txnWarning = '';
  if (!txnResult.ok) {
    txnWarning = `Transactions row write failed: ${txnResult.error}. Add it manually so the bank reconciles.`;
  }

  return json({
    ok: true,
    type, period, amount, datePaid, ref,
    pdfDriveUrl,
    remittanceRange: remResult.updates?.updatedRange,
    transactionRange: txnResult.ok ? txnResult.updates?.updatedRange : null,
    warnings: [pdfWarning, txnWarning].filter(Boolean),
  });
}

// ── Drive helpers (mirrored from receipt/upload.js) ──

async function uploadToDrive(accessToken, folderId, filename, mimeType, binary) {
  const boundary = '-------remit-boundary-' + Math.random().toString(36).slice(2);
  const metadata = { name: filename, parents: [folderId], mimeType };
  const preamble = `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n--${boundary}\r\nContent-Type: ${mimeType}\r\n\r\n`;
  const epilogue = `\r\n--${boundary}--`;

  const enc = new TextEncoder();
  const part1 = enc.encode(preamble);
  const part3 = enc.encode(epilogue);
  const fullBody = new Uint8Array(part1.length + binary.length + part3.length);
  fullBody.set(part1, 0);
  fullBody.set(binary, part1.length);
  fullBody.set(part3, part1.length + binary.length);

  const uploadResp = await fetch(`${UPLOAD_API}?uploadType=multipart`, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': `multipart/related; boundary=${boundary}`,
    },
    body: fullBody,
  });
  const uploadData = await uploadResp.json();
  if (uploadData.error) throw new Error(uploadData.error.message || 'Drive upload failed');

  // Make the file viewable by anyone with the link (so MNP can open it).
  await fetch(`${DRIVE_API}/${uploadData.id}/permissions`, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ role: 'reader', type: 'anyone' }),
  });

  return `https://drive.google.com/file/d/${uploadData.id}/view`;
}

async function findOrCreateFolder(accessToken, name, parentId) {
  const escapedName = name.replace(/'/g, "\\'");
  let query = `name='${escapedName}' and mimeType='application/vnd.google-apps.folder' and trashed=false`;
  if (parentId) query += ` and '${parentId}' in parents`;

  const searchResp = await fetch(
    `${DRIVE_API}?q=${encodeURIComponent(query)}&fields=files(id,name)`,
    { headers: { 'Authorization': `Bearer ${accessToken}` } }
  );
  const searchData = await searchResp.json();
  if (searchData.files && searchData.files.length) return searchData.files[0].id;

  const metadata = { name, mimeType: 'application/vnd.google-apps.folder' };
  if (parentId) metadata.parents = [parentId];

  const createResp = await fetch(DRIVE_API, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(metadata),
  });
  const createData = await createResp.json();
  return createData.id || null;
}
