// ════════════════════════════════════════════════════════════════════
// GET /api/yearend/list-statements?fy=FY2026&fye=March+31
//
// Returns the list of statement PDFs uploaded for a given fiscal year, plus
// the per-month roll-up so the year-end checklist can show ✓/✗ per month
// per bank.
//
// Query params:
//   fy   — fiscal year label, e.g. 'FY2026' (optional; defaults to current)
//   fye  — fiscal year end string, e.g. 'March 31' (required for current-FY default)
//
// Response:
//   {
//     ok, fiscalYear, months: [{month: '2025-04', bmo: {…} | null, amex: {…} | null, other: [...]}],
//     allFiles: [{name, month, bank, driveUrl, fileId, createdTime}],
//   }
// ════════════════════════════════════════════════════════════════════

import { getGoogleAccessToken } from '../../_google.js';
import { authenticateRequest, json, options, fiscalYearOf, fiscalYearMonths } from '../../_shared.js';
import { findOrCreateFolder, listFolderFiles } from '../../_drive.js';

const PARENT_FOLDER = 'AI Bookkeeper Year-End';

export const onRequestOptions = () => options();

export async function onRequest({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  const userId = auth.userId;

  const url = new URL(request.url);
  const fyParam = url.searchParams.get('fy');
  const fye = url.searchParams.get('fye') || 'December 31';

  // Resolve the fiscal year window. If a label like 'FY2026' was passed,
  // synthesize the matching window without depending on the current date.
  let fy;
  if (fyParam && /^FY\d{4}$/.test(fyParam)) {
    const fyEndYear = parseInt(fyParam.slice(2), 10);
    // Pick a date inside that fiscal year (the FY end date is safe).
    const { month, day } = parseFromFye(fye);
    fy = fiscalYearOf(fye, new Date(Date.UTC(fyEndYear, month - 1, day)));
  } else {
    fy = fiscalYearOf(fye);
  }

  const tok = await getGoogleAccessToken(env, userId);
  if (!tok.ok) return json({ ok: false, error: tok.error || 'Drive token unavailable' }, 401);

  // Resolve folders. If they don't exist, we just have an empty list — that's
  // a valid state (user hasn't uploaded anything yet).
  let files = [];
  try {
    const parentId = await findOrCreateFolder(tok.accessToken, PARENT_FOLDER, null);
    if (parentId) {
      const fyFolderId = await findOrCreateFolder(tok.accessToken, fy.fyLabel, parentId);
      if (fyFolderId) {
        const stmtFolderId = await findOrCreateFolder(tok.accessToken, 'Statements', fyFolderId);
        if (stmtFolderId) {
          files = await listFolderFiles(tok.accessToken, stmtFolderId);
        }
      }
    }
  } catch (err) {
    return json({ ok: false, error: 'Drive list failed: ' + err.message }, 500);
  }

  // Parse each filename for bank + month. Pattern: BANK_YYYY-MM.{pdf|png|jpg}
  const allFiles = files.map(f => {
    const m = f.name.match(/^(.+?)_(\d{4}-\d{2})\./);
    return {
      name: f.name,
      bank: m ? m[1] : 'Unknown',
      month: m ? m[2] : '',
      driveUrl: f.webViewLink || `https://drive.google.com/file/d/${f.id}/view`,
      fileId: f.id,
      createdTime: f.createdTime,
      size: f.size,
    };
  });

  // Build per-month rollup. We use the canonical fy month list so the
  // checklist always shows all 12 slots even when nothing's uploaded.
  const months = fiscalYearMonths(fy.startDate, fy.endDate).map(({ year, month }) => {
    const ym = `${year}-${String(month).padStart(2, '0')}`;
    const matches = allFiles.filter(f => f.month === ym);
    return {
      month: ym,
      bmo: matches.find(f => /^BMO/i.test(f.bank)) || null,
      amex: matches.find(f => /^AMEX/i.test(f.bank)) || null,
      other: matches.filter(f => !/^BMO|^AMEX/i.test(f.bank)),
    };
  });

  return json({
    ok: true,
    fiscalYear: fy.fyLabel,
    startDate: fy.startISO,
    endDate: fy.endISO,
    months,
    allFiles,
  });
}

// Local helper — needed only here because the shared helper is in _shared.js
// and we don't want to import the entire FY parsing chain again.
function parseFromFye(s) {
  const map = {january:1,february:2,march:3,april:4,may:5,june:6,july:7,august:8,september:9,october:10,november:11,december:12};
  const m = String(s).match(/^([a-zA-Z]+)\s+(\d{1,2})$/);
  if (!m) return { month: 12, day: 31 };
  return { month: map[m[1].toLowerCase()] || 12, day: parseInt(m[2], 10) || 31 };
}
