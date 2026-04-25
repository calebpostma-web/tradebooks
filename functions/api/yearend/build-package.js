// ════════════════════════════════════════════════════════════════════
// POST /api/yearend/build-package
//
// Assembles a year-end package folder in Drive that MNP can open with a single
// shareable link. Includes:
//
//   Postma_YearEnd_FY2026/
//     Cover Letter (Google Doc)            — auto-generated summary + index
//     Books snapshot FY2026.xlsx           — full spreadsheet export
//     Statements/                          — shortcut to the FY statement folder
//     CRA Receipts/                        — copies of CRA receipt PDFs for this FY
//     Expense Receipts/                    — shortcuts to relevant calendar-year receipt folders
//
// Why this mix (copies vs. shortcuts):
//   - Statements live in a dedicated FY folder already, so a folder shortcut is clean.
//   - CRA receipts span two calendar-year folders; we copy the FY-window subset
//     so MNP gets exactly what they should see, not extras.
//   - Expense receipts can be 100+ per year; copying takes Drive space and time,
//     so we shortcut the calendar-year folders and let MNP filter by date.
//
// All originals get marked "anyone with link" so the shortcuts resolve when MNP
// opens the package.
//
// Request body:
//   { fy: 'FY2026', fye: 'March 31', businessName?: 'Postma Contracting Inc.' }
// ════════════════════════════════════════════════════════════════════

import { getGoogleAccessToken, getUserSheetId } from '../../_google.js';
import { readRange } from '../../_sheets.js';
import { authenticateRequest, json, options, fiscalYearOf, parseFiscalYearEnd } from '../../_shared.js';
import {
  findOrCreateFolder, createFolder, makeShareable, uploadFile,
  createShortcut, listFolderFiles, exportSpreadsheetAsXlsx,
  createCoverLetterDoc,
} from '../../_drive.js';

const REM_TAB = '📑 CRA Remittances';
const TXN_TAB = '📒 Transactions';
const INV_TAB = '🧾 Invoices';
const PAY_TAB = '💼 Payroll';

const PACKAGES_PARENT = 'AI Bookkeeper Year-End Packages';
const STATEMENTS_PARENT = 'AI Bookkeeper Year-End';
const CRA_PARENT = 'AI Bookkeeper CRA Remittances';
const RECEIPTS_PARENT = 'AI Bookkeeper Receipts';

const DRIVE_API = 'https://www.googleapis.com/drive/v3/files';

export const onRequestOptions = () => options();

export async function onRequestPost({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  const userId = auth.userId;

  let body;
  try { body = await request.json(); }
  catch { return json({ ok: false, error: 'Invalid JSON' }, 400); }

  const fye = body.fye || 'December 31';
  const businessName = (body.businessName || 'Business').replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '') || 'Business';

  const fy = resolveFiscalYear(body.fy, fye);

  const tok = await getGoogleAccessToken(env, userId);
  if (!tok.ok) return json({ ok: false, error: tok.error || 'Drive token unavailable' }, 401);

  const spreadsheetId = await getUserSheetId(env, userId);
  if (!spreadsheetId) return json({ ok: false, error: 'No sheet connected' }, 400);

  const log = [];   // step-by-step record we return so the UI can show what happened
  const warnings = [];

  try {
    // ── 1. Create the package folder ──────────────────────────────────────
    const packagesParentId = await findOrCreateFolder(tok.accessToken, PACKAGES_PARENT, null);
    if (!packagesParentId) throw new Error('Could not create packages parent folder');

    const stamp = new Date().toISOString().slice(0, 10);
    const packageName = `${businessName}_YearEnd_${fy.fyLabel}_${stamp}`;
    const packageFolderId = await createFolder(tok.accessToken, packageName, packagesParentId);
    log.push(`Created package folder: ${packageName}`);

    // Make the package folder shareable up-front so subsequent operations can be tested.
    await makeShareable(tok.accessToken, packageFolderId);

    // ── 2. Export the spreadsheet as XLSX into the package folder ─────────
    try {
      const xlsxBytes = await exportSpreadsheetAsXlsx(tok.accessToken, spreadsheetId);
      await uploadFile(
        tok.accessToken,
        packageFolderId,
        `Books snapshot ${fy.fyLabel}.xlsx`,
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        xlsxBytes,
      );
      log.push(`Exported books snapshot (XLSX) — ${(xlsxBytes.length / 1024).toFixed(1)} KB`);
    } catch (err) {
      warnings.push(`Books snapshot export failed: ${err.message}`);
    }

    // ── 3. Pull CRA Remittances rows for this FY ──────────────────────────
    const remittancesForFy = [];
    try {
      const remRead = await readRange(env, userId, `'${REM_TAB}'!B12:J500`);
      if (remRead.ok) {
        for (const row of (remRead.values || [])) {
          if (!row[0]) continue;
          const t = Date.parse(row[0]);
          if (!Number.isFinite(t)) continue;
          if (t >= fy.startDate.getTime() && t <= fy.endDate.getTime()) {
            remittancesForFy.push({
              datePaid: row[0], type: row[1], period: row[2], amount: parseFloat(row[3]) || 0,
              confirmation: row[4], account: row[5], pdfUrl: row[6], notes: row[7], ref: row[8],
            });
          }
        }
        log.push(`Found ${remittancesForFy.length} CRA remittances in ${fy.fyLabel}`);
      } else {
        warnings.push(`Could not read CRA Remittances tab: ${remRead.error}`);
      }
    } catch (err) {
      warnings.push(`CRA Remittances read failed: ${err.message}`);
    }

    // ── 4. Build the Statements/ shortcut ─────────────────────────────────
    try {
      const statementsParent = await findOrCreateFolder(tok.accessToken, STATEMENTS_PARENT, null);
      const fyFolderId = statementsParent ? await findOrCreateFolder(tok.accessToken, fy.fyLabel, statementsParent) : null;
      const stmtFolderId = fyFolderId ? await findOrCreateFolder(tok.accessToken, 'Statements', fyFolderId) : null;
      if (stmtFolderId) {
        await makeShareable(tok.accessToken, stmtFolderId);
        await createShortcut(tok.accessToken, stmtFolderId, packageFolderId, `Statements (${fy.fyLabel})`);
        const stmtFiles = await listFolderFiles(tok.accessToken, stmtFolderId);
        log.push(`Linked Statements folder (${stmtFiles.length} files)`);
      } else {
        warnings.push('Statements folder not found — no monthly statements have been uploaded yet.');
      }
    } catch (err) {
      warnings.push(`Statements link failed: ${err.message}`);
    }

    // ── 5. CRA Receipts/ — copy individual PDFs that fall in this FY ──────
    let craReceiptsCopied = 0;
    let craFolderId = null;
    try {
      craFolderId = await createFolder(tok.accessToken, 'CRA Receipts', packageFolderId);
      for (const r of remittancesForFy) {
        if (!r.pdfUrl) continue;
        const fileId = extractDriveFileId(r.pdfUrl);
        if (!fileId) continue;
        try {
          const safeType = String(r.type || 'Other').replace(/[^A-Za-z0-9]+/g, '');
          const safePeriod = String(r.period || '').replace(/[^A-Za-z0-9]+/g, '-').replace(/^-+|-+$/g, '') || 'Period';
          const copyName = `${safeType}_${safePeriod}_${r.datePaid}.pdf`;
          await copyDriveFile(tok.accessToken, fileId, craFolderId, copyName);
          craReceiptsCopied++;
        } catch (copyErr) {
          warnings.push(`Could not copy CRA receipt for ${r.type} ${r.period}: ${copyErr.message}`);
        }
      }
      log.push(`Copied ${craReceiptsCopied} of ${remittancesForFy.length} CRA receipt PDFs into package`);
    } catch (err) {
      warnings.push(`CRA Receipts folder failed: ${err.message}`);
    }

    // ── 6. Expense Receipts/ — shortcut(s) to calendar-year folders ───────
    try {
      const expenseFolderId = await createFolder(tok.accessToken, 'Expense Receipts', packageFolderId);
      const receiptsParent = await findOrCreateFolder(tok.accessToken, RECEIPTS_PARENT, null);
      if (receiptsParent) {
        const startYear = fy.startDate.getUTCFullYear();
        const endYear = fy.endDate.getUTCFullYear();
        let yearsLinked = 0;
        for (let y = startYear; y <= endYear; y++) {
          const yearFolderId = await findOrCreateFolder(tok.accessToken, String(y), receiptsParent);
          if (yearFolderId) {
            await makeShareable(tok.accessToken, yearFolderId);
            await createShortcut(tok.accessToken, yearFolderId, expenseFolderId, `Receipts ${y}`);
            yearsLinked++;
          }
        }
        log.push(`Linked ${yearsLinked} calendar-year receipt folder(s) — MNP filters by date inside`);
      } else {
        warnings.push('Receipts parent folder missing — no expense receipts have been scanned yet.');
      }
    } catch (err) {
      warnings.push(`Expense Receipts link failed: ${err.message}`);
    }

    // ── 7. Cover letter Doc ───────────────────────────────────────────────
    try {
      const coverBody = renderCoverLetter({
        businessName: body.businessName || 'My Business',
        fy,
        remittances: remittancesForFy,
        craReceiptsCopied,
        warnings,
      });
      const docTitle = `Year-End Package — ${body.businessName || 'My Business'} — ${fy.fyLabel}`;
      const doc = await createCoverLetterDoc(tok.accessToken, packageFolderId, docTitle, coverBody);
      await makeShareable(tok.accessToken, doc.id);
      log.push(`Created cover letter Google Doc`);
    } catch (err) {
      warnings.push(`Cover letter creation failed: ${err.message}`);
    }

    // ── 8. Done — return shareable folder URL ─────────────────────────────
    const folderUrl = `https://drive.google.com/drive/folders/${packageFolderId}`;
    return json({
      ok: true,
      fiscalYear: fy.fyLabel,
      packageFolderId,
      packageUrl: folderUrl,
      packageName,
      log,
      warnings,
    });
  } catch (err) {
    return json({ ok: false, error: 'Package build failed: ' + err.message, log, warnings }, 500);
  }
}

// ── Cover-letter generator ─────────────────────────────────────────

function renderCoverLetter({ businessName, fy, remittances, craReceiptsCopied, warnings }) {
  const totalCra = remittances.reduce((s, r) => s + (r.amount || 0), 0);
  const byType = remittances.reduce((acc, r) => {
    const k = r.type || 'Other';
    acc[k] = (acc[k] || 0) + (r.amount || 0);
    return acc;
  }, {});
  const today = new Date().toISOString().slice(0, 10);

  let lines = [];
  lines.push(`YEAR-END PACKAGE`);
  lines.push(`${businessName}`);
  lines.push(`Fiscal Year: ${fy.fyLabel}  (${fy.startISO} → ${fy.endISO})`);
  lines.push(`Prepared: ${today}`);
  lines.push('');
  lines.push('─────────────────────────────────────────────────');
  lines.push('WHAT IS IN THIS FOLDER');
  lines.push('─────────────────────────────────────────────────');
  lines.push('');
  lines.push(`1. Books snapshot ${fy.fyLabel}.xlsx`);
  lines.push(`   The complete bookkeeping spreadsheet for the fiscal year. All tabs included:`);
  lines.push(`   Dashboard, Categories, Transactions, Invoices, HST Returns, Year-End,`);
  lines.push(`   Payroll, Work Log, CRA Remittances.`);
  lines.push('');
  lines.push(`2. Statements (${fy.fyLabel})`);
  lines.push(`   Monthly bank + AMEX statement PDFs uploaded for this fiscal year.`);
  lines.push(`   Filenames follow the pattern BANK_YYYY-MM.pdf so missing months are obvious.`);
  lines.push('');
  lines.push(`3. CRA Receipts/`);
  lines.push(`   ${craReceiptsCopied} CRA payment confirmation PDFs (HST, payroll source`);
  lines.push(`   deductions, corporate tax instalments) for this fiscal year. These were`);
  lines.push(`   copied from the master CRA Remittances Drive folder.`);
  lines.push('');
  lines.push(`4. Expense Receipts/`);
  lines.push(`   Shortcut(s) to the master receipt folders covering this fiscal year.`);
  lines.push(`   Filter by date inside to find receipts for ${fy.startISO} → ${fy.endISO}.`);
  lines.push('');
  lines.push('─────────────────────────────────────────────────');
  lines.push('CRA REMITTANCES — TOTAL PAID THIS YEAR');
  lines.push('─────────────────────────────────────────────────');
  lines.push('');
  for (const [type, amt] of Object.entries(byType)) {
    lines.push(`  ${type.padEnd(36)} $${amt.toFixed(2)}`);
  }
  lines.push('  ' + '─'.repeat(45));
  lines.push(`  ${'TOTAL'.padEnd(36)} $${totalCra.toFixed(2)}`);
  lines.push('');
  lines.push('See the CRA Remittances tab in the books snapshot for date / period /');
  lines.push('confirmation # detail. PDF copies are in the CRA Receipts folder.');
  lines.push('');
  lines.push('─────────────────────────────────────────────────');
  lines.push('NOTES FOR YOUR ACCOUNTANT');
  lines.push('─────────────────────────────────────────────────');
  lines.push('');
  lines.push('• Books are kept on a CASH BASIS. Revenue is recognized when deposits hit');
  lines.push('  the bank, not when invoices are issued. The HST Returns tab pulls from');
  lines.push('  the Transactions tab on this same basis.');
  lines.push('');
  lines.push('• Internal Transfer is the category used for movements between own');
  lines.push('  accounts (BMO ↔ AMEX bill payments) AND payments to CRA. These are');
  lines.push('  excluded from P&L and HST math by formula.');
  lines.push('');
  lines.push('• Expense receipts are stored in Google Drive. The Expense Receipts');
  lines.push('  shortcut(s) point to the master folders — please request access if');
  lines.push('  you cannot open them with the share link.');
  lines.push('');
  if (warnings.length > 0) {
    lines.push('─────────────────────────────────────────────────');
    lines.push('BUILD WARNINGS');
    lines.push('─────────────────────────────────────────────────');
    lines.push('');
    for (const w of warnings) {
      lines.push(`  • ${w}`);
    }
    lines.push('');
  }
  lines.push('─────────────────────────────────────────────────');
  lines.push(`Generated by tradebooks · ${today}`);
  lines.push('');
  return lines.join('\n');
}

// ── Helpers ────────────────────────────────────────────────────────

function extractDriveFileId(url) {
  if (!url) return null;
  const m = String(url).match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : null;
}

async function copyDriveFile(accessToken, fileId, parentFolderId, newName) {
  const resp = await fetch(`${DRIVE_API}/${fileId}/copy?fields=id,name`, {
    method: 'POST',
    headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ name: newName, parents: [parentFolderId] }),
  });
  const data = await resp.json();
  if (data.error) throw new Error(data.error.message || 'Copy failed');
  return data.id;
}

function resolveFiscalYear(fyParam, fye) {
  if (fyParam && /^FY\d{4}$/.test(fyParam)) {
    const fyEndYear = parseInt(fyParam.slice(2), 10);
    const { month, day } = parseFiscalYearEnd(fye);
    return fiscalYearOf(fye, new Date(Date.UTC(fyEndYear, month - 1, day)));
  }
  return fiscalYearOf(fye);
}
