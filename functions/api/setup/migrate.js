// ════════════════════════════════════════════════════════════════════
// POST /api/setup/migrate           (applies migrations, returns summary)
// GET  /api/setup/migrate?dryRun=1  (lists pending migrations, no writes)
//
// Idempotent schema migration for users whose Google Sheet was set up before
// new tabs / columns existed. Reads current spreadsheet metadata, applies the
// minimum set of changes to bring it to the latest schema, returns a summary
// of what was added (or "Already up to date").
//
// dryRun is the same flow with all writes skipped — used by the front-end
// to detect pending migrations on app load and show an "update available"
// banner without touching the user's sheet.
//
// Each migration is a small, self-contained block. Add new ones by appending
// another `await applyXxx(...)` call below; the existing logic is untouched.
//
// SAFETY: every migration must be safe to re-run. Check first, then write.
// ════════════════════════════════════════════════════════════════════

import { getGoogleAccessToken, getUserSheetId } from '../../_google.js';
import { getSpreadsheetMetadata, spreadsheetsBatchUpdate, writeRange } from '../../_sheets.js';
import { authenticateRequest, json, options } from '../../_shared.js';

const SHEETS_API = 'https://sheets.googleapis.com/v4/spreadsheets';

// ── Color palette + format constants (mirrored from google-setup.js so
//    migrations stay self-contained when google-setup.js evolves) ──
const COLORS = {
  teal:     { red: 0.078, green: 0.278, blue: 0.247 },
  tealTint: { red: 0.949, green: 0.980, blue: 0.976 },
  blue:     { red: 0.043, green: 0.231, blue: 0.400 },
  blueTint: { red: 0.929, green: 0.949, blue: 0.973 },
  white:    { red: 1, green: 1, blue: 1 },
  textMuted:{ red: 0.45, green: 0.45, blue: 0.45 },
};
const FMT_DATE = { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } };
const FMT_CURRENCY = { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } };

export const onRequestOptions = () => options();

// POST → apply migrations
export async function onRequestPost({ request, env }) {
  return runMigrations(request, env, /* dryRun */ false);
}

// GET → check what's pending (dryRun by default — never writes via GET)
export async function onRequestGet({ request, env }) {
  return runMigrations(request, env, /* dryRun */ true);
}

async function runMigrations(request, env, dryRunDefault) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  const userId = auth.userId;

  const url = new URL(request.url);
  const dryRunParam = url.searchParams.get('dryRun');
  const dryRun = dryRunParam ? dryRunParam !== 'false' && dryRunParam !== '0' : dryRunDefault;

  const meta = await getSpreadsheetMetadata(env, userId);
  if (!meta.ok) {
    // Pass through needsReauth + noSheet so the front-end can route the user to
    // the right recovery flow (reconnect Google vs. create a sheet) instead of
    // showing a dead-end "Could not read sheet" error.
    return json({
      ok: false,
      error: meta.error,
      needsReauth: !!meta.needsReauth,
      noSheet: !!meta.noSheet,
    });
  }

  const sheetsByTitle = Object.fromEntries(meta.sheets.map(s => [s.title, s]));
  const changes = [];
  const errors = [];

  // ── Migration 1: ensure 📑 CRA Remittances tab exists ──
  await applyCraRemittancesTab(env, userId, sheetsByTitle, changes, errors, dryRun);

  // ── Migration 2: ensure Invoices tab has deposit columns O/P/Q ──
  await applyInvoiceDepositColumns(env, userId, sheetsByTitle, changes, errors, dryRun);

  // ── Migration 3: ensure Transactions tab has Total (incl HST) column N ──
  await applyTransactionsTotalColumn(env, userId, sheetsByTitle, changes, errors, dryRun);

  // ── Migration 4: HST Returns C3 — smart FY start (data-driven default) ──
  await applyHstReturnsSmartFyStart(env, userId, sheetsByTitle, changes, errors, dryRun);

  // ── Migration 5: 🏦 Account Balances tab (bank reconciliation) ──
  await applyAccountBalancesTab(env, userId, sheetsByTitle, changes, errors, dryRun);

  // ── Migration 6: Year-End per-category breakdown (QUERY pivot) ──
  await applyYearEndPerCategoryBreakdown(env, userId, sheetsByTitle, changes, errors, dryRun);

  // ── Migration 7: 📓 Adjusting Entries tab (year-end accruals) ──
  await applyAdjustingEntriesTab(env, userId, sheetsByTitle, changes, errors, dryRun);

  // ── Migration 8: 🛠 Fixed Assets tab (CCA tracker) ──
  await applyFixedAssetsTab(env, userId, sheetsByTitle, changes, errors, dryRun);

  // ── Migration 9: 📊 T2 Worksheet (consolidated T2-prep view) ──
  await applyT2Worksheet(env, userId, sheetsByTitle, changes, errors, dryRun);

  return json({
    ok: true,
    dryRun,
    changes,
    errors,
    upToDate: changes.length === 0 && errors.length === 0,
  });
}

// ════════════════════════════════════════════════════════════════════
// Migration 1: 📑 CRA Remittances tab
// ════════════════════════════════════════════════════════════════════
async function applyCraRemittancesTab(env, userId, sheetsByTitle, changes, errors, dryRun) {
  const TITLE = '📑 CRA Remittances';
  if (sheetsByTitle[TITLE]) {
    return;  // already exists — nothing to do
  }

  // In dry-run mode, just record that this migration is pending and skip writes.
  if (dryRun) {
    changes.push(`Add '${TITLE}' tab — needed for the CRA Payments Log to track HST, payroll, and corp tax remittances with PDF receipts.`);
    return;
  }

  // Pick a sheetId that doesn't collide with anything existing. 1000 is the
  // canonical ID per google-setup.js, but if (for any reason) something else
  // grabbed it, fall back to the next free 100-block.
  const usedIds = new Set(Object.values(sheetsByTitle).map(s => s.sheetId));
  let sheetId = 1000;
  while (usedIds.has(sheetId)) sheetId += 100;

  // ── Step 1: create the tab ──
  const addSheetReq = {
    addSheet: {
      properties: {
        sheetId,
        title: TITLE,
        index: Object.keys(sheetsByTitle).length,
        gridProperties: { rowCount: 500, columnCount: 11, frozenRowCount: 11 },
        tabColor: COLORS.teal,
      },
    },
  };
  const addRes = await spreadsheetsBatchUpdate(env, userId, [addSheetReq]);
  if (!addRes.ok) {
    errors.push(`Could not create '${TITLE}' tab: ${addRes.error}`);
    return;
  }

  // ── Step 2: styling ──
  // Banner row, section header for stats, header row for data, banding,
  // currency/date formatting, dropdown validation. Same shape as google-setup.js.
  const styling = [
    // Banner (row 0)
    { repeatCell: {
        range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 11 },
        cell: { userEnteredFormat: {
          backgroundColor: COLORS.teal,
          textFormat: { foregroundColor: COLORS.white, bold: true, fontSize: 12 },
          horizontalAlignment: 'LEFT',
          verticalAlignment: 'MIDDLE',
          padding: { top: 8, left: 12, right: 12, bottom: 8 },
        }},
        fields: 'userEnteredFormat',
    }},
    { mergeCells: {
        range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 11 },
        mergeType: 'MERGE_ALL',
    }},
    { updateDimensionProperties: {
        range: { sheetId, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
        properties: { pixelSize: 36 },
        fields: 'pixelSize',
    }},
    // Section row (row 1) for "REMITTANCE TOTALS"
    { repeatCell: {
        range: { sheetId, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 1, endColumnIndex: 7 },
        cell: { userEnteredFormat: {
          backgroundColor: COLORS.teal,
          textFormat: { foregroundColor: COLORS.white, bold: true, fontSize: 10 },
        }},
        fields: 'userEnteredFormat',
    }},
    // Stats labels (rows 2-9, col B) — bold
    { repeatCell: {
        range: { sheetId, startRowIndex: 2, endRowIndex: 10, startColumnIndex: 1, endColumnIndex: 2 },
        cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 10 } }},
        fields: 'userEnteredFormat.textFormat',
    }},
    // Stats values (rows 2-9, cols C-F) — currency tint
    { repeatCell: {
        range: { sheetId, startRowIndex: 2, endRowIndex: 10, startColumnIndex: 2, endColumnIndex: 6 },
        cell: { userEnteredFormat: {
          backgroundColor: COLORS.tealTint,
          numberFormat: FMT_CURRENCY.numberFormat,
        }},
        fields: 'userEnteredFormat',
    }},
    // Header row (row 10)
    { repeatCell: {
        range: { sheetId, startRowIndex: 10, endRowIndex: 11, startColumnIndex: 1, endColumnIndex: 11 },
        cell: { userEnteredFormat: {
          backgroundColor: COLORS.teal,
          textFormat: { foregroundColor: COLORS.white, bold: true, fontSize: 10 },
          horizontalAlignment: 'LEFT',
          verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        }},
        fields: 'userEnteredFormat',
    }},
    { updateDimensionProperties: {
        range: { sheetId, dimension: 'ROWS', startIndex: 10, endIndex: 11 },
        properties: { pixelSize: 36 },
        fields: 'pixelSize',
    }},
    // Banding for data rows (alt row colour)
    { addBanding: {
        bandedRange: {
          range: { sheetId, startRowIndex: 11, endRowIndex: 500, startColumnIndex: 1, endColumnIndex: 11 },
          rowProperties: {
            firstBandColor: COLORS.white,
            secondBandColor: COLORS.tealTint,
          },
        },
    }},
    // Date format on col B (Date Paid)
    { repeatCell: {
        range: { sheetId, startRowIndex: 11, endRowIndex: 500, startColumnIndex: 1, endColumnIndex: 2 },
        cell: { userEnteredFormat: { numberFormat: FMT_DATE.numberFormat }},
        fields: 'userEnteredFormat.numberFormat',
    }},
    // Currency format on col E (Amount)
    { repeatCell: {
        range: { sheetId, startRowIndex: 11, endRowIndex: 500, startColumnIndex: 4, endColumnIndex: 5 },
        cell: { userEnteredFormat: { numberFormat: FMT_CURRENCY.numberFormat }},
        fields: 'userEnteredFormat.numberFormat',
    }},
    // Type dropdown on col C
    { setDataValidation: {
        range: { sheetId, startRowIndex: 11, endRowIndex: 500, startColumnIndex: 2, endColumnIndex: 3 },
        rule: { condition: { type: 'ONE_OF_LIST', values: [
          { userEnteredValue: 'HST' },
          { userEnteredValue: 'Payroll (PD7A)' },
          { userEnteredValue: 'Corporate Tax Instalment' },
          { userEnteredValue: 'Corporate Tax Final' },
          { userEnteredValue: 'Other' },
        ]}, showCustomUi: true },
    }},
    // Column widths
    colWidthReq(sheetId, 0, 1, 30),    // A gutter
    colWidthReq(sheetId, 1, 2, 100),   // B Date Paid
    colWidthReq(sheetId, 2, 3, 170),   // C Type
    colWidthReq(sheetId, 3, 4, 130),   // D Period Covered
    colWidthReq(sheetId, 4, 5, 110),   // E Amount
    colWidthReq(sheetId, 5, 6, 140),   // F Confirmation #
    colWidthReq(sheetId, 6, 7, 80),    // G Account
    colWidthReq(sheetId, 7, 8, 220),   // H PDF Drive Link
    colWidthReq(sheetId, 8, 9, 220),   // I Notes
    colWidthReq(sheetId, 9, 10, 130),  // J Linked Txn Ref
  ];
  const styleRes = await spreadsheetsBatchUpdate(env, userId, styling);
  if (!styleRes.ok) {
    errors.push(`Created '${TITLE}' but styling failed: ${styleRes.error}`);
    // Continue — the tab exists and is usable, just plain.
  }

  // ── Step 3: populate banner, stats formulas, header row ──
  const valueWrites = [
    [`'${TITLE}'!A1`, [[`CRA REMITTANCES LOG  —  HST · Payroll source deductions · Corp tax`]]],
    [`'${TITLE}'!B2`, [['REMITTANCE TOTALS  (YTD)']]],
    [`'${TITLE}'!B3:F3`, [['HST paid to CRA',                    '=IFERROR(SUMIF(C12:C500,"HST",E12:E500),0)', '', '', '']]],
    [`'${TITLE}'!B4:F4`, [['Payroll source deductions paid',     '=IFERROR(SUMIF(C12:C500,"Payroll (PD7A)",E12:E500),0)', '', '', '']]],
    [`'${TITLE}'!B5:F5`, [['Corporate tax instalments paid',     '=IFERROR(SUMIF(C12:C500,"Corporate Tax Instalment",E12:E500),0)', '', '', '']]],
    [`'${TITLE}'!B6:F6`, [['Corporate tax (final) paid',         '=IFERROR(SUMIF(C12:C500,"Corporate Tax Final",E12:E500),0)', '', '', '']]],
    [`'${TITLE}'!B7:F7`, [['TOTAL paid to CRA',                  '=IFERROR(SUM(E12:E500),0)', '', '', '']]],
    [`'${TITLE}'!B8:F8`, [['Receipts attached (count)',          '=IFERROR(COUNTIF(H12:H500,"<>"),0)', '', '', '']]],
    [`'${TITLE}'!B9:F9`, [['Missing receipts (count)',           '=IFERROR(COUNTA(B12:B500)-COUNTIF(H12:H500,"<>"),0)', '', '', '']]],
    [`'${TITLE}'!B11:J11`, [[
      'Date Paid', 'Type', 'Period Covered', 'Amount', 'Confirmation #', 'Account',
      'PDF Receipt (Drive)', 'Notes', 'Linked Txn Ref',
    ]]],
  ];
  for (const [range, values] of valueWrites) {
    const res = await writeRange(env, userId, range, values);
    if (!res.ok) errors.push(`Failed to populate ${range}: ${res.error}`);
  }

  changes.push(`Created '${TITLE}' tab with stats formulas, headers, validation, and styling.`);
}

// ════════════════════════════════════════════════════════════════════
// Migration 2: Invoices tab — deposit columns O / P / Q
// ════════════════════════════════════════════════════════════════════
async function applyInvoiceDepositColumns(env, userId, sheetsByTitle, changes, errors, dryRun) {
  const TITLE = '🧾 Invoices';
  const inv = sheetsByTitle[TITLE];
  if (!inv) {
    errors.push(`'${TITLE}' tab missing — can't add deposit columns. Run a fresh setup or contact support.`);
    return;
  }

  const sheetId = inv.sheetId;
  const currentColCount = inv.gridProperties?.columnCount || 0;

  // The Invoices schema needs at least 17 columns (A gutter + B-Q data).
  if (currentColCount >= 17) {
    return;  // already migrated — nothing to do
  }

  if (dryRun) {
    changes.push(`Add deposit columns (Deposit Amount / Deposit Date Received / Balance Due) to '${TITLE}' — enables tracking down payments on jobs.`);
    return;
  }

  // ── Step 1: expand the grid to 17 columns ──
  const expandReq = {
    updateSheetProperties: {
      properties: { sheetId, gridProperties: { columnCount: 17, rowCount: 500, frozenRowCount: 11 }},
      fields: 'gridProperties.columnCount,gridProperties.rowCount,gridProperties.frozenRowCount',
    },
  };
  const expandRes = await spreadsheetsBatchUpdate(env, userId, [expandReq]);
  if (!expandRes.ok) {
    errors.push(`Could not expand '${TITLE}' grid: ${expandRes.error}`);
    return;
  }

  // ── Step 2: format new columns + extend header styling/banding/validation ──
  // Re-applying these to the original 14-col range is fine (idempotent), but
  // we only need to apply to the new 3 cols (O/P/Q) and update the status
  // dropdown to include the new states.
  const styling = [
    // Currency format on O (Deposit Amount)
    { repeatCell: {
        range: { sheetId, startRowIndex: 11, endRowIndex: 500, startColumnIndex: 14, endColumnIndex: 15 },
        cell: { userEnteredFormat: { numberFormat: FMT_CURRENCY.numberFormat }},
        fields: 'userEnteredFormat.numberFormat',
    }},
    // Date format on P (Deposit Date Received)
    { repeatCell: {
        range: { sheetId, startRowIndex: 11, endRowIndex: 500, startColumnIndex: 15, endColumnIndex: 16 },
        cell: { userEnteredFormat: { numberFormat: FMT_DATE.numberFormat }},
        fields: 'userEnteredFormat.numberFormat',
    }},
    // Currency format on Q (Balance Due)
    { repeatCell: {
        range: { sheetId, startRowIndex: 11, endRowIndex: 500, startColumnIndex: 16, endColumnIndex: 17 },
        cell: { userEnteredFormat: { numberFormat: FMT_CURRENCY.numberFormat }},
        fields: 'userEnteredFormat.numberFormat',
    }},
    // Header styling on O11:Q11 (teal background, white bold text)
    { repeatCell: {
        range: { sheetId, startRowIndex: 10, endRowIndex: 11, startColumnIndex: 14, endColumnIndex: 17 },
        cell: { userEnteredFormat: {
          backgroundColor: COLORS.blue,
          textFormat: { foregroundColor: COLORS.white, bold: true, fontSize: 10 },
          horizontalAlignment: 'LEFT',
          verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        }},
        fields: 'userEnteredFormat',
    }},
    // Status dropdown rebuild — extend with new statuses. Re-applying overwrites
    // the existing rule; we add the full list (idempotent).
    { setDataValidation: {
        range: { sheetId, startRowIndex: 11, endRowIndex: 500, startColumnIndex: 10, endColumnIndex: 11 },
        rule: { condition: { type: 'ONE_OF_LIST', values: [
          { userEnteredValue: 'Unpaid' },
          { userEnteredValue: 'Awaiting Deposit' },
          { userEnteredValue: 'Deposit Received' },
          { userEnteredValue: 'Paid' },
          { userEnteredValue: 'Overdue' },
          { userEnteredValue: 'Cancelled' },
        ]}, showCustomUi: true },
    }},
    // Column widths for O/P/Q
    colWidthReq(sheetId, 14, 15, 110),  // O Deposit Amount
    colWidthReq(sheetId, 15, 16, 130),  // P Deposit Date Received
    colWidthReq(sheetId, 16, 17, 110),  // Q Balance Due
  ];
  const styleRes = await spreadsheetsBatchUpdate(env, userId, styling);
  if (!styleRes.ok) {
    errors.push(`Expanded '${TITLE}' but styling failed: ${styleRes.error}`);
  }

  // ── Step 3: write the header labels in O11/P11/Q11 ──
  const headerRes = await writeRange(env, userId, `'${TITLE}'!O11:Q11`,
    [['Deposit Amount', 'Deposit Date Received', 'Balance Due']]);
  if (!headerRes.ok) {
    errors.push(`Failed to write deposit headers: ${headerRes.error}`);
  }

  // ── Step 4: refresh the stats formulas in B8:C9 to use the deposit-aware versions ──
  // Old: SUMIF on Status='Unpaid'. New: deposit-aware totals.
  const statsRes1 = await writeRange(env, userId, `'${TITLE}'!B8:C8`,
    [['Outstanding — not yet paid', '=SUM(H12:H500)-SUMIF(K12:K500,"Paid",H12:H500)-SUMIF(K12:K500,"Deposit Received",O12:O500)']]);
  if (!statsRes1.ok) errors.push(`Failed to update outstanding formula: ${statsRes1.error}`);

  const statsRes2 = await writeRange(env, userId, `'${TITLE}'!B9:C9`,
    [['Collected — paid + deposits', '=SUMIF(K12:K500,"Paid",H12:H500)+SUMIF(K12:K500,"Deposit Received",O12:O500)']]);
  if (!statsRes2.ok) errors.push(`Failed to update collected formula: ${statsRes2.error}`);

  changes.push(`Added deposit columns O/P/Q to '${TITLE}' tab and updated outstanding/collected formulas to be deposit-aware.`);
}

// ════════════════════════════════════════════════════════════════════
// Migration 3: Transactions tab — Total (incl HST) column N
// ════════════════════════════════════════════════════════════════════
// Adds a calculated column to the right of Match Status that shows the gross
// signed amount (Amount + HST in same direction). Lets users reconcile against
// bank statements directly: the bank shows the total charged, this column
// matches it. Safe to run on any user's Transactions tab (legacy emoji
// prefixes get resolved via title suffix matching).
async function applyTransactionsTotalColumn(env, userId, sheetsByTitle, changes, errors, dryRun) {
  // Resolve Transactions tab regardless of emoji prefix (📒 vs legacy variants)
  const txnTab = Object.values(sheetsByTitle).find(s => /transactions/i.test(s.title));
  if (!txnTab) {
    // No Transactions tab to migrate — not an error, just skip. The user
    // probably needs to rebuild from scratch with manualCreateSheet.
    return;
  }

  const sheetId = txnTab.sheetId;
  const currentColCount = txnTab.gridProperties?.columnCount || 0;

  // Need at least 14 columns (A gutter + B-N data). If the column already
  // exists AND has the header populated, skip — idempotent.
  if (currentColCount >= 14) {
    // Could check the header cell content to be extra-safe but reading every
    // sheet's header on every dryRun is wasteful. Trust column count.
    return;
  }

  if (dryRun) {
    changes.push(`Add 'Total (incl HST)' column to '${txnTab.title}' — auto-calculates gross signed amount so you can reconcile against bank statements directly.`);
    return;
  }

  // ── Step 1: expand the grid to 14 columns ──
  const expandReq = {
    updateSheetProperties: {
      properties: { sheetId, gridProperties: { columnCount: 14, rowCount: txnTab.gridProperties?.rowCount || 1000, frozenRowCount: 11 }},
      fields: 'gridProperties.columnCount,gridProperties.rowCount,gridProperties.frozenRowCount',
    },
  };
  const expandRes = await spreadsheetsBatchUpdate(env, userId, [expandReq]);
  if (!expandRes.ok) {
    errors.push(`Could not expand '${txnTab.title}' grid: ${expandRes.error}`);
    return;
  }

  // ── Step 2: format the new column (currency) + header styling + col width ──
  const styling = [
    // Currency format on N (Total)
    { repeatCell: {
        range: { sheetId, startRowIndex: 11, endRowIndex: 1000, startColumnIndex: 13, endColumnIndex: 14 },
        cell: { userEnteredFormat: { numberFormat: FMT_CURRENCY.numberFormat }},
        fields: 'userEnteredFormat.numberFormat',
    }},
    // Header styling on N11 (teal background, white bold)
    { repeatCell: {
        range: { sheetId, startRowIndex: 10, endRowIndex: 11, startColumnIndex: 13, endColumnIndex: 14 },
        cell: { userEnteredFormat: {
          backgroundColor: COLORS.teal,
          textFormat: { foregroundColor: COLORS.white, bold: true, fontSize: 10 },
          horizontalAlignment: 'LEFT',
          verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        }},
        fields: 'userEnteredFormat',
    }},
    colWidthReq(sheetId, 13, 14, 110),  // N Total
  ];
  const styleRes = await spreadsheetsBatchUpdate(env, userId, styling);
  if (!styleRes.ok) {
    errors.push(`Expanded '${txnTab.title}' but styling failed: ${styleRes.error}`);
  }

  // ── Step 3: write the header label in N11 ──
  const headerRes = await writeRange(env, userId, `'${txnTab.title}'!N11`, [['Total (incl HST)']]);
  if (!headerRes.ok) errors.push(`Failed to write Total header: ${headerRes.error}`);

  // ── Step 4: write the ARRAYFORMULA in N12 — populates all rows automatically ──
  const formulaRes = await writeRange(
    env, userId, `'${txnTab.title}'!N12`,
    [['=ARRAYFORMULA(IF(E12:E1000="","",E12:E1000+H12:H1000*SIGN(E12:E1000)))']]
  );
  if (!formulaRes.ok) errors.push(`Failed to write Total formula: ${formulaRes.error}`);

  changes.push(`Added 'Total (incl HST)' column N to '${txnTab.title}' — every row now shows the gross signed amount that matches what hit the bank.`);
}

// ════════════════════════════════════════════════════════════════════
// Migration 4: HST Returns — smart Fiscal Year Start (C3) formula
// ════════════════════════════════════════════════════════════════════
// Old formula defaulted C3 to today's FY-in-progress. If the user just
// imported transactions for the FY that JUST ENDED, the HST tab would show
// an empty quarterly window because it was looking at next year. New formula
// derives FY from the latest transaction date, falling back to today.
// User can still type any date in C3 manually to view past FYs.
async function applyHstReturnsSmartFyStart(env, userId, sheetsByTitle, changes, errors, dryRun) {
  const hstTab = Object.values(sheetsByTitle).find(s => /hst returns/i.test(s.title));
  if (!hstTab) return;  // No HST Returns tab — likely a partial sheet

  // We can't easily detect "is the formula already updated" without reading C3.
  // To stay idempotent + cheap on dry-run, just always rewrite both C3 and the
  // label. A re-run is harmless because the values are the same.
  if (dryRun) {
    changes.push(`Update '${hstTab.title}' C3 to auto-detect Fiscal Year from your latest transaction (was hardcoded to today's FY, which broke when importing prior-year data).`);
    return;
  }

  const labelRes = await writeRange(env, userId, `'${hstTab.title}'!B3`,
    [['Fiscal Year Start (auto-detected — type a date here to view a different FY):']]);
  if (!labelRes.ok) errors.push(`Failed to update HST FY label: ${labelRes.error}`);

  // The formula references the Transactions tab — resolve its actual title
  // (handles legacy emoji prefixes).
  const txnTab = Object.values(sheetsByTitle).find(s => /transactions/i.test(s.title));
  const txnTitle = txnTab ? txnTab.title : '📒 Transactions';
  const formula = `=IFERROR(DATE(YEAR(MAX('${txnTitle}'!B12:B1000))-IF(MONTH(MAX('${txnTitle}'!B12:B1000))<4,1,0),4,1),DATE(YEAR(TODAY())-IF(MONTH(TODAY())<4,1,0),4,1))`;
  const formulaRes = await writeRange(env, userId, `'${hstTab.title}'!C3`, [[formula]]);
  if (!formulaRes.ok) errors.push(`Failed to update HST FY formula: ${formulaRes.error}`);

  changes.push(`Updated '${hstTab.title}' C3 to auto-detect Fiscal Year from your latest transaction.`);
}

// ════════════════════════════════════════════════════════════════════
// Migration 5: 🏦 Account Balances tab (bank reconciliation)
// ════════════════════════════════════════════════════════════════════
// Adds the per-account-per-period reconciliation tab. User enters opening
// balance + actual closing from each bank statement. Sheet computes expected
// closing from Transactions and flags any difference. Catches missed /
// duplicated / wrong-sign rows automatically. Foundational for "MNP can
// trust these books" — a reconciled book is a believable book.
async function applyAccountBalancesTab(env, userId, sheetsByTitle, changes, errors, dryRun) {
  const TITLE = '🏦 Account Balances';
  if (sheetsByTitle[TITLE]) return;  // already exists

  if (dryRun) {
    changes.push(`Add '${TITLE}' tab — bank reconciliation: enter opening + closing from each statement, sheet flags any discrepancy with your imported transactions.`);
    return;
  }

  // Pick a sheetId that doesn't collide
  const usedIds = new Set(Object.values(sheetsByTitle).map(s => s.sheetId));
  let sheetId = 1100;
  while (usedIds.has(sheetId)) sheetId += 100;

  // Step 1: create the tab
  const addSheetReq = {
    addSheet: {
      properties: {
        sheetId, title: TITLE,
        index: Object.keys(sheetsByTitle).length,
        gridProperties: { rowCount: 500, columnCount: 13, frozenRowCount: 11 },
        tabColor: COLORS.teal,  // close enough to green; using existing palette
      },
    },
  };
  const addRes = await spreadsheetsBatchUpdate(env, userId, [addSheetReq]);
  if (!addRes.ok) {
    errors.push(`Could not create '${TITLE}' tab: ${addRes.error}`);
    return;
  }

  // Step 2: styling — banner, section, header, banding, currency/date formats
  const styling = [
    { repeatCell: {
        range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 13 },
        cell: { userEnteredFormat: {
          backgroundColor: COLORS.teal,
          textFormat: { foregroundColor: COLORS.white, bold: true, fontSize: 12 },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
          padding: { top: 8, left: 12, right: 12, bottom: 8 },
        }},
        fields: 'userEnteredFormat',
    }},
    { mergeCells: {
        range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 13 },
        mergeType: 'MERGE_ALL',
    }},
    { updateDimensionProperties: {
        range: { sheetId, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
        properties: { pixelSize: 36 }, fields: 'pixelSize',
    }},
    // Section row label
    { repeatCell: {
        range: { sheetId, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 1, endColumnIndex: 7 },
        cell: { userEnteredFormat: {
          backgroundColor: COLORS.teal,
          textFormat: { foregroundColor: COLORS.white, bold: true, fontSize: 10 },
        }},
        fields: 'userEnteredFormat',
    }},
    // Stats labels (rows 2-7, col B) bold
    { repeatCell: {
        range: { sheetId, startRowIndex: 2, endRowIndex: 8, startColumnIndex: 1, endColumnIndex: 2 },
        cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 10 } }},
        fields: 'userEnteredFormat.textFormat',
    }},
    // Header row (10)
    { repeatCell: {
        range: { sheetId, startRowIndex: 10, endRowIndex: 11, startColumnIndex: 1, endColumnIndex: 13 },
        cell: { userEnteredFormat: {
          backgroundColor: COLORS.teal,
          textFormat: { foregroundColor: COLORS.white, bold: true, fontSize: 10 },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        }},
        fields: 'userEnteredFormat',
    }},
    { updateDimensionProperties: {
        range: { sheetId, dimension: 'ROWS', startIndex: 10, endIndex: 11 },
        properties: { pixelSize: 36 }, fields: 'pixelSize',
    }},
    { addBanding: {
        bandedRange: {
          range: { sheetId, startRowIndex: 11, endRowIndex: 500, startColumnIndex: 1, endColumnIndex: 13 },
          rowProperties: { firstBandColor: COLORS.white, secondBandColor: COLORS.tealTint },
        },
    }},
    // Date columns D + E
    { repeatCell: {
        range: { sheetId, startRowIndex: 11, endRowIndex: 500, startColumnIndex: 3, endColumnIndex: 5 },
        cell: { userEnteredFormat: { numberFormat: FMT_DATE.numberFormat }},
        fields: 'userEnteredFormat.numberFormat',
    }},
    // Currency columns F-J
    { repeatCell: {
        range: { sheetId, startRowIndex: 11, endRowIndex: 500, startColumnIndex: 5, endColumnIndex: 10 },
        cell: { userEnteredFormat: { numberFormat: FMT_CURRENCY.numberFormat }},
        fields: 'userEnteredFormat.numberFormat',
    }},
    colWidthReq(sheetId, 0, 1, 30),    // A gutter
    colWidthReq(sheetId, 1, 2, 110),   // B Period
    colWidthReq(sheetId, 2, 3, 100),   // C Account
    colWidthReq(sheetId, 3, 4, 100),   // D Period Start
    colWidthReq(sheetId, 4, 5, 100),   // E Period End
    colWidthReq(sheetId, 5, 6, 110),   // F Opening
    colWidthReq(sheetId, 6, 7, 110),   // G Sum Activity (formula)
    colWidthReq(sheetId, 7, 8, 130),   // H Expected Closing (formula)
    colWidthReq(sheetId, 8, 9, 130),   // I Actual Closing
    colWidthReq(sheetId, 9, 10, 110),  // J Difference (formula)
    colWidthReq(sheetId, 10, 11, 110), // K Match (formula)
    colWidthReq(sheetId, 11, 12, 200), // L Notes
  ];
  const styleRes = await spreadsheetsBatchUpdate(env, userId, styling);
  if (!styleRes.ok) errors.push(`Created '${TITLE}' but styling failed: ${styleRes.error}`);

  // Step 3: populate values + formulas (resolve Transactions tab name for legacy compat)
  const txnTab = Object.values(sheetsByTitle).find(s => /transactions/i.test(s.title));
  const txnTitle = txnTab ? txnTab.title : '📒 Transactions';

  const writes = [
    [`'${TITLE}'!A1`, [['BANK RECONCILIATION  —  Opening + activity = expected closing  ·  Compare to actual to catch errors']]],
    [`'${TITLE}'!B2`, [['RECONCILIATION SUMMARY']]],
    [`'${TITLE}'!B3:F3`, [['Periods reconciled', '=COUNTA(B12:B500)', '', '', '']]],
    [`'${TITLE}'!B4:F4`, [['Periods balanced (✓)', '=COUNTIF(K12:K500,"✓ Balanced")', '', '', '']]],
    [`'${TITLE}'!B5:F5`, [['Periods OFF (action needed)', '=COUNTIF(K12:K500,"⚠*")', '', '', '']]],
    [`'${TITLE}'!B6:F6`, [['Total off-by-amount across periods', '=SUMIF(K12:K500,"⚠*",J12:J500)', '', '', '']]],
    [`'${TITLE}'!B7:F7`, [['How to use this tab', 'Enter period dates + opening + closing balance from each statement. The sheet computes expected closing from your Transactions and flags any difference. A non-zero difference = a missed row, duplicate, wrong sign, or bad amount somewhere.', '', '', '']]],
    [`'${TITLE}'!B11:L11`, [[
      'Period', 'Account', 'Period Start', 'Period End', 'Opening Balance',
      'Sum Activity (auto)', 'Expected Closing (auto)', 'Actual Closing',
      'Difference (auto)', 'Match (auto)', 'Notes',
    ]]],
    [`'${TITLE}'!G12`, [[
      `=ARRAYFORMULA(IF(C12:C500="","",IFERROR(SUMIFS('${txnTitle}'!N12:N1000,'${txnTitle}'!I12:I1000,C12:C500,'${txnTitle}'!B12:B1000,">="&D12:D500,'${txnTitle}'!B12:B1000,"<="&E12:E500),0)))`
    ]]],
    [`'${TITLE}'!H12`, [['=ARRAYFORMULA(IF(C12:C500="","",F12:F500+G12:G500))']]],
    [`'${TITLE}'!J12`, [['=ARRAYFORMULA(IF(I12:I500="","",H12:H500-I12:I500))']]],
    [`'${TITLE}'!K12`, [['=ARRAYFORMULA(IF(I12:I500="","",IF(ABS(J12:J500)<0.01,"✓ Balanced","⚠ Off by $"&TEXT(ROUND(J12:J500,2),"0.00"))))']]],
  ];
  for (const [range, values] of writes) {
    const res = await writeRange(env, userId, range, values);
    if (!res.ok) errors.push(`Failed to populate ${range}: ${res.error}`);
  }

  changes.push(`Added '${TITLE}' tab — enter opening + closing from each statement, sheet flags any discrepancy with your imported transactions.`);
}

// ════════════════════════════════════════════════════════════════════
// Migration 6: Year-End per-category breakdown (QUERY pivot)
// ════════════════════════════════════════════════════════════════════
// Adds a dynamic per-category breakdown to the Year-End tab. Uses QUERY to
// pivot Transactions by category and sum the Total (incl HST) column. Captures
// every category in use including user-custom ones, sorted biggest first.
// Replicates the column-totals visual her old spreadsheet had.
async function applyYearEndPerCategoryBreakdown(env, userId, sheetsByTitle, changes, errors, dryRun) {
  const yeTab = Object.values(sheetsByTitle).find(s => /year.?end/i.test(s.title));
  if (!yeTab) return;  // No Year-End tab — partial sheet

  // Need to make sure the grid is large enough for the QUERY rows.
  // We don't have a clean "is the breakdown already there" check without
  // reading A62, so re-running is harmless (overwrites with same content).
  if (dryRun) {
    changes.push(`Add per-category breakdown to '${yeTab.title}' — auto-populated table showing every category in use with row count + total. Mirrors the column-totals layout you used before.`);
    return;
  }

  // Step 1: ensure grid is big enough (need at least 200 rows for the QUERY
  // output). Older Year-End tabs were 80 rows.
  const currentRowCount = yeTab.gridProperties?.rowCount || 80;
  if (currentRowCount < 200) {
    const expandReq = {
      updateSheetProperties: {
        properties: { sheetId: yeTab.sheetId, gridProperties: { rowCount: 200, columnCount: yeTab.gridProperties?.columnCount || 8 }},
        fields: 'gridProperties.rowCount,gridProperties.columnCount',
      },
    };
    const expandRes = await spreadsheetsBatchUpdate(env, userId, [expandReq]);
    if (!expandRes.ok) errors.push(`Could not expand '${yeTab.title}' to 200 rows: ${expandRes.error}`);
  }

  // Step 2: write the section header + QUERY formula. Resolve Transactions
  // tab name for legacy compat.
  const txnTab = Object.values(sheetsByTitle).find(s => /transactions/i.test(s.title));
  const txnTitle = txnTab ? txnTab.title : '📒 Transactions';

  const headerRes = await writeRange(env, userId, `'${yeTab.title}'!A62`,
    [['  PER-CATEGORY BREAKDOWN  (all-time, every category in use, biggest first)']]);
  if (!headerRes.ok) errors.push(`Failed to write breakdown header: ${headerRes.error}`);

  const formula = `=IFERROR(QUERY('${txnTitle}'!B12:N1000, "SELECT F, COUNT(F), SUM(N) WHERE F IS NOT NULL AND F <> '' AND F <> 'Internal Transfer' GROUP BY F ORDER BY SUM(N) DESC LABEL F 'Category', COUNT(F) '# of rows', SUM(N) 'Total (incl HST)'", 0), "No transactions yet — import a statement to populate this breakdown.")`;
  const formulaRes = await writeRange(env, userId, `'${yeTab.title}'!B63`, [[formula]]);
  if (!formulaRes.ok) errors.push(`Failed to write breakdown formula: ${formulaRes.error}`);

  changes.push(`Added per-category breakdown to '${yeTab.title}' — see PER-CATEGORY BREAKDOWN section near the bottom for auto-populated category totals.`);
}

// ════════════════════════════════════════════════════════════════════
// Migration 7: 📓 Adjusting Entries tab (year-end accruals)
// ════════════════════════════════════════════════════════════════════
// Adds the linchpin tab for accrual adjustments at year-end. User enters one
// row per adjustment (AR, AP, prepaid, accrued). Year-End summary picks them
// up so cash-basis books become T2-ready. Without this tab, MNP has to make
// the entries themselves — which means full prep, not review-only.
async function applyAdjustingEntriesTab(env, userId, sheetsByTitle, changes, errors, dryRun) {
  const TITLE = '📓 Adjusting Entries';
  if (sheetsByTitle[TITLE]) return;

  if (dryRun) {
    changes.push(`Add '${TITLE}' tab — log year-end accrual adjustments (AR, AP, prepaids) so the books become accrual-ready for the T2.`);
    return;
  }

  const usedIds = new Set(Object.values(sheetsByTitle).map(s => s.sheetId));
  let sheetId = 1200;
  while (usedIds.has(sheetId)) sheetId += 100;

  // Create tab
  const addRes = await spreadsheetsBatchUpdate(env, userId, [{
    addSheet: {
      properties: {
        sheetId, title: TITLE,
        index: Object.keys(sheetsByTitle).length,
        gridProperties: { rowCount: 200, columnCount: 13, frozenRowCount: 11 },
        tabColor: COLORS.teal,
      },
    },
  }]);
  if (!addRes.ok) {
    errors.push(`Could not create '${TITLE}': ${addRes.error}`);
    return;
  }

  // Styling
  const styling = [
    { repeatCell: {
        range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 13 },
        cell: { userEnteredFormat: {
          backgroundColor: COLORS.teal,
          textFormat: { foregroundColor: COLORS.white, bold: true, fontSize: 12 },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
          padding: { top: 8, left: 12, right: 12, bottom: 8 },
        }},
        fields: 'userEnteredFormat',
    }},
    { mergeCells: { range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 13 }, mergeType: 'MERGE_ALL' }},
    { updateDimensionProperties: { range: { sheetId, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' }},
    { repeatCell: {
        range: { sheetId, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 1, endColumnIndex: 7 },
        cell: { userEnteredFormat: { backgroundColor: COLORS.teal, textFormat: { foregroundColor: COLORS.white, bold: true, fontSize: 10 }}},
        fields: 'userEnteredFormat',
    }},
    { repeatCell: {
        range: { sheetId, startRowIndex: 2, endRowIndex: 10, startColumnIndex: 1, endColumnIndex: 2 },
        cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 10 }}},
        fields: 'userEnteredFormat.textFormat',
    }},
    { repeatCell: {
        range: { sheetId, startRowIndex: 10, endRowIndex: 11, startColumnIndex: 1, endColumnIndex: 13 },
        cell: { userEnteredFormat: {
          backgroundColor: COLORS.teal,
          textFormat: { foregroundColor: COLORS.white, bold: true, fontSize: 10 },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
        }},
        fields: 'userEnteredFormat',
    }},
    { updateDimensionProperties: { range: { sheetId, dimension: 'ROWS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 36 }, fields: 'pixelSize' }},
    { addBanding: {
        bandedRange: {
          range: { sheetId, startRowIndex: 11, endRowIndex: 200, startColumnIndex: 1, endColumnIndex: 13 },
          rowProperties: { firstBandColor: COLORS.white, secondBandColor: COLORS.tealTint },
        },
    }},
    { repeatCell: {
        range: { sheetId, startRowIndex: 11, endRowIndex: 200, startColumnIndex: 1, endColumnIndex: 2 },
        cell: { userEnteredFormat: { numberFormat: FMT_DATE.numberFormat }},
        fields: 'userEnteredFormat.numberFormat',
    }},
    { repeatCell: {
        range: { sheetId, startRowIndex: 11, endRowIndex: 200, startColumnIndex: 5, endColumnIndex: 8 },
        cell: { userEnteredFormat: { numberFormat: FMT_CURRENCY.numberFormat }},
        fields: 'userEnteredFormat.numberFormat',
    }},
    { setDataValidation: {
        range: { sheetId, startRowIndex: 11, endRowIndex: 200, startColumnIndex: 2, endColumnIndex: 3 },
        rule: { condition: { type: 'ONE_OF_LIST', values: [
          { userEnteredValue: 'Accounts Receivable (AR)' },
          { userEnteredValue: 'Accounts Payable (AP)' },
          { userEnteredValue: 'Prepaid Expense' },
          { userEnteredValue: 'Accrued Expense' },
          { userEnteredValue: 'Accrued Revenue' },
          { userEnteredValue: 'Depreciation (CCA)' },
          { userEnteredValue: 'Other Adjustment' },
        ]}, showCustomUi: true },
    }},
    colWidthReq(sheetId, 0, 1, 30),
    colWidthReq(sheetId, 1, 2, 100),
    colWidthReq(sheetId, 2, 3, 180),
    colWidthReq(sheetId, 3, 4, 280),
    colWidthReq(sheetId, 4, 5, 150),
    colWidthReq(sheetId, 5, 6, 110),
    colWidthReq(sheetId, 6, 7, 100),
    colWidthReq(sheetId, 7, 8, 110),
    colWidthReq(sheetId, 8, 9, 200),
    colWidthReq(sheetId, 9, 10, 130),
    colWidthReq(sheetId, 10, 11, 200),
    colWidthReq(sheetId, 11, 12, 80),
  ];
  const styleRes = await spreadsheetsBatchUpdate(env, userId, styling);
  if (!styleRes.ok) errors.push(`Created '${TITLE}' but styling failed: ${styleRes.error}`);

  // Populate values
  const writes = [
    [`'${TITLE}'!A1`, [['YEAR-END ADJUSTING ENTRIES  —  Convert cash-basis books to accrual-basis for the T2  ·  AR / AP / Prepaids / Accruals']]],
    [`'${TITLE}'!B2`, [['ADJUSTMENT TOTALS  (by type)']]],
    [`'${TITLE}'!B3:F3`, [['Accounts Receivable (revenue to add)', '=SUMIF(C12:C200,"Accounts Receivable (AR)",F12:F200)', '', '', '']]],
    [`'${TITLE}'!B4:F4`, [['Accounts Payable (expenses to add)', '=SUMIF(C12:C200,"Accounts Payable (AP)",F12:F200)', '', '', '']]],
    [`'${TITLE}'!B5:F5`, [['Prepaid Expenses (deferred to next FY)', '=SUMIF(C12:C200,"Prepaid Expense",F12:F200)', '', '', '']]],
    [`'${TITLE}'!B6:F6`, [['Accrued Expenses (booked, not yet paid)', '=SUMIF(C12:C200,"Accrued Expense",F12:F200)', '', '', '']]],
    [`'${TITLE}'!B7:F7`, [['Accrued Revenue (earned, not yet billed)', '=SUMIF(C12:C200,"Accrued Revenue",F12:F200)', '', '', '']]],
    [`'${TITLE}'!B8:F8`, [['Depreciation (CCA from Fixed Assets)', '=SUMIF(C12:C200,"Depreciation (CCA)",F12:F200)', '', '', '']]],
    [`'${TITLE}'!B9:F9`, [['Total adjustments', '=SUM(F12:F200)', '', '', '']]],
    [`'${TITLE}'!B11:L11`, [[
      'Date', 'Type', 'Description', 'Counterparty', 'Net Amount',
      'HST', 'Total (auto)', 'Effect on Books', 'Linked Invoice / Bill', 'Notes', 'FY',
    ]]],
    [`'${TITLE}'!H12`, [['=ARRAYFORMULA(IF(F12:F200="","",F12:F200+G12:G200*SIGN(F12:F200)))']]],
  ];
  for (const [range, values] of writes) {
    const res = await writeRange(env, userId, range, values);
    if (!res.ok) errors.push(`Failed to populate ${range}: ${res.error}`);
  }

  changes.push(`Added '${TITLE}' tab — log year-end accrual adjustments to make the books T2-ready (the linchpin for MNP review-only engagement).`);
}

// ════════════════════════════════════════════════════════════════════
// Migration 8: 🛠 Fixed Assets tab (CCA tracker)
// ════════════════════════════════════════════════════════════════════
// Adds the CCA tracker for fixed assets (vehicles, tools, equipment, etc).
// Tracks UCC year-over-year, computes annual CCA for the T2 Schedule 8.
// Pre-loaded with common CCA classes for trades businesses.
async function applyFixedAssetsTab(env, userId, sheetsByTitle, changes, errors, dryRun) {
  const TITLE = '🛠 Fixed Assets';
  if (sheetsByTitle[TITLE]) return;

  if (dryRun) {
    changes.push(`Add '${TITLE}' tab — track capital assets (trucks, tools, equipment) and compute annual CCA for the T2 Schedule 8.`);
    return;
  }

  const usedIds = new Set(Object.values(sheetsByTitle).map(s => s.sheetId));
  let sheetId = 1300;
  while (usedIds.has(sheetId)) sheetId += 100;

  const addRes = await spreadsheetsBatchUpdate(env, userId, [{
    addSheet: {
      properties: {
        sheetId, title: TITLE,
        index: Object.keys(sheetsByTitle).length,
        gridProperties: { rowCount: 200, columnCount: 13, frozenRowCount: 11 },
        tabColor: COLORS.blue,
      },
    },
  }]);
  if (!addRes.ok) {
    errors.push(`Could not create '${TITLE}': ${addRes.error}`);
    return;
  }

  // Styling
  const styling = [
    { repeatCell: {
        range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 13 },
        cell: { userEnteredFormat: {
          backgroundColor: COLORS.blue,
          textFormat: { foregroundColor: COLORS.white, bold: true, fontSize: 12 },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
          padding: { top: 8, left: 12, right: 12, bottom: 8 },
        }},
        fields: 'userEnteredFormat',
    }},
    { mergeCells: { range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 13 }, mergeType: 'MERGE_ALL' }},
    { updateDimensionProperties: { range: { sheetId, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' }},
    { repeatCell: {
        range: { sheetId, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 1, endColumnIndex: 7 },
        cell: { userEnteredFormat: { backgroundColor: COLORS.blue, textFormat: { foregroundColor: COLORS.white, bold: true, fontSize: 10 }}},
        fields: 'userEnteredFormat',
    }},
    { repeatCell: {
        range: { sheetId, startRowIndex: 2, endRowIndex: 10, startColumnIndex: 1, endColumnIndex: 2 },
        cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 10 }}},
        fields: 'userEnteredFormat.textFormat',
    }},
    { repeatCell: {
        range: { sheetId, startRowIndex: 10, endRowIndex: 11, startColumnIndex: 1, endColumnIndex: 13 },
        cell: { userEnteredFormat: {
          backgroundColor: COLORS.blue,
          textFormat: { foregroundColor: COLORS.white, bold: true, fontSize: 10 },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
        }},
        fields: 'userEnteredFormat',
    }},
    { updateDimensionProperties: { range: { sheetId, dimension: 'ROWS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 36 }, fields: 'pixelSize' }},
    { addBanding: {
        bandedRange: {
          range: { sheetId, startRowIndex: 11, endRowIndex: 200, startColumnIndex: 1, endColumnIndex: 13 },
          rowProperties: { firstBandColor: COLORS.white, secondBandColor: COLORS.blueTint },
        },
    }},
    { repeatCell: { range: { sheetId, startRowIndex: 11, endRowIndex: 200, startColumnIndex: 3, endColumnIndex: 4 }, cell: { userEnteredFormat: { numberFormat: FMT_DATE.numberFormat }}, fields: 'userEnteredFormat.numberFormat' }},
    { repeatCell: { range: { sheetId, startRowIndex: 11, endRowIndex: 200, startColumnIndex: 4, endColumnIndex: 7 }, cell: { userEnteredFormat: { numberFormat: FMT_CURRENCY.numberFormat }}, fields: 'userEnteredFormat.numberFormat' }},
    { repeatCell: { range: { sheetId, startRowIndex: 11, endRowIndex: 200, startColumnIndex: 7, endColumnIndex: 8 }, cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' }}}, fields: 'userEnteredFormat.numberFormat' }},
    { repeatCell: { range: { sheetId, startRowIndex: 11, endRowIndex: 200, startColumnIndex: 8, endColumnIndex: 10 }, cell: { userEnteredFormat: { numberFormat: FMT_CURRENCY.numberFormat }}, fields: 'userEnteredFormat.numberFormat' }},
    { setDataValidation: {
        range: { sheetId, startRowIndex: 11, endRowIndex: 200, startColumnIndex: 1, endColumnIndex: 2 },
        rule: { condition: { type: 'ONE_OF_LIST', values: [
          { userEnteredValue: 'Class 8 (20%) — Tools, equipment, furniture' },
          { userEnteredValue: 'Class 10 (30%) — Vehicles, work trucks <$36k' },
          { userEnteredValue: 'Class 10.1 (30%) — Luxury vehicles ≥$36k' },
          { userEnteredValue: 'Class 12 (100%) — Small tools <$500, software' },
          { userEnteredValue: 'Class 14.1 (5%) — Goodwill, intangibles' },
          { userEnteredValue: 'Class 16 (40%) — Heavy trucks >11,788kg' },
          { userEnteredValue: 'Class 50 (55%) — Computers, computer software' },
          { userEnteredValue: 'Class 53 (50%) — Manufacturing equipment' },
          { userEnteredValue: 'Class 6 (10%) — Wood-frame buildings' },
          { userEnteredValue: 'Class 1 (4%) — Other buildings' },
          { userEnteredValue: 'Other (enter rate manually)' },
        ]}, showCustomUi: true },
    }},
    { setDataValidation: {
        range: { sheetId, startRowIndex: 11, endRowIndex: 200, startColumnIndex: 10, endColumnIndex: 11 },
        rule: { condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: 'No' }, { userEnteredValue: 'Yes' }]}, showCustomUi: true },
    }},
    colWidthReq(sheetId, 0, 1, 30),
    colWidthReq(sheetId, 1, 2, 280),
    colWidthReq(sheetId, 2, 3, 220),
    colWidthReq(sheetId, 3, 4, 110),
    colWidthReq(sheetId, 4, 5, 110),
    colWidthReq(sheetId, 5, 6, 110),
    colWidthReq(sheetId, 6, 7, 110),
    colWidthReq(sheetId, 7, 8, 80),
    colWidthReq(sheetId, 8, 9, 110),
    colWidthReq(sheetId, 9, 10, 110),
    colWidthReq(sheetId, 10, 11, 80),
    colWidthReq(sheetId, 11, 12, 200),
  ];
  const styleRes = await spreadsheetsBatchUpdate(env, userId, styling);
  if (!styleRes.ok) errors.push(`Created '${TITLE}' but styling failed: ${styleRes.error}`);

  // Populate values + formulas
  const writes = [
    [`'${TITLE}'!A1`, [['FIXED ASSETS  —  CCA Tracker  ·  Capital Cost Allowance for T2 Schedule 8']]],
    [`'${TITLE}'!B2`, [['CCA SUMMARY  (this fiscal year)']]],
    [`'${TITLE}'!B3:F3`, [['Total assets tracked', '=COUNTA(B12:B200)', '', '', '']]],
    [`'${TITLE}'!B4:F4`, [['Total Original Cost (ever)', '=SUM(E12:E200)', '', '', '']]],
    [`'${TITLE}'!B5:F5`, [['Total Opening UCC (this FY)', '=SUM(F12:F200)', '', '', '']]],
    [`'${TITLE}'!B6:F6`, [['Total Additions This FY', '=SUM(G12:G200)', '', '', '']]],
    [`'${TITLE}'!B7:F7`, [['Total CCA This FY (T2 deduction)', '=SUM(I12:I200)', '', '', '']]],
    [`'${TITLE}'!B8:F8`, [['Total Ending UCC (carries to next FY)', '=SUM(J12:J200)', '', '', '']]],
    [`'${TITLE}'!B9:F9`, [['How to use', "Add one row per asset. At each year end: enter Opening UCC + any new Additions. CCA + Ending UCC compute automatically. Carry Ending UCC into next year's Opening UCC.", '', '', '']]],
    [`'${TITLE}'!B11:L11`, [[
      'CCA Class', 'Description', 'Date Acquired', 'Original Cost', 'Opening UCC',
      'Additions This FY', 'CCA Rate', 'CCA This FY', 'Ending UCC', 'Disposed?', 'Notes',
    ]]],
    [`'${TITLE}'!H12`, [['=ARRAYFORMULA(IF(B12:B200="","",IFERROR(VALUE(REGEXEXTRACT(B12:B200,"\\((\\d+)%\\)"))/100,0)))']]],
    [`'${TITLE}'!I12`, [['=ARRAYFORMULA(IF(B12:B200="","",IF(K12:K200="Yes",0, ROUND((F12:F200 + G12:G200/2) * H12:H200, 2))))']]],
    [`'${TITLE}'!J12`, [['=ARRAYFORMULA(IF(B12:B200="","",F12:F200 + G12:G200 - I12:I200))']]],
  ];
  for (const [range, values] of writes) {
    const res = await writeRange(env, userId, range, values);
    if (!res.ok) errors.push(`Failed to populate ${range}: ${res.error}`);
  }

  changes.push(`Added '${TITLE}' tab — track capital assets and compute CCA automatically for the T2 Schedule 8.`);
}

// ════════════════════════════════════════════════════════════════════
// Migration 9: 📊 T2 Worksheet (consolidated T2-prep view)
// ════════════════════════════════════════════════════════════════════
// The deliverable MNP reviews. Pulls from every other tab to produce the
// numbers MNP needs to file the T2 — no transcription required, just review.
async function applyT2Worksheet(env, userId, sheetsByTitle, changes, errors, dryRun) {
  const TITLE = '📊 T2 Worksheet';
  if (sheetsByTitle[TITLE]) return;

  if (dryRun) {
    changes.push(`Add '${TITLE}' tab — consolidated T2-prep view that MNP reviews instead of preparing themselves. Pulls Income Statement, Schedule 1 adjustments, CCA summary, and rough Balance Sheet from your other tabs.`);
    return;
  }

  const usedIds = new Set(Object.values(sheetsByTitle).map(s => s.sheetId));
  let sheetId = 1400;
  while (usedIds.has(sheetId)) sheetId += 100;

  // Resolve other tab names for legacy compat
  const txnTab = Object.values(sheetsByTitle).find(s => /transactions/i.test(s.title));
  const txnTitle = txnTab ? txnTab.title : '📒 Transactions';

  const addRes = await spreadsheetsBatchUpdate(env, userId, [{
    addSheet: {
      properties: {
        sheetId, title: TITLE,
        index: Object.keys(sheetsByTitle).length,
        gridProperties: { rowCount: 100, columnCount: 8 },
        tabColor: COLORS.teal,
      },
    },
  }]);
  if (!addRes.ok) {
    errors.push(`Could not create '${TITLE}': ${addRes.error}`);
    return;
  }

  // Light styling
  const styling = [
    { repeatCell: {
        range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 8 },
        cell: { userEnteredFormat: {
          backgroundColor: COLORS.teal,
          textFormat: { foregroundColor: COLORS.white, bold: true, fontSize: 12 },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
          padding: { top: 8, left: 12, right: 12, bottom: 8 },
        }},
        fields: 'userEnteredFormat',
    }},
    { mergeCells: { range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 8 }, mergeType: 'MERGE_ALL' }},
    { updateDimensionProperties: { range: { sheetId, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' }},
    { repeatCell: { range: { sheetId, startRowIndex: 1, endRowIndex: 99, startColumnIndex: 2, endColumnIndex: 3 }, cell: { userEnteredFormat: { numberFormat: FMT_CURRENCY.numberFormat, horizontalAlignment: 'RIGHT' }}, fields: 'userEnteredFormat.numberFormat,userEnteredFormat.horizontalAlignment' }},
    colWidthReq(sheetId, 0, 1, 30),
    colWidthReq(sheetId, 1, 2, 360),
    colWidthReq(sheetId, 2, 3, 130),
    colWidthReq(sheetId, 3, 4, 90),
    colWidthReq(sheetId, 4, 5, 240),
  ];
  await spreadsheetsBatchUpdate(env, userId, styling);

  // Load business name + FYE for the banner
  let bizName = 'Your Business', fye = 'March 31';
  try {
    const profileRow = await env.DB.prepare('SELECT business_name, fiscal_year_end FROM profiles WHERE user_id = ?')
      .bind(userId).first();
    if (profileRow) {
      bizName = profileRow.business_name || bizName;
      fye = profileRow.fiscal_year_end || fye;
    }
  } catch (e) { /* fall back to defaults */ }

  // Populate values + formulas
  const writes = [
    [`'${TITLE}'!A1`, [[`T2 WORKSHEET  —  ${bizName}  ·  Year-end ${fye}  ·  Hand this to MNP for review`]]],

    // Schedule 125 — Income Statement
    [`'${TITLE}'!A3`, [['  SCHEDULE 125 — INCOME STATEMENT (GIFI)']]],
    [`'${TITLE}'!B5:E5`, [['REVENUE', '', '', '']]],
    [`'${TITLE}'!B6:E10`, [
      ['Sales/services revenue (cash basis)',
        `=SUMIFS('${txnTitle}'!E12:E1000,'${txnTitle}'!E12:E1000,">0",'${txnTitle}'!F12:F1000,"<>Internal Transfer")`,
        '8089', '← Total positive Transactions excluding Internal Transfer'],
      ['Add: Accrued Revenue (FYE adjustments)',
        `=IFERROR(SUMIF('📓 Adjusting Entries'!C12:C200,"Accrued Revenue",'📓 Adjusting Entries'!F12:F200)+SUMIF('📓 Adjusting Entries'!C12:C200,"Accounts Receivable (AR)",'📓 Adjusting Entries'!F12:F200),0)`,
        '8089', '← AR + Accrued Revenue from Adjusting Entries'],
      ['TOTAL REVENUE (accrual basis)', '=C6+C7', '8299', '← For Schedule 125 line 8299'],
      ['', '', '', ''], ['', '', '', ''],
    ]],
    [`'${TITLE}'!B12:E12`, [['EXPENSES', '', '', '']]],
    [`'${TITLE}'!B13:E25`, [
      ['Total operating expenses (cash basis)',
        `=-SUMIFS('${txnTitle}'!E12:E1000,'${txnTitle}'!E12:E1000,"<0",'${txnTitle}'!F12:F1000,"<>Internal Transfer")`,
        'multiple', '← Total negative Transactions excluding Internal Transfer'],
      ['Add: Accrued Expenses + AP (FYE adjustments)',
        `=IFERROR(SUMIF('📓 Adjusting Entries'!C12:C200,"Accrued Expense",'📓 Adjusting Entries'!F12:F200)+SUMIF('📓 Adjusting Entries'!C12:C200,"Accounts Payable (AP)",'📓 Adjusting Entries'!F12:F200),0)`,
        'multiple', '← AP + Accrued Expenses from Adjusting Entries'],
      ['Less: Prepaid Expenses (defer to next FY)',
        `=IFERROR(SUMIF('📓 Adjusting Entries'!C12:C200,"Prepaid Expense",'📓 Adjusting Entries'!F12:F200),0)`,
        'adj', '← Negative entries on Adjusting tab reduce expense'],
      ['CCA (depreciation per Schedule 8)',
        `=IFERROR(SUM('🛠 Fixed Assets'!I12:I200),0)`,
        '8670', '← Total CCA from Fixed Assets tab'],
      ['TOTAL EXPENSES (accrual + CCA)', '=C13+C14+C15+C16', '9367', '← For Schedule 125'],
      ['', '', '', ''],
      ['NET INCOME PER BOOKS', '=C8-C17', '9999', '← Revenue minus Expenses'],
      ['', '', '', ''], ['', '', '', ''], ['', '', '', ''], ['', '', '', ''], ['', '', '', ''], ['', '', '', ''],
    ]],

    // Schedule 1 — Book-to-Tax
    [`'${TITLE}'!A28`, [['  SCHEDULE 1 — BOOK-TO-TAX ADJUSTMENTS']]],
    [`'${TITLE}'!B30:E37`, [
      ['Net Income per Books (from above)', '=C19', '', '← Starting point'],
      ['Add back: 50% of Meals & Entertainment (non-deductible)',
        `=ROUND(-SUMIFS('${txnTitle}'!E12:E1000,'${txnTitle}'!F12:F1000,"Meals & Entertainment",'${txnTitle}'!E12:E1000,"<0")*0.5,2)`,
        '101', '← CRA only allows 50% of meals'],
      ['Add back: Amortization per books', 0, '104', '← Tradebooks doesn\'t book amortization separately; usually $0'],
      ['Less: CCA per Schedule 8', '=-C16', '', '← CCA is deducted on tax side instead of book amortization'],
      ['Less: Other tax adjustments', 0, '', '← Manual entry if any'],
      ['', '', '', ''],
      ['TAXABLE INCOME (for T2 line 600)', '=C30+C31+C32+C33+C34', '', '← Net income for tax purposes'],
      ['', '', '', ''],
    ]],

    // Schedule 8 — CCA summary
    [`'${TITLE}'!A39`, [['  SCHEDULE 8 — CCA SUMMARY (per Fixed Assets tab)']]],
    [`'${TITLE}'!B41:E45`, [
      ['Total Opening UCC (start of FY)', `=IFERROR(SUM('🛠 Fixed Assets'!F12:F200),0)`, '', ''],
      ['Total Additions This FY', `=IFERROR(SUM('🛠 Fixed Assets'!G12:G200),0)`, '', ''],
      ['Total CCA Claimed This FY', `=IFERROR(SUM('🛠 Fixed Assets'!I12:I200),0)`, '', '← Goes to Schedule 1 above'],
      ['Total Ending UCC (carries forward)', `=IFERROR(SUM('🛠 Fixed Assets'!J12:J200),0)`, '', ''],
      ['', '', '', ''],
    ]],

    // Schedule 100 — Balance Sheet (rough)
    [`'${TITLE}'!A47`, [['  SCHEDULE 100 — BALANCE SHEET (rough — verify each line)']]],
    [`'${TITLE}'!B49:E55`, [
      ['ASSETS', '', '', ''],
      ['Cash on hand (per latest reconciled bank balances)',
        `=IFERROR(SUMIFS('🏦 Account Balances'!I12:I500,'🏦 Account Balances'!K12:K500,"✓ Balanced"),0)`,
        '1001', '← From Account Balances tab'],
      ['Accounts Receivable (AR adjustments at FYE)',
        `=IFERROR(SUMIF('📓 Adjusting Entries'!C12:C200,"Accounts Receivable (AR)",'📓 Adjusting Entries'!H12:H200),0)`,
        '1062', '← From Adjusting Entries tab'],
      ['Net Fixed Assets (Ending UCC)', '=C44', '1781', '← From Schedule 8 above'],
      ['TOTAL ASSETS', '=C50+C51+C52', '2599', ''],
      ['', '', '', ''], ['', '', '', ''],
    ]],
    [`'${TITLE}'!B57:E62`, [
      ['LIABILITIES', '', '', ''],
      ['Accounts Payable (AP adjustments at FYE)',
        `=IFERROR(SUMIF('📓 Adjusting Entries'!C12:C200,"Accounts Payable (AP)",'📓 Adjusting Entries'!H12:H200),0)`,
        '2620', '← From Adjusting Entries tab'],
      ['Other liabilities (manual)', 0, '2960', '← Loans, owner advances, etc.'],
      ['TOTAL LIABILITIES', '=C58+C59', '3499', ''],
      ['', '', '', ''],
      ['EQUITY (plug — Assets minus Liabilities)', '=C53-C60', '3640', '← Should reconcile with retained earnings + share capital'],
    ]],

    // Tax Calculation
    [`'${TITLE}'!A64`, [['  TAX CALCULATION (rough — MNP to verify with actual T2 software)']]],
    [`'${TITLE}'!B66:E70`, [
      ['Taxable Income (from Schedule 1 above)', '=C36', '', ''],
      ['Combined ON+Fed small business rate', 0.125, '', '← 12.5% est. — verify with advisor'],
      ['Estimated Corporate Tax Owing', '=C66*C67', '', '← Rough — actual rate depends on TOSI, SBD eligibility, etc.'],
      ['', '', '', ''],
      ['HANDOFF NOTE', 'Hand this worksheet to MNP. They review, verify against their own T2 software, and file. Cash-basis books are clean; accrual adjustments are in 📓 Adjusting Entries; CCA is in 🛠 Fixed Assets. All source data is in the books snapshot included with the year-end package.', '', ''],
    ]],
  ];
  for (const [range, values] of writes) {
    const res = await writeRange(env, userId, range, values);
    if (!res.ok) errors.push(`Failed to populate ${range}: ${res.error}`);
  }

  changes.push(`Added '${TITLE}' tab — consolidated T2-prep view that MNP reviews instead of preparing.`);
}

// ── Helpers ──

function colWidthReq(sheetId, startIdx, endIdx, pixelSize) {
  return {
    updateDimensionProperties: {
      range: { sheetId, dimension: 'COLUMNS', startIndex: startIdx, endIndex: endIdx },
      properties: { pixelSize },
      fields: 'pixelSize',
    },
  };
}
