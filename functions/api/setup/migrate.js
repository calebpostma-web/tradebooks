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
