import { authenticateRequest } from '../_shared.js';
import { saveGoogleRefreshToken } from '../_google.js';

// functions/api/google-setup.js
// Handles Google OAuth token exchange and automatic sheet + script creation
// POST /api/google-setup with { code } or { action: 'create-sheet', token }
//
// v2: Full styled template matching Postma_Corporate_Bookkeeping
//     - 7 tabs with colored titles, summary blocks, data tables, tab colors
//     - Frozen headers, number formats, data validations, alt-row colors
//     - Year-End tab pulls category totals via formulas

export async function onRequestPost(context) {
  const { request, env } = context;
  const origin = request.headers.get('Origin') || '';

  const headers = {
    'Content-Type': 'application/json',
    'Access-Control-Allow-Origin': origin,
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization',
  };

  try {
    const body = await request.json();
    const auth = await authenticateRequest(request, env);
    const userId = auth?.userId || null;

    if (body.code) {
      return handleTokenExchange(body.code, origin, env, headers, userId);
    }

    if (body.action === 'create-sheet' && body.accessToken) {
      return handleCreateSheet(body.accessToken, body.profile || {}, env, headers, userId);
    }

    return new Response(JSON.stringify({ ok: false, error: 'Invalid request' }), { headers });

  } catch (err) {
    return new Response(JSON.stringify({ ok: false, error: err.message }), { headers, status: 500 });
  }
}

export async function onRequestOptions() {
  return new Response(null, {
    headers: {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization',
    }
  });
}


// ════════════════════════════════════════════════════════════════════
// STEP 1: Exchange authorization code for access + refresh tokens
// ════════════════════════════════════════════════════════════════════
async function handleTokenExchange(code, origin, env, headers, userId) {
  const tokenResp = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      code,
      client_id: env.GOOGLE_CLIENT_ID,
      client_secret: env.GOOGLE_CLIENT_SECRET,
      redirect_uri: origin + '/app/',
      grant_type: 'authorization_code',
    }),
  });

  const tokenData = await tokenResp.json();

  if (tokenData.error) {
    return new Response(JSON.stringify({
      ok: false,
      error: tokenData.error_description || tokenData.error,
    }), { headers });
  }

  const userResp = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
    headers: { 'Authorization': `Bearer ${tokenData.access_token}` },
  });
  const userInfo = await userResp.json();

  // Persist refresh token so the Worker can act on the user's sheet without re-OAuth
  if (userId && tokenData.refresh_token) {
    await saveGoogleRefreshToken(
      env,
      userId,
      tokenData.refresh_token,
      tokenData.access_token,
      tokenData.expires_in,
    );
  }

  return new Response(JSON.stringify({
    ok: true,
    accessToken: tokenData.access_token,
    expiresIn: tokenData.expires_in,
    email: userInfo.email,
    name: userInfo.name,
  }), { headers });
}


// ════════════════════════════════════════════════════════════════════
// COLOR PALETTE (matches original template)
// ════════════════════════════════════════════════════════════════════
const COLORS = {
  green:        { red: 0.173, green: 0.322, blue: 0.204 },
  greenLight:   { red: 0.878, green: 0.925, blue: 0.890 },
  greenTint:    { red: 0.941, green: 0.965, blue: 0.945 },
  red:          { red: 0.431, green: 0.133, blue: 0.133 },
  redLight:     { red: 0.984, green: 0.918, blue: 0.918 },
  redTint:      { red: 0.996, green: 0.969, blue: 0.969 },
  blue:         { red: 0.090, green: 0.216, blue: 0.388 },
  blueLight:    { red: 0.890, green: 0.925, blue: 0.976 },
  blueTint:     { red: 0.961, green: 0.976, blue: 0.996 },
  teal:         { red: 0.078, green: 0.278, blue: 0.247 },
  tealLight:    { red: 0.859, green: 0.929, blue: 0.918 },
  tealTint:     { red: 0.949, green: 0.980, blue: 0.976 },
  brown:        { red: 0.420, green: 0.255, blue: 0.098 },
  brownLight:   { red: 0.961, green: 0.933, blue: 0.886 },
  brownTint:    { red: 0.988, green: 0.976, blue: 0.957 },
  yellow:       { red: 1.000, green: 0.976, blue: 0.843 },
  white:        { red: 1.000, green: 1.000, blue: 1.000 },
  grey:         { red: 0.950, green: 0.950, blue: 0.950 },
  textDark:     { red: 0.129, green: 0.145, blue: 0.161 },
  textMuted:    { red: 0.450, green: 0.450, blue: 0.450 },
};


// ════════════════════════════════════════════════════════════════════
// STEP 2: Create Google Sheet with full styling
// ════════════════════════════════════════════════════════════════════
async function handleCreateSheet(accessToken, profile, env, headers, userId) {
  const authHeader = { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' };

  // Create the Google Sheet with all 7 tabs
  const sheetResp = await fetch('https://sheets.googleapis.com/v4/spreadsheets', {
    method: 'POST',
    headers: authHeader,
    body: JSON.stringify({
      properties: {
        title: `${profile.businessName || 'My Business'} — TradeBooks`,
        locale: 'en_CA',
        timeZone: 'America/Toronto',
      },
      sheets: [
        { properties: { sheetId: 100, title: '📊 Dashboard', index: 0,
            gridProperties: { rowCount: 60, columnCount: 12, frozenRowCount: 1 },
            tabColor: COLORS.green } },
        { properties: { sheetId: 200, title: '⚙️ Config', index: 1,
            gridProperties: { rowCount: 100, columnCount: 10 },
            tabColor: COLORS.green } },
        // 📒 Transactions — single signed ledger replacing Income + Expenses.
        // Columns B-M: Date | Party | Description | Amount (signed) | Category |
        //              HST flag | HST amount | Account | Source | Ref |
        //              Related Invoice | Match Status
        // Positive Amount = money in; negative = money out. Category "Internal
        // Transfer" is excluded from P&L and HST math client-side.
        { properties: { sheetId: 300, title: '📒 Transactions', index: 2,
            gridProperties: { rowCount: 1000, columnCount: 13, frozenRowCount: 11 },
            tabColor: COLORS.teal } },
        { properties: { sheetId: 500, title: '🧾 Invoices', index: 3,
            gridProperties: { rowCount: 500, columnCount: 17, frozenRowCount: 11 },
            tabColor: COLORS.blue } },
        { properties: { sheetId: 600, title: '📋 HST Returns', index: 4,
            gridProperties: { rowCount: 40, columnCount: 10 },
            tabColor: COLORS.teal } },
        { properties: { sheetId: 700, title: '📅 Year-End', index: 5,
            gridProperties: { rowCount: 80, columnCount: 8 },
            tabColor: COLORS.brown } },
        // 💼 Payroll — pay run history, YTD tracking, remittance due dates.
        // 16 cols B-Q: Pay Date, Employee, Age, Business, Work Description,
        // Hours, Rate, Gross, CPP (ee), EI (ee), Fed Tax, ON Tax, Net Pay,
        // YTD Gross, Remittance Due, Status
        { properties: { sheetId: 800, title: '💼 Payroll', index: 6,
            gridProperties: { rowCount: 500, columnCount: 17, frozenRowCount: 11 },
            tabColor: COLORS.brown } },
        // 📝 Work Log — contemporaneous entries for CRA audit defence.
        // 8 cols B-I: Date, Employee, Business, Task Description, Hours,
        // Rate, Notes, Entry Audit (server-timestamp + "corrected" flag)
        { properties: { sheetId: 900, title: '📝 Work Log', index: 7,
            gridProperties: { rowCount: 1000, columnCount: 9, frozenRowCount: 11 },
            tabColor: COLORS.brown } },
      ],
    }),
  });

  const sheetData = await sheetResp.json();
  if (sheetData.error) {
    return new Response(JSON.stringify({ ok: false, error: 'Sheet creation failed: ' + (sheetData.error.message || JSON.stringify(sheetData.error)) }), { headers });
  }

  const spreadsheetId = sheetData.spreadsheetId;
  const spreadsheetUrl = sheetData.spreadsheetUrl;

  // Apply styling and populate values — capture errors so the frontend can surface them
  const setupErrors = [];
  try {
    const stylingResult = await applyStyling(accessToken, spreadsheetId);
    if (stylingResult && !stylingResult.ok) {
      setupErrors.push(`Styling failed: ${stylingResult.error}`);
    }
  } catch (e) {
    setupErrors.push(`Styling threw: ${e.message}`);
  }
  try {
    const valuesResult = await populateValues(accessToken, spreadsheetId, profile);
    if (valuesResult && !valuesResult.ok) {
      setupErrors.push(`Population failed: ${valuesResult.error}`);
    }
  } catch (e) {
    setupErrors.push(`Population threw: ${e.message}`);
  }

  // Persist sheet_id to the user's profile
  if (userId) {
    try {
      await env.DB.prepare('UPDATE profiles SET sheet_id = ? WHERE user_id = ?')
        .bind(spreadsheetId, userId)
        .run();
    } catch (e) { /* non-blocking */ }
  }

  return new Response(JSON.stringify({
    ok: true,
    spreadsheetId,
    spreadsheetUrl,
    setupErrors: setupErrors.length ? setupErrors : null,
  }), { headers });
}


// ════════════════════════════════════════════════════════════════════
// STYLING HELPERS
// ════════════════════════════════════════════════════════════════════

function cellFormat(sheetId, r1, c1, r2, c2, format) {
  const fields = [];
  if ('backgroundColor' in format) fields.push('userEnteredFormat.backgroundColor');
  if ('horizontalAlignment' in format) fields.push('userEnteredFormat.horizontalAlignment');
  if ('verticalAlignment' in format) fields.push('userEnteredFormat.verticalAlignment');
  if ('textFormat' in format) fields.push('userEnteredFormat.textFormat');
  if ('numberFormat' in format) fields.push('userEnteredFormat.numberFormat');
  if ('wrapStrategy' in format) fields.push('userEnteredFormat.wrapStrategy');
  if ('padding' in format) fields.push('userEnteredFormat.padding');
  if ('borders' in format) fields.push('userEnteredFormat.borders');
  return {
    repeatCell: {
      range: { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 },
      cell: { userEnteredFormat: format },
      fields: fields.join(','),
    }
  };
}

function bannerRequest(sheetId, row, startCol, endCol, bgColor) {
  return [
    { mergeCells: {
        range: { sheetId, startRowIndex: row, endRowIndex: row + 1, startColumnIndex: startCol, endColumnIndex: endCol },
        mergeType: 'MERGE_ALL',
    }},
    { repeatCell: {
        range: { sheetId, startRowIndex: row, endRowIndex: row + 1, startColumnIndex: startCol, endColumnIndex: endCol },
        cell: { userEnteredFormat: {
            backgroundColor: bgColor,
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
            textFormat: { foregroundColor: COLORS.white, bold: true, fontSize: 12 },
        }},
        fields: 'userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat)',
    }},
    { updateDimensionProperties: {
        range: { sheetId, dimension: 'ROWS', startIndex: row, endIndex: row + 1 },
        properties: { pixelSize: 32 }, fields: 'pixelSize',
    }},
  ];
}

function sectionRequest(sheetId, row, startCol, endCol, bgColor) {
  return [
    { mergeCells: {
        range: { sheetId, startRowIndex: row, endRowIndex: row + 1, startColumnIndex: startCol, endColumnIndex: endCol },
        mergeType: 'MERGE_ALL',
    }},
    { repeatCell: {
        range: { sheetId, startRowIndex: row, endRowIndex: row + 1, startColumnIndex: startCol, endColumnIndex: endCol },
        cell: { userEnteredFormat: {
            backgroundColor: bgColor,
            horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
            textFormat: { foregroundColor: COLORS.white, bold: true, fontSize: 10 },
            padding: { left: 8 },
        }},
        fields: 'userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat,padding)',
    }},
    { updateDimensionProperties: {
        range: { sheetId, dimension: 'ROWS', startIndex: row, endIndex: row + 1 },
        properties: { pixelSize: 26 }, fields: 'pixelSize',
    }},
  ];
}

function headerRowRequest(sheetId, row, startCol, endCol, bgColor) {
  return {
    repeatCell: {
      range: { sheetId, startRowIndex: row, endRowIndex: row + 1, startColumnIndex: startCol, endColumnIndex: endCol },
      cell: { userEnteredFormat: {
          backgroundColor: bgColor,
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          textFormat: { foregroundColor: COLORS.white, bold: true, fontSize: 10 },
          wrapStrategy: 'WRAP',
      }},
      fields: 'userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat,wrapStrategy)',
    }
  };
}

function bandingRequest(sheetId, startRow, endRow, startCol, endCol, tintColor) {
  return {
    addBanding: {
      bandedRange: {
        range: { sheetId, startRowIndex: startRow, endRowIndex: endRow, startColumnIndex: startCol, endColumnIndex: endCol },
        rowProperties: { firstBandColor: COLORS.white, secondBandColor: tintColor },
      }
    }
  };
}

const FMT_CURRENCY = { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00;("$"#,##0.00)' } };
const FMT_DATE     = { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } };
const FMT_PERCENT  = { numberFormat: { type: 'PERCENT', pattern: '0.0%' } };
const FMT_EDITABLE = { backgroundColor: COLORS.yellow, textFormat: { foregroundColor: { red: 0.15, green: 0.35, blue: 0.65 } } };

function colWidth(sheetId, startIdx, endIdx, px) {
  return {
    updateDimensionProperties: {
      range: { sheetId, dimension: 'COLUMNS', startIndex: startIdx, endIndex: endIdx },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }
  };
}


// ════════════════════════════════════════════════════════════════════
// APPLY STYLING — batchUpdate for all visual formatting
// ════════════════════════════════════════════════════════════════════
async function applyStyling(accessToken, spreadsheetId) {
  const authHeader = { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' };
  const DASH = 100, CFG = 200, TXN = 300, INV = 500, HST = 600, YE = 700, PAY = 800, WLOG = 900;
  const requests = [];

  // ─── DASHBOARD ───
  requests.push(...bannerRequest(DASH, 0, 0, 7, COLORS.green));
  requests.push(...bannerRequest(DASH, 2, 0, 7, COLORS.green));
  requests.push({
    repeatCell: {
      range: { sheetId: DASH, startRowIndex: 4, endRowIndex: 10, startColumnIndex: 0, endColumnIndex: 7 },
      cell: { userEnteredFormat: { backgroundColor: COLORS.greenTint } },
      fields: 'userEnteredFormat.backgroundColor',
    }
  });
  requests.push(cellFormat(DASH, 3, 0, 10, 5, FMT_CURRENCY));
  requests.push(cellFormat(DASH, 4, 1, 10, 2, { textFormat: { bold: true, fontSize: 10 } }));
  requests.push(cellFormat(DASH, 8, 0, 9, 7, { backgroundColor: COLORS.greenLight, textFormat: { bold: true } }));
  requests.push(cellFormat(DASH, 9, 0, 10, 7, { backgroundColor: COLORS.tealLight, textFormat: { bold: true } }));
  requests.push(colWidth(DASH, 0, 1, 30));
  requests.push(colWidth(DASH, 1, 2, 220));
  requests.push(colWidth(DASH, 2, 7, 130));

  // ─── CONFIG ───
  requests.push(...bannerRequest(CFG, 0, 0, 6, COLORS.green));
  requests.push(cellFormat(CFG, 1, 0, 2, 6, { horizontalAlignment: 'CENTER', textFormat: { italic: true, foregroundColor: COLORS.textMuted, fontSize: 9 } }));
  requests.push(...sectionRequest(CFG, 3, 0, 6, COLORS.green));
  requests.push(cellFormat(CFG, 4, 1, 13, 2, { textFormat: { bold: true, fontSize: 10 } }));
  requests.push(cellFormat(CFG, 4, 2, 13, 3, FMT_EDITABLE));
  requests.push(cellFormat(CFG, 4, 4, 13, 5, { textFormat: { bold: true, fontSize: 10 } }));
  requests.push(cellFormat(CFG, 4, 5, 13, 6, FMT_EDITABLE));
  requests.push(...sectionRequest(CFG, 14, 0, 6, COLORS.green));
  requests.push(cellFormat(CFG, 15, 1, 16, 2, { textFormat: { italic: true, foregroundColor: COLORS.textMuted, fontSize: 9 } }));
  requests.push(cellFormat(CFG, 15, 4, 16, 5, { textFormat: { italic: true, foregroundColor: COLORS.textMuted, fontSize: 9 } }));
  requests.push(cellFormat(CFG, 16, 4, 36, 5, FMT_EDITABLE));
  requests.push(...sectionRequest(CFG, 49, 0, 6, COLORS.green));
  requests.push(cellFormat(CFG, 50, 1, 51, 2, { textFormat: { italic: true, foregroundColor: COLORS.textMuted, fontSize: 9 } }));
  requests.push(cellFormat(CFG, 50, 4, 51, 5, { textFormat: { italic: true, foregroundColor: COLORS.textMuted, fontSize: 9 } }));
  requests.push(cellFormat(CFG, 51, 4, 60, 5, FMT_EDITABLE));
  requests.push(...sectionRequest(CFG, 61, 0, 6, COLORS.green));
  requests.push(cellFormat(CFG, 62, 1, 63, 2, { textFormat: { bold: true, foregroundColor: COLORS.textMuted, fontSize: 9 } }));
  requests.push(cellFormat(CFG, 62, 4, 63, 5, { textFormat: { bold: true, foregroundColor: COLORS.textMuted, fontSize: 9 } }));
  requests.push(cellFormat(CFG, 63, 1, 74, 2, FMT_EDITABLE));
  requests.push(cellFormat(CFG, 63, 4, 74, 5, FMT_EDITABLE));
  // Employees section — B-I (8 data columns). Bold header row; yellow editable body.
  requests.push(...sectionRequest(CFG, 74, 0, 9, COLORS.green));
  requests.push(cellFormat(CFG, 75, 1, 76, 9, { textFormat: { bold: true, foregroundColor: COLORS.textMuted, fontSize: 9 } }));
  requests.push(cellFormat(CFG, 76, 1, 86, 9, FMT_EDITABLE));
  // DOB (col C idx 2) and Start Date (col F idx 5) get date formatting.
  requests.push(cellFormat(CFG, 76, 2, 86, 3, { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } }));
  requests.push(cellFormat(CFG, 76, 5, 86, 6, { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } }));
  // Default Rate (col G idx 6) as currency per hour.
  requests.push(cellFormat(CFG, 76, 6, 86, 7, { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } }));
  requests.push(colWidth(CFG, 0, 1, 30));
  requests.push(colWidth(CFG, 1, 2, 180));
  requests.push(colWidth(CFG, 2, 3, 260));
  requests.push(colWidth(CFG, 3, 4, 30));
  requests.push(colWidth(CFG, 4, 5, 180));
  requests.push(colWidth(CFG, 5, 6, 260));

  // ─── TRANSACTIONS ───
  // Single signed ledger: + for money in, − for money out. Replaces Income + Expenses.
  // Columns (0-indexed): A=gutter, B=Date, C=Party, D=Description, E=Amount,
  // F=Category, G=HST Flag, H=HST Amount, I=Account, J=Source, K=Ref,
  // L=Related Invoice, M=Match Status
  requests.push(...bannerRequest(TXN, 0, 0, 13, COLORS.teal));
  requests.push(...sectionRequest(TXN, 1, 1, 6, COLORS.teal));
  requests.push(cellFormat(TXN, 2, 1, 10, 2, { textFormat: { bold: true, fontSize: 10 } }));
  requests.push(cellFormat(TXN, 2, 2, 10, 6, { backgroundColor: COLORS.tealTint, numberFormat: FMT_CURRENCY.numberFormat }));
  requests.push(headerRowRequest(TXN, 10, 1, 13, COLORS.teal));
  requests.push({ updateDimensionProperties: { range: { sheetId: TXN, dimension: 'ROWS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  requests.push(bandingRequest(TXN, 11, 1000, 1, 13, COLORS.tealTint));
  // Date col B
  requests.push(cellFormat(TXN, 11, 1, 1000, 2, FMT_DATE));
  // Amount col E — signed currency (negatives shown in parens via FMT_CURRENCY)
  requests.push(cellFormat(TXN, 11, 4, 1000, 5, FMT_CURRENCY));
  // HST Amount col H
  requests.push(cellFormat(TXN, 11, 7, 1000, 8, FMT_CURRENCY));
  // HST Flag validation — col G (index 6)
  requests.push({
    setDataValidation: {
      range: { sheetId: TXN, startRowIndex: 11, endRowIndex: 1000, startColumnIndex: 6, endColumnIndex: 7 },
      rule: { condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: 'Yes' }, { userEnteredValue: 'No' }] }, showCustomUi: true },
    }
  });
  // Match Status validation — col M (index 12)
  requests.push({
    setDataValidation: {
      range: { sheetId: TXN, startRowIndex: 11, endRowIndex: 1000, startColumnIndex: 12, endColumnIndex: 13 },
      rule: { condition: { type: 'ONE_OF_LIST', values: [
        { userEnteredValue: 'Matched' },
        { userEnteredValue: 'Unmatched' },
        { userEnteredValue: 'N/A' },
      ]}, showCustomUi: true },
    }
  });
  requests.push(colWidth(TXN, 0, 1, 30));   // A gutter
  requests.push(colWidth(TXN, 1, 2, 100));  // B Date
  requests.push(colWidth(TXN, 2, 3, 180));  // C Party
  requests.push(colWidth(TXN, 3, 4, 220));  // D Description
  requests.push(colWidth(TXN, 4, 5, 110));  // E Amount (signed)
  requests.push(colWidth(TXN, 5, 6, 170));  // F Category
  requests.push(colWidth(TXN, 6, 7, 70));   // G HST Flag
  requests.push(colWidth(TXN, 7, 8, 110));  // H HST Amount
  requests.push(colWidth(TXN, 8, 9, 90));   // I Account
  requests.push(colWidth(TXN, 9, 10, 90));  // J Source
  requests.push(colWidth(TXN, 10, 11, 220));// K Ref
  requests.push(colWidth(TXN, 11, 12, 100));// L Related Invoice
  requests.push(colWidth(TXN, 12, 13, 100));// M Match Status

  // ─── INVOICES ───
  // 14 cols: A gutter, B Invoice #, C Date, D Client, E Description, F Amount excl HST,
  // G HST, H Total, I HST flag, J Due, K Status, L Date Paid, M Notes, N Revenue Category,
  // O Deposit Amount, P Deposit Date Received, Q Balance Due
  requests.push(...bannerRequest(INV, 0, 0, 17, COLORS.blue));
  requests.push(...sectionRequest(INV, 1, 1, 6, COLORS.blue));
  requests.push(cellFormat(INV, 2, 1, 10, 2, { textFormat: { bold: true, fontSize: 10 } }));
  requests.push(cellFormat(INV, 2, 2, 10, 6, { backgroundColor: COLORS.blueTint }));
  requests.push(headerRowRequest(INV, 10, 1, 17, COLORS.blue));
  requests.push({ updateDimensionProperties: { range: { sheetId: INV, dimension: 'ROWS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  requests.push(bandingRequest(INV, 11, 500, 1, 17, COLORS.blueTint));
  requests.push(cellFormat(INV, 11, 2, 500, 3, FMT_DATE));
  requests.push(cellFormat(INV, 11, 5, 500, 6, FMT_CURRENCY));
  requests.push(cellFormat(INV, 11, 6, 500, 7, FMT_CURRENCY));
  requests.push(cellFormat(INV, 11, 7, 500, 8, FMT_CURRENCY));
  requests.push(cellFormat(INV, 11, 9, 500, 10, FMT_DATE));
  requests.push(cellFormat(INV, 11, 11, 500, 12, FMT_DATE));
  requests.push(cellFormat(INV, 11, 14, 500, 15, FMT_CURRENCY));  // O Deposit Amount
  requests.push(cellFormat(INV, 11, 15, 500, 16, FMT_DATE));      // P Deposit Date Received
  requests.push(cellFormat(INV, 11, 16, 500, 17, FMT_CURRENCY));  // Q Balance Due
  requests.push({
    setDataValidation: {
      range: { sheetId: INV, startRowIndex: 11, endRowIndex: 500, startColumnIndex: 8, endColumnIndex: 9 },
      rule: { condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: 'Yes' }, { userEnteredValue: 'No' }] }, showCustomUi: true },
    }
  });
  requests.push({
    setDataValidation: {
      range: { sheetId: INV, startRowIndex: 11, endRowIndex: 500, startColumnIndex: 10, endColumnIndex: 11 },
      rule: { condition: { type: 'ONE_OF_LIST', values: [
        { userEnteredValue: 'Unpaid' },
        { userEnteredValue: 'Awaiting Deposit' },
        { userEnteredValue: 'Deposit Received' },
        { userEnteredValue: 'Paid' },
        { userEnteredValue: 'Overdue' },
        { userEnteredValue: 'Cancelled' },
      ] }, showCustomUi: true },
    }
  });
  requests.push(colWidth(INV, 0, 1, 30));
  requests.push(colWidth(INV, 1, 2, 90));
  requests.push(colWidth(INV, 2, 3, 110));
  requests.push(colWidth(INV, 3, 4, 180));
  requests.push(colWidth(INV, 4, 5, 260));
  requests.push(colWidth(INV, 5, 6, 110));
  requests.push(colWidth(INV, 6, 7, 90));
  requests.push(colWidth(INV, 7, 8, 110));
  requests.push(colWidth(INV, 8, 9, 70));
  requests.push(colWidth(INV, 9, 10, 130));  // K Status — wider for new labels
  requests.push(colWidth(INV, 10, 11, 90));
  requests.push(colWidth(INV, 11, 12, 110));
  requests.push(colWidth(INV, 12, 13, 200));
  requests.push(colWidth(INV, 13, 14, 150));  // N Revenue Category
  requests.push(colWidth(INV, 14, 15, 110));  // O Deposit Amount
  requests.push(colWidth(INV, 15, 16, 130));  // P Deposit Date Received
  requests.push(colWidth(INV, 16, 17, 110));  // Q Balance Due

  // ─── HST RETURNS ───
  requests.push(...bannerRequest(HST, 0, 0, 7, COLORS.teal));
  requests.push(cellFormat(HST, 1, 0, 2, 7, { horizontalAlignment: 'CENTER', textFormat: { italic: true, foregroundColor: COLORS.textMuted, fontSize: 9 } }));
  // Row 3 (index 2): Fiscal Year Start param row. B3 label right-aligned bold; C3 editable-yellow with date format.
  requests.push(cellFormat(HST, 2, 1, 3, 2, {
    horizontalAlignment: 'RIGHT',
    textFormat: { bold: true, fontSize: 10, foregroundColor: COLORS.textDark },
  }));
  requests.push(cellFormat(HST, 2, 2, 3, 3, {
    ...FMT_EDITABLE,
    numberFormat: FMT_DATE.numberFormat,
    horizontalAlignment: 'LEFT',
  }));
  requests.push(headerRowRequest(HST, 3, 1, 7, COLORS.teal));
  requests.push({ updateDimensionProperties: { range: { sheetId: HST, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });
  requests.push(cellFormat(HST, 4, 1, 8, 2, { textFormat: { bold: true, fontSize: 10 } }));
  requests.push(cellFormat(HST, 4, 2, 8, 7, { backgroundColor: COLORS.tealTint, numberFormat: FMT_CURRENCY.numberFormat }));
  requests.push(cellFormat(HST, 7, 1, 8, 7, { backgroundColor: COLORS.tealLight, textFormat: { bold: true } }));
  requests.push(...sectionRequest(HST, 10, 0, 7, COLORS.teal));
  requests.push(cellFormat(HST, 11, 1, 18, 7, { textFormat: { fontSize: 10 }, wrapStrategy: 'WRAP' }));
  requests.push(colWidth(HST, 0, 1, 30));
  requests.push(colWidth(HST, 1, 2, 260));
  requests.push(colWidth(HST, 2, 7, 120));

  // ─── YEAR-END ───
  requests.push(...bannerRequest(YE, 0, 0, 6, COLORS.brown));
  requests.push(cellFormat(YE, 1, 0, 2, 6, { horizontalAlignment: 'LEFT', textFormat: { italic: true, foregroundColor: COLORS.textMuted, fontSize: 9 }, padding: { left: 8 } }));
  requests.push(...sectionRequest(YE, 3, 0, 6, COLORS.brown));
  requests.push(cellFormat(YE, 4, 1, 10, 2, { textFormat: { fontSize: 10 } }));
  requests.push(cellFormat(YE, 4, 2, 10, 3, FMT_CURRENCY));
  requests.push(cellFormat(YE, 4, 4, 10, 6, { textFormat: { italic: true, foregroundColor: COLORS.textMuted, fontSize: 9 } }));
  requests.push(cellFormat(YE, 9, 1, 10, 6, { backgroundColor: COLORS.brownLight, textFormat: { bold: true } }));
  requests.push(...sectionRequest(YE, 11, 0, 6, COLORS.brown));
  requests.push(cellFormat(YE, 12, 1, 37, 2, { textFormat: { fontSize: 10 } }));
  requests.push(cellFormat(YE, 12, 2, 37, 3, FMT_CURRENCY));
  requests.push(cellFormat(YE, 12, 4, 37, 6, { textFormat: { italic: true, foregroundColor: COLORS.textMuted, fontSize: 9 } }));
  requests.push(cellFormat(YE, 36, 1, 37, 6, { backgroundColor: COLORS.brownLight, textFormat: { bold: true } }));
  requests.push(...sectionRequest(YE, 38, 0, 6, COLORS.brown));
  requests.push(cellFormat(YE, 39, 1, 44, 2, { textFormat: { fontSize: 10 } }));
  requests.push(cellFormat(YE, 39, 2, 44, 3, FMT_CURRENCY));
  requests.push(cellFormat(YE, 40, 2, 41, 3, FMT_PERCENT));
  requests.push(cellFormat(YE, 42, 1, 43, 6, { backgroundColor: COLORS.brownLight, textFormat: { bold: true } }));
  requests.push(cellFormat(YE, 39, 4, 44, 6, { textFormat: { italic: true, foregroundColor: COLORS.textMuted, fontSize: 9 } }));
  requests.push(...sectionRequest(YE, 45, 0, 6, COLORS.brown));
  requests.push(cellFormat(YE, 46, 1, 50, 2, { textFormat: { fontSize: 10 } }));
  requests.push(cellFormat(YE, 46, 2, 50, 3, FMT_CURRENCY));
  requests.push(cellFormat(YE, 46, 4, 50, 6, { textFormat: { italic: true, foregroundColor: COLORS.textMuted, fontSize: 9 } }));
  requests.push(cellFormat(YE, 48, 1, 49, 6, { backgroundColor: COLORS.brownLight, textFormat: { bold: true } }));
  requests.push(...sectionRequest(YE, 50, 0, 6, COLORS.brown));
  requests.push(cellFormat(YE, 51, 1, 60, 2, { textFormat: { fontSize: 10 } }));
  requests.push(cellFormat(YE, 51, 2, 60, 5, { horizontalAlignment: 'CENTER', textFormat: { fontSize: 10 } }));
  requests.push(cellFormat(YE, 51, 5, 60, 6, { textFormat: { italic: true, foregroundColor: COLORS.textMuted, fontSize: 9 } }));
  requests.push(colWidth(YE, 0, 1, 30));
  requests.push(colWidth(YE, 1, 2, 260));
  requests.push(colWidth(YE, 2, 3, 140));
  requests.push(colWidth(YE, 3, 4, 30));
  requests.push(colWidth(YE, 4, 5, 30));
  requests.push(colWidth(YE, 5, 6, 260));

  // ─── PAYROLL ───
  // 17 cols (A gutter + B-Q data): Pay Date, Employee, Age, Business,
  // Work Description, Hours, Rate, Gross, CPP (ee), EI (ee), Fed Tax,
  // ON Tax, Net Pay, YTD Gross, Remittance Due, Status
  requests.push(...bannerRequest(PAY, 0, 0, 17, COLORS.brown));
  requests.push(...sectionRequest(PAY, 1, 1, 7, COLORS.brown));
  requests.push(cellFormat(PAY, 2, 1, 10, 2, { textFormat: { bold: true, fontSize: 10 } }));
  requests.push(cellFormat(PAY, 2, 2, 10, 7, { backgroundColor: COLORS.brownTint, numberFormat: FMT_CURRENCY.numberFormat }));
  requests.push(headerRowRequest(PAY, 10, 1, 17, COLORS.brown));
  requests.push({ updateDimensionProperties: { range: { sheetId: PAY, dimension: 'ROWS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });
  requests.push(bandingRequest(PAY, 11, 500, 1, 17, COLORS.brownTint));
  // Dates: col B (Pay Date), col P (Remittance Due)
  requests.push(cellFormat(PAY, 11, 1, 500, 2, FMT_DATE));
  requests.push(cellFormat(PAY, 11, 15, 500, 16, FMT_DATE));
  // Currency: col H (Gross) through col N (YTD Gross) — 7 currency columns
  requests.push(cellFormat(PAY, 11, 7, 500, 14, FMT_CURRENCY));
  // Status dropdown (col Q, index 16)
  requests.push({
    setDataValidation: {
      range: { sheetId: PAY, startRowIndex: 11, endRowIndex: 500, startColumnIndex: 16, endColumnIndex: 17 },
      rule: { condition: { type: 'ONE_OF_LIST', values: [
        { userEnteredValue: 'Pending' }, { userEnteredValue: 'Paid' },
        { userEnteredValue: 'Remitted' }, { userEnteredValue: 'Cancelled' },
      ]}, showCustomUi: true },
    }
  });
  requests.push(colWidth(PAY, 0, 1, 30));    // A gutter
  requests.push(colWidth(PAY, 1, 2, 95));    // B Pay Date
  requests.push(colWidth(PAY, 2, 3, 130));   // C Employee
  requests.push(colWidth(PAY, 3, 4, 50));    // D Age
  requests.push(colWidth(PAY, 4, 5, 100));   // E Business
  requests.push(colWidth(PAY, 5, 6, 200));   // F Work Description
  requests.push(colWidth(PAY, 6, 7, 60));    // G Hours
  requests.push(colWidth(PAY, 7, 8, 70));    // H Rate
  requests.push(colWidth(PAY, 8, 9, 90));    // I Gross
  requests.push(colWidth(PAY, 9, 10, 80));   // J CPP
  requests.push(colWidth(PAY, 10, 11, 70));  // K EI
  requests.push(colWidth(PAY, 11, 12, 80));  // L Fed Tax
  requests.push(colWidth(PAY, 12, 13, 80));  // M ON Tax
  requests.push(colWidth(PAY, 13, 14, 90));  // N Net Pay
  requests.push(colWidth(PAY, 14, 15, 95));  // O YTD Gross
  requests.push(colWidth(PAY, 15, 16, 110)); // P Remittance Due
  requests.push(colWidth(PAY, 16, 17, 90));  // Q Status

  // ─── WORK LOG ───
  // 9 cols (A gutter + B-I data): Date, Employee, Business, Task Description,
  // Hours, Rate, Notes, Entry Audit. Entry Audit is a stringified server
  // timestamp + edit flag — locked from silent edits for CRA audit defence.
  requests.push(...bannerRequest(WLOG, 0, 0, 9, COLORS.brown));
  requests.push(...sectionRequest(WLOG, 1, 1, 8, COLORS.brown));
  requests.push(cellFormat(WLOG, 2, 1, 10, 2, { textFormat: { bold: true, fontSize: 10 } }));
  requests.push(cellFormat(WLOG, 2, 2, 10, 8, { backgroundColor: COLORS.brownTint }));
  requests.push(headerRowRequest(WLOG, 10, 1, 9, COLORS.brown));
  requests.push({ updateDimensionProperties: { range: { sheetId: WLOG, dimension: 'ROWS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  requests.push(bandingRequest(WLOG, 11, 1000, 1, 9, COLORS.brownTint));
  requests.push(cellFormat(WLOG, 11, 1, 1000, 2, FMT_DATE));   // B Date
  requests.push(cellFormat(WLOG, 11, 5, 1000, 6, { numberFormat: { type: 'NUMBER', pattern: '0.00' } })); // F Hours
  requests.push(cellFormat(WLOG, 11, 6, 1000, 7, FMT_CURRENCY)); // G Rate
  // Entry Audit column (col I, idx 8) — read-only mono font, muted
  requests.push(cellFormat(WLOG, 11, 8, 1000, 9, {
    textFormat: { fontSize: 8, foregroundColor: COLORS.textMuted },
    backgroundColor: COLORS.grey,
  }));
  requests.push(colWidth(WLOG, 0, 1, 30));    // A gutter
  requests.push(colWidth(WLOG, 1, 2, 95));    // B Date
  requests.push(colWidth(WLOG, 2, 3, 130));   // C Employee
  requests.push(colWidth(WLOG, 3, 4, 100));   // D Business
  requests.push(colWidth(WLOG, 4, 5, 260));   // E Task Description
  requests.push(colWidth(WLOG, 5, 6, 60));    // F Hours
  requests.push(colWidth(WLOG, 6, 7, 70));    // G Rate
  requests.push(colWidth(WLOG, 7, 8, 200));   // H Notes
  requests.push(colWidth(WLOG, 8, 9, 180));   // I Entry Audit

  const stylingResp = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`, {
    method: 'POST', headers: authHeader,
    body: JSON.stringify({ requests }),
  });
  const stylingData = await stylingResp.json();
  if (stylingData.error) {
    return { ok: false, step: 'applyStyling', error: stylingData.error.message || JSON.stringify(stylingData.error) };
  }
  return { ok: true };
}


// ════════════════════════════════════════════════════════════════════
// POPULATE VALUES — text, formulas, headers, config, Year-End
// ════════════════════════════════════════════════════════════════════
async function populateValues(accessToken, spreadsheetId, profile) {
  const authHeader = { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' };

  const prov = profile.province || 'ON';
  const taxRate = {
    ON:0.13, BC:0.12, AB:0.05, SK:0.11, MB:0.12,
    QC:0.14975, NB:0.15, NS:0.15, PE:0.15, NL:0.15
  }[prov] || 0.13;
  const taxPct = Math.round(taxRate * 100);
  const bizName = profile.businessName || 'My Business';
  const fye = profile.fiscalYearEnd || 'December 31';

  const incomeCats = [
    'Consulting Revenue', 'Service Revenue', 'Sales Revenue',
    'Sales Revenue — Materials', 'Rental Income', 'Other Income'
  ];
  const expenseCats = [
    ['Wages & Salaries', 'T2 Line 9060'],
    ['Professional Fees', 'T2 Line 8860'],
    ['Supplies', ''],
    ['Small Tools', ''],
    ['Uniforms', ''],
    ['Fuel', ''],
    ['Telephone', ''],
    ['Dues & Memberships', ''],
    ['Repairs & Maintenance', ''],
    ['Vehicle Repairs', ''],
    ['Utilities', ''],
    ['Interest & Bank Charges', 'no HST component'],
    ['Insurance', 'no HST component'],
    ['Office Supplies', ''],
    ['Advertising & Marketing', ''],
    ['Subcontractors', ''],
    ['Training & Education', ''],
    ['Rent', ''],
    ['Home Office', ''],
    ['Travel', ''],
    ['Permits & Licenses', ''],
    ['Property Tax', ''],
    ['Meals & Entertainment (50%)', 'only 50% deductible'],
    ['Other', ''],
  ];

  const data = [];

  // DASHBOARD
  // Reads from single signed 📒 Transactions tab.
  // E = Amount (signed: + income, − expense). F = Category. H = HST Amount.
  // Internal Transfer rows excluded from P&L and HST via SUMIFS on category.
  data.push({ range: "'📊 Dashboard'!A1", values: [[`${bizName.toUpperCase()}  ·  FISCAL YE ${fye.toUpperCase()}  ·  ${prov} HST ${taxPct}%`]] });
  data.push({ range: "'📊 Dashboard'!A3", values: [['PROFIT & LOSS SUMMARY  (Cash Basis)']] });
  data.push({ range: "'📊 Dashboard'!B5:B10", values: [
    ['Total Revenue (excl HST)'], ['HST Collected'],
    ['Total Expenses (excl HST)'], ['HST Paid — ITCs'],
    ['NET INCOME before tax'], ['HST OWING / (Refund)']
  ]});
  data.push({ range: "'📊 Dashboard'!D5:D10", values: [
    // Revenue: sum of positive amounts, excluding Internal Transfer
    ["=SUMIFS('📒 Transactions'!E12:E1000,'📒 Transactions'!E12:E1000,\">0\",'📒 Transactions'!F12:F1000,\"<>Internal Transfer\")"],
    // HST collected: HST amount on positive (income) rows
    ["=SUMIFS('📒 Transactions'!H12:H1000,'📒 Transactions'!E12:E1000,\">0\",'📒 Transactions'!F12:F1000,\"<>Internal Transfer\")"],
    // Expenses: absolute value of negative amounts, excluding Internal Transfer
    ["=-SUMIFS('📒 Transactions'!E12:E1000,'📒 Transactions'!E12:E1000,\"<0\",'📒 Transactions'!F12:F1000,\"<>Internal Transfer\")"],
    // ITCs: HST amount on negative (expense) rows
    ["=SUMIFS('📒 Transactions'!H12:H1000,'📒 Transactions'!E12:E1000,\"<0\",'📒 Transactions'!F12:F1000,\"<>Internal Transfer\")"],
    ['=D5-D7'],
    ['=D6-D8']
  ]});

  // CONFIG
  data.push({ range: "'⚙️ Config'!A1", values: [['⚙️ TRADEBOOKS CONFIGURATION']] });
  data.push({ range: "'⚙️ Config'!A2", values: [['Edit the yellow cells below. The app reads these settings automatically.']] });
  data.push({ range: "'⚙️ Config'!A4", values: [['  BUSINESS INFORMATION']] });
  data.push({ range: "'⚙️ Config'!B5:C13", values: [
    ['Business Name',      profile.businessName || ''],
    ['Trading / DBA Name', profile.tradingName || ''],
    ['Owner Name',         profile.ownerName || ''],
    ['City',               profile.city || ''],
    ['Province',           prov],
    ['HST / GST Rate',     taxRate],
    ['HST / GST Number',   profile.hstNumber || ''],
    ['Business Type',      profile.businessType || 'sole_prop'],
    ['Fiscal Year End',    fye],
  ]});
  data.push({ range: "'⚙️ Config'!E5:F13", values: [
    ['Primary Bank',        profile.primaryBank || 'BMO'],
    ['Credit Card',         profile.creditCard || 'AMEX'],
    ['Starting Invoice #',  profile.invoiceStart || 1001],
    ['Home Office %',       profile.homeOfficePercent || 0],
    ['Email',               profile.email || ''],
    ['Apps Script URL',     ''],
    ['Corporate Structure', profile.structure || ''],
    ['Business Activities', profile.activities || ''],
    ['', ''],
  ]});
  data.push({ range: "'⚙️ Config'!A15", values: [['  CUSTOM EXPENSE CATEGORIES — add your own below (one per row)']] });
  data.push({ range: "'⚙️ Config'!B16", values: [['Default Categories (do not edit)']] });
  data.push({ range: "'⚙️ Config'!E16", values: [['Your Custom Categories (edit these)']] });
  data.push({ range: "'⚙️ Config'!B17:B48", values: [
    ['Meals & Entertainment'], ['Professional Fees'], ['Wages & Salaries'], ['Small Tools'],
    ['Supplies'], ['Uniforms'], ['Dues & Memberships'], ['Vehicle Repairs'], ['Fuel'],
    ['Insurance'], ['Interest & Bank Charges'], ['Repairs & Maintenance'], ['Property Tax'],
    ['Telephone'], ['Utilities'], ['Conferences'], ['Equipment Purchase'],
    ['Inventory — Materials (COGS)'], ['Advertising & Marketing'], ['Subcontractors'],
    ['Office Supplies'], ['Rent'], ['Home Office'], ['Vehicle Lease/Payments'], ['Travel'],
    ['Training & Education'], ['Permits & Licenses'], ['Tax Payments'], ['Other'],
    ['Internal Transfer'], ['Owner Draw / Distribution'], ['SKIP — not a business expense'],
  ]});
  data.push({ range: "'⚙️ Config'!A50", values: [['  CUSTOM INCOME CATEGORIES — add your own below (one per row)']] });
  data.push({ range: "'⚙️ Config'!B51", values: [['Default Categories (do not edit)']] });
  data.push({ range: "'⚙️ Config'!E51", values: [['Your Custom Categories (edit these)']] });
  data.push({ range: "'⚙️ Config'!B52:B59", values: [
    ['Consulting Revenue'], ['Service Revenue'], ['Sales Revenue'],
    ['Sales Revenue — Materials'], ['Rental Income'], ['Interest Income'],
    ['Income Received'], ['Other Income'],
  ]});
  data.push({ range: "'⚙️ Config'!A62", values: [['  CLIENTS — for invoice autocomplete (one per row)']] });
  data.push({ range: "'⚙️ Config'!B63:E63", values: [['Client Name', '', '', 'Default Income Category']] });
  // EMPLOYEES — richer schema for age-aware payroll. One row per employee.
  // DOB drives CPP age branching (under-18 exempt). Relationship drives the
  // family-EI-exempt flag. SIN optional at setup (fill before year-end T4 gen).
  // TD1 claim codes default to '1' (federal/ON basic personal amount only);
  // override if the employee claims dependents or other credits.
  data.push({ range: "'⚙️ Config'!A75", values: [['  EMPLOYEES — for payroll and T4 generation (one per row)']] });
  data.push({ range: "'⚙️ Config'!B76:I76", values: [[
    'Name', 'DOB (YYYY-MM-DD)', 'SIN (optional)', 'Relationship',
    'Start Date', 'Default Rate $/hr', 'TD1 Fed Code', 'TD1 ON Code',
  ]]});
  // Blank placeholder rows — user enters via the app on first pay run.
  // (10 rows; add more by inserting below before row 86.)
  data.push({ range: "'⚙️ Config'!B77:I86", values: Array.from({length: 10}, () => ['', '', '', '', '', '', '', '']) });

  // TRANSACTIONS — single signed ledger (cash basis)
  // Replaces Income + Expenses. Positive E = money in, negative E = money out.
  // Category "Internal Transfer" is excluded from P&L and HST math.
  // HST amount (col H) is written directly by the importer / invoice confirmation —
  // no array formula needed because the API populates H at write time.
  data.push({ range: "'📒 Transactions'!A1", values: [[`TRANSACTIONS LEDGER  —  Cash basis  ·  One row per money movement  ·  + income, − expense`]] });
  data.push({ range: "'📒 Transactions'!B2", values: [['LEDGER SUMMARY']] });
  data.push({ range: "'📒 Transactions'!B3:F3", values: [['Total Revenue (excl HST)',  "='📊 Dashboard'!D5", '', '', '']]});
  data.push({ range: "'📒 Transactions'!B4:F4", values: [['Total HST Collected',       "='📊 Dashboard'!D6", '', '', '']]});
  data.push({ range: "'📒 Transactions'!B5:F5", values: [['Total Expenses (excl HST)', "='📊 Dashboard'!D7", '', '', '']]});
  data.push({ range: "'📒 Transactions'!B6:F6", values: [['Total HST Paid (ITCs)',     "='📊 Dashboard'!D8", '', '', '']]});
  data.push({ range: "'📒 Transactions'!B7:F7", values: [['NET INCOME before tax',     "='📊 Dashboard'!D9", '', '', '']]});
  data.push({ range: "'📒 Transactions'!B11:M11", values: [[
    'Date', 'Party (Client/Vendor)', 'Description', 'Amount (signed, excl HST)', 'Category',
    'HST?', `HST (${taxPct}%)`, 'Account', 'Source', 'Source Ref', 'Related Invoice #', 'Match Status',
  ]]});

  // INVOICES
  data.push({ range: "'🧾 Invoices'!A1", values: [[`INVOICE LOG  —  All invoices issued  ·  Outstanding tracking  ·  Status by colour`]] });
  data.push({ range: "'🧾 Invoices'!B2", values: [['INVOICE STATS']] });
  data.push({ range: "'🧾 Invoices'!B3:C3", values: [['Total invoices issued',   '=COUNTA(B12:B500)']] });
  data.push({ range: "'🧾 Invoices'!B4:C4", values: [['Average invoice value',   '=IFERROR(AVERAGE(F12:F500),0)']] });
  data.push({ range: "'🧾 Invoices'!B5:C5", values: [['Total invoiced excl HST', '=SUM(F12:F500)']] });
  data.push({ range: "'🧾 Invoices'!B6:C6", values: [['Total HST on invoices',   '=SUM(G12:G500)']] });
  data.push({ range: "'🧾 Invoices'!B7:C7", values: [['Total invoiced incl HST', '=SUM(H12:H500)']] });
  // Collected = full Total on Paid invoices + Deposit Amount on Deposit-Received invoices.
  // Outstanding = everything invoiced minus what's been collected.
  // (Awaiting Deposit + Unpaid contribute their full Total to outstanding.)
  data.push({ range: "'🧾 Invoices'!B8:C8", values: [['Outstanding — not yet paid', '=SUM(H12:H500)-SUMIF(K12:K500,"Paid",H12:H500)-SUMIF(K12:K500,"Deposit Received",O12:O500)']] });
  data.push({ range: "'🧾 Invoices'!B9:C9", values: [['Collected — paid + deposits', '=SUMIF(K12:K500,"Paid",H12:H500)+SUMIF(K12:K500,"Deposit Received",O12:O500)']] });
  data.push({ range: "'🧾 Invoices'!B11:Q11", values: [[
    'Invoice #', 'Date Issued', 'Client', 'Service Description', 'Amount (excl HST)',
    `HST (${taxPct}%)`, 'Total Invoiced', 'HST?', 'Due Date', 'Status', 'Date Paid', 'Notes',
    'Revenue Category', 'Deposit Amount', 'Deposit Date Received', 'Balance Due',
  ]]});

  // HST RETURNS
  data.push({ range: "'📋 HST Returns'!A1", values: [[`HST RETURN WORKBOOK  ·  ${prov} ${taxPct}%  ·  Quarterly Filing  ·  ${fye}`]] });
  data.push({ range: "'📋 HST Returns'!A2", values: [['Figures auto-pull from the 📒 Transactions tab. Enter on CRA My Business Account each quarter.']] });

  // Fiscal Year Start cell at C3 — parameterizes the quarterly period windows.
  // Defaults to April 1 of the current fiscal year (assumes Apr–Mar fiscal year,
  // which is standard for Ontario trades businesses). User can edit this one cell
  // to roll to a new fiscal year or shift to a calendar-year filing window.
  data.push({ range: "'📋 HST Returns'!B3", values: [['Fiscal Year Start (edit to roll forward):']] });
  data.push({ range: "'📋 HST Returns'!C3", values: [[
    '=DATE(YEAR(TODAY())-IF(MONTH(TODAY())<4,1,0),4,1)'
  ]]});

  data.push({ range: "'📋 HST Returns'!B4:G4", values: [[
    'LINE ITEM', 'Q1 Apr–Jun\nDue Jul 31', 'Q2 Jul–Sep\nDue Oct 31', 'Q3 Oct–Dec\nDue Jan 31', 'Q4 Jan–Mar\nDue Apr 30', 'FULL YEAR',
  ]]});

  // HST lines pull from 📒 Transactions.
  // Line 101 = Total Sales incl HST = sum of (Amount + HST) on positive rows excl Internal Transfer
  // Line 103 = HST Collected = sum of HST on positive rows excl Internal Transfer
  // Line 106 = ITCs = sum of HST on negative rows excl Internal Transfer
  // Date windows derive from $C$3 (Fiscal Year Start): Q1 = +0..2 months,
  // Q2 = +3..5, Q3 = +6..8, Q4 = +9..11 (all end-of-month via EOMONTH).
  const TXN_E = "'📒 Transactions'!E12:E1000";
  const TXN_F = "'📒 Transactions'!F12:F1000";
  const TXN_H = "'📒 Transactions'!H12:H1000";
  const TXN_B = "'📒 Transactions'!B12:B1000";
  const excl = `${TXN_F},"<>Internal Transfer"`;
  // q-indexed start/end offsets in months from FY start: Q1 = [0, 2], Q2 = [3, 5], ...
  const qStart = q => (q - 1) * 3;              // months to add to $C$3 for window start
  const qEndOffset = q => (q - 1) * 3 + 2;      // months to feed into EOMONTH for window end
  const dateGt = q => `${TXN_B},">="&EDATE($C$3,${qStart(q)})`;
  const dateLt = q => `${TXN_B},"<="&EOMONTH($C$3,${qEndOffset(q)})`;

  // Line 101 — Q1..Q4: (sum of Amount + sum of HST) within window
  const line101 = q =>
    `=SUMIFS(${TXN_E},${TXN_E},">0",${excl},${dateGt(q)},${dateLt(q)})`
    + `+SUMIFS(${TXN_H},${TXN_E},">0",${excl},${dateGt(q)},${dateLt(q)})`;
  // Line 103 — HST collected, positive rows, within window
  const line103 = q =>
    `=SUMIFS(${TXN_H},${TXN_E},">0",${excl},${dateGt(q)},${dateLt(q)})`;
  // Line 106 — ITCs, negative rows, within window
  const line106 = q =>
    `=SUMIFS(${TXN_H},${TXN_E},"<0",${excl},${dateGt(q)},${dateLt(q)})`;

  data.push({ range: "'📋 HST Returns'!B5:G8", values: [
    ['Total Sales (incl HST) — Line 101',
      line101(1), line101(2), line101(3), line101(4),
      "=SUM(C5:F5)"],
    ['HST Collected on Sales — Line 103',
      line103(1), line103(2), line103(3), line103(4),
      "=SUM(C6:F6)"],
    ['Input Tax Credits (ITCs) — Line 106',
      line106(1), line106(2), line106(3), line106(4),
      "=SUM(C7:F7)"],
    ['NET HST OWING (103 − 106) — Line 109','=C6-C7', '=D6-D7', '=E6-E7', '=F6-F7', '=G6-G7'],
  ]});
  data.push({ range: "'📋 HST Returns'!B11", values: [['FILING NOTES & REMINDERS']] });
  data.push({ range: "'📋 HST Returns'!B12:B18", values: [
    ['• Quarterly due: Q1 Jul 31  ·  Q2 Oct 31  ·  Q3 Jan 31  ·  Q4 Apr 30 (all year-end).'],
    ['• Pay via CRA My Business Account or online banking (Business Number + RT0001).'],
    ['• Meals & Entertainment: only 50% of HST paid is claimable — manually reduce Line 106 by the other 50%.'],
    ['• Keep ALL receipts 6 years — CRA audit window for HST is 4 years, income tax is 6 years.'],
    ['• Negative Line 109 = CRA owes you a refund — claim on My Business Account within 4 years.'],
    ['• AMEX CSV: download monthly statement, import via app — rows land in Transactions tab signed as expenses.'],
    ['• Quarterly windows derive from the Fiscal Year Start date in cell C3. Edit C3 to roll to a new fiscal year, or set Jan 1 for calendar-year filing.'],
  ]});

  // YEAR-END
  data.push({ range: "'📅 Year-End'!A1", values: [[`YEAR-END SUMMARY  ·  Corporate Tax Preparation  ·  ${fye} Year-End`]] });
  data.push({ range: "'📅 Year-End'!A2", values: [['⚠️ Working draft — review all figures with your T2 preparer before filing']] });
  data.push({ range: "'📅 Year-End'!A4", values: [['  INCOME SUMMARY']] });

  // Income rows pull positive-sign Transactions by category.
  const incomeRows = incomeCats.map(cat => [
    cat,
    `=SUMIFS('📒 Transactions'!E12:E1000,'📒 Transactions'!F12:F1000,"${cat}",'📒 Transactions'!E12:E1000,">0")`,
    '', '', '← from Transactions (cash basis)'
  ]);
  // TOTAL REVENUE uses an all-inclusive SUMIFS so it agrees with Dashboard D5.
  // Named category rows above show the bucketed breakdown; anything posted as
  // 'Income Received' (matched deposits without an explicit revenue category)
  // is included in the total but not in a named row. Delta = unbucketed deposits.
  data.push({ range: "'📅 Year-End'!B5:F11", values: [
    ...incomeRows,
    ['TOTAL REVENUE',
      "=SUMIFS('📒 Transactions'!E12:E1000,'📒 Transactions'!E12:E1000,\">0\",'📒 Transactions'!F12:F1000,\"<>Internal Transfer\")",
      '', '', '← all positive Transactions (T2 Schedule 1 Line 8000)'],
  ]});

  data.push({ range: "'📅 Year-End'!A12", values: [['  DEDUCTIBLE EXPENSES (T2 Schedule 1)']] });
  // Expense rows pull negative-sign Transactions by category, then negate so display is positive.
  const expRows = expenseCats.map(([cat, note]) => {
    let formula;
    if (cat === 'Meals & Entertainment (50%)') {
      formula = `=-SUMIFS('📒 Transactions'!E12:E1000,'📒 Transactions'!F12:F1000,"Meals & Entertainment",'📒 Transactions'!E12:E1000,"<0")*0.5`;
    } else {
      formula = `=-SUMIFS('📒 Transactions'!E12:E1000,'📒 Transactions'!F12:F1000,"${cat}",'📒 Transactions'!E12:E1000,"<0")`;
    }
    return [cat, formula, '', '', note || ''];
  });
  data.push({ range: "'📅 Year-End'!B13:F37", values: [
    ...expRows,
    ['TOTAL EXPENSES', `=SUM(C13:C${13 + expRows.length - 1})`, '', '', '← T2 Schedule 1 total'],
  ]});

  data.push({ range: "'📅 Year-End'!A39", values: [['  NET INCOME & TAX ESTIMATE']] });
  data.push({ range: "'📅 Year-End'!B40:F44", values: [
    ['Net Income before tax',                   '=C11-C37', '', '', ''],
    ['Small Business rate (ON+Fed combined)',   0.125,      '', '', '12.5% est. — verify with advisor'],
    ['Estimated corporate tax',                 '=C40*C41', '', '', ''],
    ['NET INCOME after estimated tax',          '=C40-C42', '', '', ''],
    ['', '', '', '', ''],
  ]});

  data.push({ range: "'📅 Year-End'!A46", values: [['  HST RECONCILIATION']] });
  data.push({ range: "'📅 Year-End'!B47:F50", values: [
    ['HST Collected (annual)', "=SUMIFS('📒 Transactions'!H12:H1000,'📒 Transactions'!E12:E1000,\">0\",'📒 Transactions'!F12:F1000,\"<>Internal Transfer\")",   '', '', 'from Transactions'],
    ['ITCs Claimed (annual)',  "=SUMIFS('📒 Transactions'!H12:H1000,'📒 Transactions'!E12:E1000,\"<0\",'📒 Transactions'!F12:F1000,\"<>Internal Transfer\")", '', '', 'from Transactions'],
    ['Net HST Owing',          '=C47-C48',                      '', '', ''],
    ['', '', '', '', ''],
  ]});

  data.push({ range: "'📅 Year-End'!A51", values: [['  TAX MINIMIZATION CHECKLIST']] });
  data.push({ range: "'📅 Year-End'!B52:F60", values: [
    ['Salary vs dividend split reviewed?',              '☐ Yes', '☐ No', '☐ N/A', ''],
    ['RRSP for owner-employees maximized?',             '☐ Yes', '☐ No', '☐ N/A', ''],
    ['CCA on equipment/vehicles claimed?',              '☐ Yes', '☐ No', '☐ N/A', 'Vehicles may qualify — ask advisor'],
    ['Management fees between related corps reviewed?', '☐ Yes', '☐ No', '☐ N/A', ''],
    ['Home office portion documented?',                 '☐ Yes', '☐ No', '☐ N/A', ''],
    ['Vehicle logbook maintained?',                     '☐ Yes', '☐ No', '☐ N/A', 'CRA requirement'],
    ['All HST/payroll remittances current?',            '☐ Yes', '☐ No', '☐ N/A', ''],
    ['LCGE planning if selling business?',              '☐ Yes', '☐ No', '☐ N/A', 'Lifetime Capital Gains Exemption'],
    ['', '', '', '', ''],
  ]});

  // ─── PAYROLL TAB ───
  data.push({ range: "'💼 Payroll'!A1", values: [[
    'PAYROLL LEDGER  —  Pay runs · Source deduction tracking · T4 source of truth'
  ]]});
  data.push({ range: "'💼 Payroll'!B2", values: [['PAYROLL SUMMARY  (YTD)']] });
  // Summary block — computed from Payroll rows below
  data.push({ range: "'💼 Payroll'!B3:H3", values: [[
    'Total Gross Wages (YTD)', '=IFERROR(SUM(I12:I500),0)', '', '', '', '', ''
  ]]});
  data.push({ range: "'💼 Payroll'!B4:H4", values: [[
    'CPP Withheld (YTD)',      '=IFERROR(SUM(J12:J500),0)', '', '', '', '', ''
  ]]});
  data.push({ range: "'💼 Payroll'!B5:H5", values: [[
    'Fed + ON Tax Withheld (YTD)', '=IFERROR(SUM(L12:L500)+SUM(M12:M500),0)', '', '', '', '', ''
  ]]});
  data.push({ range: "'💼 Payroll'!B6:H6", values: [[
    'Outstanding Remittance Owed', '=IFERROR(SUMIFS(J12:J500,Q12:Q500,"Paid")+SUMIFS(L12:L500,Q12:Q500,"Paid")+SUMIFS(M12:M500,Q12:Q500,"Paid"),0)', '', '', '', '', ''
  ]]});
  data.push({ range: "'💼 Payroll'!B11:Q11", values: [[
    'Pay Date', 'Employee', 'Age', 'Business', 'Work Description',
    'Hours', 'Rate', 'Gross', 'CPP (ee)', 'EI (ee)', 'Fed Tax', 'ON Tax',
    'Net Pay', 'YTD Gross', 'Remittance Due', 'Status',
  ]]});

  // ─── WORK LOG TAB ───
  data.push({ range: "'📝 Work Log'!A1", values: [[
    'WORK LOG  —  Contemporaneous entry · Timestamped · Audit-defence paper trail'
  ]]});
  data.push({ range: "'📝 Work Log'!B2", values: [[
    'Enter work as it happens. Server-stamps the entry time. A "corrected" flag appears on edits — do NOT silently modify past rows.'
  ]]});
  data.push({ range: "'📝 Work Log'!B11:I11", values: [[
    'Date', 'Employee', 'Business', 'Task Description', 'Hours', 'Rate', 'Notes', 'Entry Audit',
  ]]});

  const valuesResp = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values:batchUpdate`, {
    method: 'POST', headers: authHeader,
    body: JSON.stringify({ valueInputOption: 'USER_ENTERED', data }),
  });
  const valuesData = await valuesResp.json();
  if (valuesData.error) {
    return { ok: false, step: 'populateValues', error: valuesData.error.message || JSON.stringify(valuesData.error) };
  }
  return { ok: true };
}

