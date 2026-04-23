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
            gridProperties: { rowCount: 100, columnCount: 8 },
            tabColor: COLORS.green } },
        { properties: { sheetId: 300, title: '💰 Income', index: 2,
            gridProperties: { rowCount: 500, columnCount: 12, frozenRowCount: 11 },
            tabColor: COLORS.green } },
        { properties: { sheetId: 400, title: '💸 Expenses', index: 3,
            gridProperties: { rowCount: 500, columnCount: 13, frozenRowCount: 11 },
            tabColor: COLORS.red } },
        { properties: { sheetId: 500, title: '🧾 Invoices', index: 4,
            gridProperties: { rowCount: 500, columnCount: 13, frozenRowCount: 11 },
            tabColor: COLORS.blue } },
        { properties: { sheetId: 600, title: '📋 HST Returns', index: 5,
            gridProperties: { rowCount: 40, columnCount: 10 },
            tabColor: COLORS.teal } },
        { properties: { sheetId: 700, title: '📅 Year-End', index: 6,
            gridProperties: { rowCount: 80, columnCount: 8 },
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
  const DASH = 100, CFG = 200, INC = 300, EXP = 400, INV = 500, HST = 600, YE = 700;
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
  requests.push(...sectionRequest(CFG, 74, 0, 6, COLORS.green));
  requests.push(cellFormat(CFG, 75, 1, 76, 6, { textFormat: { bold: true, foregroundColor: COLORS.textMuted, fontSize: 9 } }));
  requests.push(cellFormat(CFG, 76, 1, 86, 6, FMT_EDITABLE));
  requests.push(colWidth(CFG, 0, 1, 30));
  requests.push(colWidth(CFG, 1, 2, 180));
  requests.push(colWidth(CFG, 2, 3, 260));
  requests.push(colWidth(CFG, 3, 4, 30));
  requests.push(colWidth(CFG, 4, 5, 180));
  requests.push(colWidth(CFG, 5, 6, 260));

  // ─── INCOME ───
  requests.push(...bannerRequest(INC, 0, 0, 11, COLORS.green));
  requests.push(...sectionRequest(INC, 1, 1, 6, COLORS.green));
  requests.push(cellFormat(INC, 2, 1, 10, 2, { textFormat: { bold: true, fontSize: 10 } }));
  requests.push(cellFormat(INC, 2, 2, 10, 6, { backgroundColor: COLORS.greenTint, numberFormat: FMT_CURRENCY.numberFormat }));
  requests.push(headerRowRequest(INC, 10, 1, 11, COLORS.green));
  requests.push({ updateDimensionProperties: { range: { sheetId: INC, dimension: 'ROWS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  requests.push(bandingRequest(INC, 11, 500, 1, 11, COLORS.greenTint));
  requests.push(cellFormat(INC, 11, 1, 500, 2, FMT_DATE));
  requests.push(cellFormat(INC, 11, 5, 500, 6, FMT_CURRENCY));
  requests.push(cellFormat(INC, 11, 7, 500, 8, FMT_CURRENCY));
  requests.push(cellFormat(INC, 11, 8, 500, 9, FMT_CURRENCY));
  requests.push(colWidth(INC, 0, 1, 30));
  requests.push(colWidth(INC, 1, 2, 100));
  requests.push(colWidth(INC, 2, 3, 180));
  requests.push(colWidth(INC, 3, 4, 220));
  requests.push(colWidth(INC, 4, 5, 90));
  requests.push(colWidth(INC, 5, 6, 120));
  requests.push(colWidth(INC, 6, 7, 170));
  requests.push(colWidth(INC, 7, 8, 110));
  requests.push(colWidth(INC, 8, 9, 120));
  requests.push(colWidth(INC, 9, 10, 120));
  requests.push(colWidth(INC, 10, 11, 200));

  // ─── EXPENSES ───
  requests.push(...bannerRequest(EXP, 0, 0, 12, COLORS.red));
  requests.push(...sectionRequest(EXP, 1, 1, 6, COLORS.red));
  requests.push(cellFormat(EXP, 2, 1, 10, 2, { textFormat: { bold: true, fontSize: 10 } }));
  requests.push(cellFormat(EXP, 2, 2, 10, 6, { backgroundColor: COLORS.redTint, numberFormat: FMT_CURRENCY.numberFormat }));
  requests.push(headerRowRequest(EXP, 10, 1, 12, COLORS.red));
  requests.push({ updateDimensionProperties: { range: { sheetId: EXP, dimension: 'ROWS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  requests.push(bandingRequest(EXP, 11, 500, 1, 12, COLORS.redTint));
  requests.push(cellFormat(EXP, 11, 1, 500, 2, FMT_DATE));
  requests.push(cellFormat(EXP, 11, 4, 500, 5, FMT_CURRENCY));
  requests.push(cellFormat(EXP, 11, 6, 500, 7, FMT_CURRENCY));
  requests.push(cellFormat(EXP, 11, 7, 500, 8, FMT_CURRENCY));
  requests.push(colWidth(EXP, 0, 1, 30));
  requests.push(colWidth(EXP, 1, 2, 100));
  requests.push(colWidth(EXP, 2, 3, 180));
  requests.push(colWidth(EXP, 3, 4, 220));
  requests.push(colWidth(EXP, 4, 5, 110));
  requests.push(colWidth(EXP, 5, 6, 170));
  requests.push(colWidth(EXP, 6, 7, 110));
  requests.push(colWidth(EXP, 7, 8, 110));
  requests.push(colWidth(EXP, 8, 9, 120));
  requests.push(colWidth(EXP, 9, 10, 80));
  requests.push(colWidth(EXP, 10, 11, 120));
  requests.push(colWidth(EXP, 11, 12, 200));

  // ─── INVOICES ───
  requests.push(...bannerRequest(INV, 0, 0, 12, COLORS.blue));
  requests.push(...sectionRequest(INV, 1, 1, 6, COLORS.blue));
  requests.push(cellFormat(INV, 2, 1, 10, 2, { textFormat: { bold: true, fontSize: 10 } }));
  requests.push(cellFormat(INV, 2, 2, 10, 6, { backgroundColor: COLORS.blueTint }));
  requests.push(headerRowRequest(INV, 10, 1, 12, COLORS.blue));
  requests.push({ updateDimensionProperties: { range: { sheetId: INV, dimension: 'ROWS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  requests.push(bandingRequest(INV, 11, 500, 1, 12, COLORS.blueTint));
  requests.push(cellFormat(INV, 11, 2, 500, 3, FMT_DATE));
  requests.push(cellFormat(INV, 11, 5, 500, 6, FMT_CURRENCY));
  requests.push(cellFormat(INV, 11, 6, 500, 7, FMT_CURRENCY));
  requests.push(cellFormat(INV, 11, 7, 500, 8, FMT_CURRENCY));
  requests.push(cellFormat(INV, 11, 9, 500, 10, FMT_DATE));
  requests.push(cellFormat(INV, 11, 11, 500, 12, FMT_DATE));
  requests.push({
    setDataValidation: {
      range: { sheetId: INV, startRowIndex: 11, endRowIndex: 500, startColumnIndex: 8, endColumnIndex: 9 },
      rule: { condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: 'Yes' }, { userEnteredValue: 'No' }] }, showCustomUi: true },
    }
  });
  requests.push({
    setDataValidation: {
      range: { sheetId: INV, startRowIndex: 11, endRowIndex: 500, startColumnIndex: 10, endColumnIndex: 11 },
      rule: { condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: 'Unpaid' }, { userEnteredValue: 'Paid' }, { userEnteredValue: 'Overdue' }, { userEnteredValue: 'Cancelled' }] }, showCustomUi: true },
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
  requests.push(colWidth(INV, 9, 10, 110));
  requests.push(colWidth(INV, 10, 11, 90));
  requests.push(colWidth(INV, 11, 12, 110));
  requests.push(colWidth(INV, 12, 13, 200));

  // ─── HST RETURNS ───
  requests.push(...bannerRequest(HST, 0, 0, 7, COLORS.teal));
  requests.push(cellFormat(HST, 1, 0, 2, 7, { horizontalAlignment: 'CENTER', textFormat: { italic: true, foregroundColor: COLORS.textMuted, fontSize: 9 } }));
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
  data.push({ range: "'📊 Dashboard'!A1", values: [[`${bizName.toUpperCase()}  ·  FISCAL YE ${fye.toUpperCase()}  ·  ${prov} HST ${taxPct}%`]] });
  data.push({ range: "'📊 Dashboard'!A3", values: [['PROFIT & LOSS SUMMARY']] });
  data.push({ range: "'📊 Dashboard'!B5:B10", values: [
    ['Total Revenue (excl HST)'], ['HST Collected'],
    ['Total Expenses (excl HST)'], ['HST Paid — ITCs'],
    ['NET INCOME before tax'], ['HST OWING / (Refund)']
  ]});
  data.push({ range: "'📊 Dashboard'!D5:D10", values: [
    ["=SUMIF('💰 Income'!G12:G500,\"<>\",'💰 Income'!F12:F500)"],
    ["=SUM('💰 Income'!H12:H500)"],
    ["=SUMIF('💸 Expenses'!F12:F500,\"<>\",'💸 Expenses'!E12:E500)"],
    ["=SUM('💸 Expenses'!G12:G500)"],
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
    ['Bill Payment / Transfer'], ['Owner Draw / Distribution'], ['SKIP — not a business expense'],
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
  data.push({ range: "'⚙️ Config'!A75", values: [['  EMPLOYEES — for payroll and T4 generation (one per row)']] });
  data.push({ range: "'⚙️ Config'!B76:F76", values: [['Employee Name', 'SIN (last 3)', '', 'Annual Salary/Wages', 'Status']] });

  // INCOME
  data.push({ range: "'💰 Income'!A1", values: [[`INCOME TRACKER  —  One row per payment received  ·  HST auto-calculated`]] });
  data.push({ range: "'💰 Income'!B2", values: [['INCOME SUMMARY']] });
  data.push({ range: "'💰 Income'!B3:F3", values: [['Total Revenue (excl HST)', "='📊 Dashboard'!D5", '', '', '']]});
  data.push({ range: "'💰 Income'!B4:F4", values: [['Total HST Collected',      "='📊 Dashboard'!D6", '', '', '']]});
  data.push({ range: "'💰 Income'!B11:L11", values: [[
    'Date', 'Client / Payer', 'Description', 'Invoice #', 'Amount (excl HST)',
    'Category', `HST (${taxPct}%)`, 'Total Received', 'Payment Method', 'Source', 'Source Ref',
  ]]});
  data.push({ range: "'💰 Income'!H12", values: [[
    `=ARRAYFORMULA(IF(F12:F500="","",IF(REGEXMATCH(G12:G500,"Interest Income|Income Received"),0,F12:F500*${taxRate})))`
  ]]});
  data.push({ range: "'💰 Income'!I12", values: [[
    '=ARRAYFORMULA(IF(F12:F500="","",F12:F500+H12:H500))'
  ]]});

  // EXPENSES
  data.push({ range: "'💸 Expenses'!A1", values: [[`EXPENSE TRACKER  —  Enter total charged  ·  HST auto-extracted  ·  Net calculated`]] });
  data.push({ range: "'💸 Expenses'!B2", values: [['EXPENSE SUMMARY']] });
  data.push({ range: "'💸 Expenses'!B3:F3", values: [['Total Expenses (excl HST)', "='📊 Dashboard'!D7", '', '', '']]});
  data.push({ range: "'💸 Expenses'!B4:F4", values: [['Total HST Paid (ITCs)',     "='📊 Dashboard'!D8", '', '', '']]});
  data.push({ range: "'💸 Expenses'!B11:M11", values: [[
    'Date', 'Vendor / Payee', 'Description', 'Amount (excl HST)', 'Category',
    `HST (auto ${taxPct}%)`, 'Total (incl HST)', 'Payment Method', 'On AMEX?', 'Source', 'Source Ref', 'Receipt',
  ]]});
  data.push({ range: "'💸 Expenses'!G12", values: [[
    `=ARRAYFORMULA(IF(E12:E500="","",IF(REGEXMATCH(F12:F500,"Wages|Insurance|Interest & Bank|Property Tax|Bill Payment|Owner Draw|Income Received|SKIP|Subcontractors"),0,E12:E500*${taxRate})))`
  ]]});
  data.push({ range: "'💸 Expenses'!H12", values: [[
    '=ARRAYFORMULA(IF(E12:E500="","",E12:E500+G12:G500))'
  ]]});

  // INVOICES
  data.push({ range: "'🧾 Invoices'!A1", values: [[`INVOICE LOG  —  All invoices issued  ·  Outstanding tracking  ·  Status by colour`]] });
  data.push({ range: "'🧾 Invoices'!B2", values: [['INVOICE STATS']] });
  data.push({ range: "'🧾 Invoices'!B3:C3", values: [['Total invoices issued',   '=COUNTA(B12:B500)']] });
  data.push({ range: "'🧾 Invoices'!B4:C4", values: [['Average invoice value',   '=IFERROR(AVERAGE(F12:F500),0)']] });
  data.push({ range: "'🧾 Invoices'!B5:C5", values: [['Total invoiced excl HST', '=SUM(F12:F500)']] });
  data.push({ range: "'🧾 Invoices'!B6:C6", values: [['Total HST on invoices',   '=SUM(G12:G500)']] });
  data.push({ range: "'🧾 Invoices'!B7:C7", values: [['Total invoiced incl HST', '=SUM(H12:H500)']] });
  data.push({ range: "'🧾 Invoices'!B8:C8", values: [['Outstanding — unpaid',    '=SUMIF(K12:K500,"Unpaid",H12:H500)']] });
  data.push({ range: "'🧾 Invoices'!B9:C9", values: [['Collected — paid',        '=SUMIF(K12:K500,"Paid",H12:H500)']] });
  data.push({ range: "'🧾 Invoices'!B11:M11", values: [[
    'Invoice #', 'Date Issued', 'Client', 'Service Description', 'Amount (excl HST)',
    `HST (${taxPct}%)`, 'Total Invoiced', 'HST?', 'Due Date', 'Status', 'Date Paid', 'Notes',
  ]]});

  // HST RETURNS
  data.push({ range: "'📋 HST Returns'!A1", values: [[`HST RETURN WORKBOOK  ·  ${prov} ${taxPct}%  ·  Quarterly Filing  ·  ${fye}`]] });
  data.push({ range: "'📋 HST Returns'!A2", values: [['Figures auto-pull from Income & Expense tabs. Enter on CRA My Business Account each quarter.']] });
  data.push({ range: "'📋 HST Returns'!B4:G4", values: [[
    'LINE ITEM', 'Q1 Apr–Jun\nDue Jul 31', 'Q2 Jul–Sep\nDue Oct 31', 'Q3 Oct–Dec\nDue Jan 31', 'Q4 Jan–Mar\nDue Apr 30', 'FULL YEAR',
  ]]});
  data.push({ range: "'📋 HST Returns'!B5:G8", values: [
    ['Total Sales (incl HST) — Line 101',  '', '', '', "=SUMPRODUCT(('💰 Income'!B12:B500<>\"\")*1*('💰 Income'!I12:I500))", "=SUM(C5:F5)"],
    ['HST Collected on Sales — Line 103',  '', '', '', "=SUMPRODUCT(('💰 Income'!B12:B500<>\"\")*1*('💰 Income'!H12:H500))", "=SUM(C6:F6)"],
    ['Input Tax Credits (ITCs) — Line 106','', '', '', "=SUMPRODUCT(('💸 Expenses'!B12:B500<>\"\")*1*('💸 Expenses'!G12:G500))", "=SUM(C7:F7)"],
    ['NET HST OWING (103 − 106) — Line 109','=C6-C7', '=D6-D7', '=E6-E7', '=F6-F7', '=G6-G7'],
  ]});
  data.push({ range: "'📋 HST Returns'!B11", values: [['FILING NOTES & REMINDERS']] });
  data.push({ range: "'📋 HST Returns'!B12:B17", values: [
    ['• Quarterly due: Q1 Jul 31  ·  Q2 Oct 31  ·  Q3 Jan 31  ·  Q4 Apr 30 (all year-end).'],
    ['• Pay via CRA My Business Account or online banking (Business Number + RT0001).'],
    ['• Meals & Entertainment: only 50% of HST paid is claimable — manually reduce Line 106 by the other 50%.'],
    ['• Keep ALL receipts 6 years — CRA audit window for HST is 4 years, income tax is 6 years.'],
    ['• Negative Line 109 = CRA owes you a refund — claim on My Business Account within 4 years.'],
    ['• AMEX CSV: download monthly statement, import to Expenses tab, categorize each line.'],
  ]});

  // YEAR-END
  data.push({ range: "'📅 Year-End'!A1", values: [[`YEAR-END SUMMARY  ·  Corporate Tax Preparation  ·  ${fye} Year-End`]] });
  data.push({ range: "'📅 Year-End'!A2", values: [['⚠️ Working draft — review all figures with your T2 preparer before filing']] });
  data.push({ range: "'📅 Year-End'!A4", values: [['  INCOME SUMMARY']] });

  const incomeRows = incomeCats.map(cat => [
    cat,
    `=SUMIF('💰 Income'!G12:G500,"${cat}",'💰 Income'!F12:F500)`,
    '', '', '← from Income tab'
  ]);
  data.push({ range: "'📅 Year-End'!B5:F11", values: [
    ...incomeRows,
    ['TOTAL REVENUE', '=SUM(C5:C10)', '', '', '← T2 Schedule 1 Line 8000'],
  ]});

  data.push({ range: "'📅 Year-End'!A12", values: [['  DEDUCTIBLE EXPENSES (T2 Schedule 1)']] });
  const expRows = expenseCats.map(([cat, note]) => {
    let formula;
    if (cat === 'Meals & Entertainment (50%)') {
      formula = `=SUMIF('💸 Expenses'!F12:F500,"Meals & Entertainment",'💸 Expenses'!E12:E500)*0.5`;
    } else {
      formula = `=SUMIF('💸 Expenses'!F12:F500,"${cat}",'💸 Expenses'!E12:E500)`;
    }
    return [cat, formula, '', '', note || ''];
  });
  data.push({ range: "'📅 Year-End'!B13:F37", values: [
    ...expRows,
    ['TOTAL EXPENSES', `=SUM(C13:C${13 + expRows.length - 1})`, '', '', '← T2 Schedule 1 total'],
  ]});

  data.push({ range: "'📅 Year-End'!A39", values: [['  NET INCOME & TAX ESTIMATE']] });
  data.push({ range: "'📅 Year-End'!B40:F44", values: [
    ['Net Income before tax',                   '=C10-C37', '', '', ''],
    ['Small Business rate (ON+Fed combined)',   0.125,      '', '', '12.5% est. — verify with advisor'],
    ['Estimated corporate tax',                 '=C40*C41', '', '', ''],
    ['NET INCOME after estimated tax',          '=C40-C42', '', '', ''],
    ['', '', '', '', ''],
  ]});

  data.push({ range: "'📅 Year-End'!A46", values: [['  HST RECONCILIATION']] });
  data.push({ range: "'📅 Year-End'!B47:F50", values: [
    ['HST Collected (annual)', "=SUM('💰 Income'!H12:H500)",   '', '', 'from HST Returns tab'],
    ['ITCs Claimed (annual)',  "=SUM('💸 Expenses'!G12:G500)", '', '', 'from HST Returns tab'],
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

