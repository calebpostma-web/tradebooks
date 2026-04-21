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

    if (body.code) {
      return handleTokenExchange(body.code, origin, env, headers);
    }

    if (body.action === 'create-sheet' && body.accessToken) {
      return handleCreateSheet(body.accessToken, body.profile || {}, env, headers);
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
async function handleTokenExchange(code, origin, env, headers) {
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

  return new Response(JSON.stringify({
    ok: true,
    accessToken: tokenData.access_token,
    refreshToken: tokenData.refresh_token,
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
async function handleCreateSheet(accessToken, profile, env, headers) {
  const authHeader = { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' };

  // Create the Google Sheet with all 7 tabs
  const sheetResp = await fetch('https://sheets.googleapis.com/v4/spreadsheets', {
    method: 'POST',
    headers: authHeader,
    body: JSON.stringify({
      properties: {
        title: `${profile.businessName || 'My Business'} — AI Bookkeeper`,
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

  // Apply styling and populate values
  await applyStyling(accessToken, spreadsheetId);
  await populateValues(accessToken, spreadsheetId, profile);

  // Create Apps Script project
  const scriptResp = await fetch('https://script.googleapis.com/v1/projects', {
    method: 'POST',
    headers: authHeader,
    body: JSON.stringify({ title: 'AI Bookkeeper Add-on', parentId: spreadsheetId }),
  });

  const scriptData = await scriptResp.json();
  if (scriptData.error) {
    return new Response(JSON.stringify({
      ok: true, spreadsheetId, spreadsheetUrl,
      scriptError: scriptData.error.message || 'Could not create Apps Script',
      scriptUrl: null,
    }), { headers });
  }

  const scriptId = scriptData.scriptId;

  // Push Apps Script code
  await fetch(`https://script.googleapis.com/v1/projects/${scriptId}/content`, {
    method: 'PUT',
    headers: authHeader,
    body: JSON.stringify({
      files: [
        { name: 'Code', type: 'SERVER_JS', source: getAppsScriptCode() },
        { name: 'appsscript', type: 'JSON', source: getAppsScriptManifest() },
      ],
    }),
  });

  // Deploy as web app
  let scriptUrl = null;
  try {
    const versionResp = await fetch(`https://script.googleapis.com/v1/projects/${scriptId}/versions`, {
      method: 'POST',
      headers: authHeader,
      body: JSON.stringify({ description: 'AI Bookkeeper v1' }),
    });
    const versionData = await versionResp.json();

    if (versionData.versionNumber) {
      const deployResp = await fetch(`https://script.googleapis.com/v1/projects/${scriptId}/deployments`, {
        method: 'POST',
        headers: authHeader,
        body: JSON.stringify({
          versionNumber: versionData.versionNumber,
          manifestFileName: 'appsscript',
          description: 'AI Bookkeeper Web App',
        }),
      });
      const deployData = await deployResp.json();
      if (deployData.entryPoints) {
        const webApp = deployData.entryPoints.find(e => e.entryPointType === 'WEB_APP');
        if (webApp && webApp.webApp && webApp.webApp.url) scriptUrl = webApp.webApp.url;
      }
    }
  } catch (deployErr) { /* non-blocking */ }

  // Write script URL back to Config F10
  if (scriptUrl) {
    try {
      await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/%E2%9A%99%EF%B8%8F%20Config!F10?valueInputOption=RAW`, {
        method: 'PUT',
        headers: authHeader,
        body: JSON.stringify({ values: [[scriptUrl]] }),
      });
    } catch (e) { /* non-critical */ }
  }

  return new Response(JSON.stringify({
    ok: true, spreadsheetId, spreadsheetUrl, scriptId, scriptUrl,
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

  await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`, {
    method: 'POST', headers: authHeader,
    body: JSON.stringify({ requests }),
  });
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
  data.push({ range: "'⚙️ Config'!A1", values: [['⚙️ AI BOOKKEEPER CONFIGURATION']] });
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
  data.push({ range: "'📅 Year-End'!B5:F10", values: [
    ...incomeRows,
    ['TOTAL REVENUE', '=SUM(C5:C9)', '', '', '← T2 Schedule 1 Line 8000'],
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

  await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values:batchUpdate`, {
    method: 'POST', headers: authHeader,
    body: JSON.stringify({ valueInputOption: 'USER_ENTERED', data }),
  });
}


// ════════════════════════════════════════════════════════════════════
// APPS SCRIPT CODE
// ════════════════════════════════════════════════════════════════════
function getAppsScriptCode() {
  return `
const EXPENSES_TAB='💸 Expenses',INCOME_TAB='💰 Income',INVOICES_TAB='🧾 Invoices',CONFIG_TAB='⚙️ Config',DATA_START_ROW=12;
const DEFAULT_INCOME_CATS=['Consulting Revenue','Service Revenue','Sales Revenue','Sales Revenue — Materials','Rental Income','Interest Income','Income Received','Other Income'];
const DEFAULT_SKIP_CATS=['Bill Payment / Transfer','Owner Draw / Distribution','SKIP — not a business expense'];
const COMMON_EXPENSE_CATS=['Fuel','Vehicle Repairs','Insurance','Telephone','Supplies','Small Tools','Meals & Entertainment','Professional Fees','Interest & Bank Charges'];

function onOpen(){SpreadsheetApp.getUi().createMenu('📚 Bookkeeping').addItem('🏥 Health Check','runHealthCheck').addItem('📋 Summary','runSummary').addItem('📊 Show Config','showConfig').addToUi()}

function showConfig(){const c=readConfig(SpreadsheetApp.getActiveSpreadsheet());SpreadsheetApp.getUi().alert('Config:\\n\\nBusiness: '+(c.businessName||'Not set')+'\\nProvince: '+(c.province||'ON')+'\\nType: '+(c.businessType||'Not set'))}

function runHealthCheck(){const r=buildHealthCheck(SpreadsheetApp.getActiveSpreadsheet());const l=['Score: '+r.score+'/100 ('+r.grade+')','',r.summary,''];r.issues.forEach(i=>{l.push((i.severity==='error'?'🔴':i.severity==='warning'?'🟡':'ℹ️')+' '+i.message)});SpreadsheetApp.getUi().alert(l.join('\\n'))}

function runSummary(){const r=buildSummary(SpreadsheetApp.getActiveSpreadsheet());SpreadsheetApp.getUi().alert('Summary\\n\\nRevenue: $'+r.totalRevenue+'\\nExpenses: $'+r.totalExpenses+'\\nNet: $'+r.netIncome+'\\nHST Owing: $'+r.hstNetOwing)}

function readConfig(ss){const c=ss.getSheetByName(CONFIG_TAB);if(!c)return{customExpenseCats:[],customIncomeCats:[],clients:[],employees:[]};const l=c.getRange('B5:C13').getValues(),r=c.getRange('E5:F13').getValues();const cfg={businessName:l[0][1],tradingName:l[1][1],ownerName:l[2][1],city:l[3][1],province:l[4][1]||'ON',taxRate:parseFloat(l[5][1])||0.13,hstNumber:l[6][1],businessType:l[7][1]||'sole_prop',fiscalYearEnd:l[8][1]||'December 31',primaryBank:r[0][1]||'BMO',creditCard:r[1][1]||'AMEX',startingInvoice:parseInt(r[2][1])||1001,homeOfficePercent:parseInt(r[3][1])||0,email:r[4][1]||'',appsScriptUrl:r[5][1]||'',structure:r[6][1]||'',activities:r[7][1]||''};cfg.customExpenseCats=c.getRange('E17:E36').getValues().flat().filter(v=>v&&String(v).trim());const all=c.getDataRange().getValues();cfg.customIncomeCats=[];cfg.clients=[];cfg.employees=[];let sec='';for(let i=0;i<all.length;i++){const a=String(all[i][0]||'');if(a.includes('CUSTOM INCOME'))sec='inc';if(a.includes('CLIENTS'))sec='cli';if(a.includes('EMPLOYEES'))sec='emp';if(sec==='inc'){const v=String(all[i][4]||'').trim();if(v&&!v.includes('Custom'))cfg.customIncomeCats.push(v)}if(sec==='cli'){const n=String(all[i][1]||'').trim(),cat=String(all[i][4]||'').trim();if(n&&n!=='Client Name')cfg.clients.push({name:n,defaultCategory:cat||'Consulting Revenue'})}if(sec==='emp'){const n=String(all[i][1]||'').trim(),s=String(all[i][2]||'').trim(),sal=parseFloat(all[i][4])||0,st=String(all[i][5]||'Active').trim();if(n&&n!=='Employee Name')cfg.employees.push({name:n,sin3:s,salary:sal,status:st})}}return cfg}

function doGet(e){try{const ss=SpreadsheetApp.getActiveSpreadsheet(),a=(e&&e.parameter&&e.parameter.action)||'';if(a==='summary')return respond(buildSummary(ss));if(a==='health')return respond(buildHealthCheck(ss));if(a==='config')return respond({ok:true,config:readConfig(ss)});if(a==='categories')return respond(buildCategoryLists(ss));const c=readConfig(ss);return respond({ok:true,message:'AI Bookkeeper v4 ✓',business:c.businessName||'Not configured',province:c.province||'ON'})}catch(e){return respond({ok:false,error:e.message})}}

function doPost(e){try{const d=JSON.parse(e.postData.contents),ss=SpreadsheetApp.getActiveSpreadsheet();if(d.type==='invoice')return handleInvoice(ss,d.invoice);if(d.type==='receipt'&&d.image)return handleReceiptUpload(ss,d);return handleRows(ss,d.rows||[],d.bank||'AMEX',d.source||'Import')}catch(e){return respond({ok:false,error:e.message})}}

function handleReceiptUpload(ss,d){try{const folder=getOrCreateReceiptFolder(),today=new Date(),year=today.getFullYear(),yf=getOrCreateYearFolder(folder,year);const bytes=Utilities.base64Decode(d.image.split(',').pop()||d.image),blob=Utilities.newBlob(bytes,d.mimeType||'image/jpeg',d.filename||('Receipt_'+today.getTime()+'.jpg'));const file=yf.createFile(blob);file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);return respond({ok:true,driveUrl:file.getUrl(),fileId:file.getId()})}catch(e){return respond({ok:false,error:'Drive upload: '+e.message})}}
function getOrCreateReceiptFolder(){const name='AI Bookkeeper Receipts',fs=DriveApp.getFoldersByName(name);if(fs.hasNext())return fs.next();return DriveApp.createFolder(name)}
function getOrCreateYearFolder(parent,year){const fs=parent.getFoldersByName(String(year));if(fs.hasNext())return fs.next();return parent.createFolder(String(year))}

function buildCategoryLists(ss){const c=readConfig(ss);const de=['Meals & Entertainment','Professional Fees','Wages & Salaries','Small Tools','Supplies','Uniforms','Dues & Memberships','Vehicle Repairs','Fuel','Insurance','Interest & Bank Charges','Repairs & Maintenance','Property Tax','Telephone','Utilities','Conferences','Equipment Purchase','Inventory — Materials (COGS)','Advertising & Marketing','Subcontractors','Office Supplies','Rent','Home Office','Vehicle Lease/Payments','Travel','Training & Education','Permits & Licenses','Tax Payments','Other','Bill Payment / Transfer','Owner Draw / Distribution','Income Received','SKIP — not a business expense'];const di=['Consulting Revenue','Service Revenue','Sales Revenue','Sales Revenue — Materials','Rental Income','Interest Income','Income Received','Other Income'];const me=[...de];const oi=me.indexOf('Other');(c.customExpenseCats||[]).forEach(x=>{if(!me.includes(x))me.splice(oi,0,x)});const mi=[...di];(c.customIncomeCats||[]).forEach(x=>{if(!mi.includes(x))mi.splice(mi.length-1,0,x)});return{ok:true,expenseCategories:me,incomeCategories:mi,clients:(c.clients||[]).map(x=>x.name),clientDefaults:c.clients||[]}}

function handleInvoice(ss,inv){const is=ss.getSheetByName(INVOICES_TAB),ic=ss.getSheetByName(INCOME_TAB);if(!is)return respond({ok:false,error:'Tab not found: '+INVOICES_TAB});if(!ic)return respond({ok:false,error:'Tab not found: '+INCOME_TAB});const ir=gnr(is);is.getRange(ir,2,1,10).setValues([[inv.invNum||'',inv.dateVal||'',inv.client||'',inv.desc||'',parseFloat(inv.sub)||0,parseFloat(inv.hstAmt)||0,parseFloat(inv.total)||0,inv.hst||'Yes',inv.dueVal||'','Unpaid']]);const icr=gnr(ic);const sr=gsr('INV',inv.dateVal,inv.sub,inv.client);ic.getRange(icr,2,1,6).setValues([[inv.dateVal||'',inv.client||'',inv.desc||'',inv.invNum||'',parseFloat(inv.sub)||0,inv.category||'Consulting Revenue']]);ic.getRange(icr,10,1,2).setValues([['Invoice',sr]]);return respond({ok:true,invoiceRow:ir,incomeRow:icr})}

function handleRows(ss,rows,bank,source){if(!rows.length)return respond({ok:false,error:'No rows'});const c=readConfig(ss),es=ss.getSheetByName(EXPENSES_TAB),is=ss.getSheetByName(INCOME_TAB);if(!es)return respond({ok:false,error:'Tab not found: '+EXPENSES_TAB});if(!is)return respond({ok:false,error:'Tab not found: '+INCOME_TAB});const ic=new Set([...DEFAULT_INCOME_CATS,...(c.customIncomeCats||[])]),sk=new Set(DEFAULT_SKIP_CATS);const er=ler(es,12),ir=ler(is,11);const eR=[],iR=[],eM=[],iM=[];let s=0,d=0;for(const row of rows){const cat=row.category||'';if(sk.has(cat)){s++;continue}const ref=gsr(bank,row.date,row.amount||row.net,row.vendor);if(ic.has(cat)){if(ir.has(ref.toLowerCase())){d++;continue}ir.add(ref.toLowerCase());iR.push([row.date||'',row.vendor||'',row.description||'','',Math.abs(parseFloat(row.net)||parseFloat(row.amount)||0),cat]);iM.push([bank,ref])}else{if(er.has(ref.toLowerCase())){d++;continue}er.add(ref.toLowerCase());eR.push([row.date||'',row.vendor||'',row.description||'',Math.abs(parseFloat(row.amount)||0),cat]);eM.push([bank,bank==='AMEX'?'Yes':'No',source||'Import',ref])}}if(eR.length){const st=gnr(es);es.getRange(st,2,eR.length,5).setValues(eR);es.getRange(st,9,eR.length,4).setValues(eM)}if(iR.length){const st=gnr(is);is.getRange(st,2,iR.length,6).setValues(iR);is.getRange(st,10,iR.length,2).setValues(iM)}return respond({ok:true,expenses:eR.length,income:iR.length,skipped:s,duplicates:d,total:rows.length})}

function buildSummary(ss){const c=readConfig(ss),ed=gtd(ss,EXPENSES_TAB),id=gtd(ss,INCOME_TAB),iv=gtd(ss,INVOICES_TAB);const ec={};let te=0,hp=0;for(const r of ed){const cat=r[4]||'Uncategorized',a=parseFloat(r[3])||0,h=parseFloat(r[5])||0;ec[cat]=(ec[cat]||0)+a;te+=a;hp+=h}const ic={};let tr=0,hc=0;for(const r of id){const cat=r[5]||'Uncategorized',a=parseFloat(r[4])||0,h=parseFloat(r[6])||0;ic[cat]=(ic[cat]||0)+a;tr+=a;hc+=h}let up=0,ua=0;for(const r of iv){if(String(r[9]||'').toLowerCase()==='unpaid'){up++;ua+=parseFloat(r[6])||0}}const lg=ed.filter(r=>(parseFloat(r[3])||0)>=1000).map(r=>({date:fds(r[0]),vendor:r[1]||'',amount:rd2(parseFloat(r[3])||0),category:r[4]||''}));const used=new Set(Object.keys(ec)),miss=COMMON_EXPENSE_CATS.filter(c=>!used.has(c));const top=Object.entries(ec).sort((a,b)=>b[1]-a[1]).slice(0,10).map(([c,a])=>({category:c,amount:rd2(a)}));return{ok:true,businessName:c.businessName||'',province:c.province||'ON',fiscalYearEnd:c.fiscalYearEnd||'',totalRevenue:rd2(tr),totalExpenses:rd2(te),netIncome:rd2(tr-te),hstCollected:rd2(hc),hstPaid:rd2(hp),hstNetOwing:rd2(hc-hp),unpaidInvoices:up,unpaidTotal:rd2(ua),expensesByCategory:ec,incomeByCategory:ic,topExpenses:top,largeExpenses:lg,missingCategories:miss,transactionCount:ed.length+id.length,invoiceCount:iv.length,customExpenseCats:c.customExpenseCats||[],customIncomeCats:c.customIncomeCats||[]}}

function buildHealthCheck(ss){const ed=gtd(ss,EXPENSES_TAB),id=gtd(ss,INCOME_TAB),iv=gtd(ss,INVOICES_TAB);const issues=[];let score=100;const uc=ed.filter(r=>!r[4]||r[4]==='Other'||r[4]==='Uncategorized');if(uc.length){score-=Math.min(20,uc.length*2);issues.push({severity:'warning',message:uc.length+' expense(s) uncategorized or "Other".'})}const lg=ed.filter(r=>(parseFloat(r[3])||0)>=500&&r[4]!=='Equipment Purchase');if(lg.length){score-=Math.min(15,lg.length*3);issues.push({severity:'warning',message:lg.length+' expense(s) over $500 may be equipment (CCA).'})}const refs=ed.map(r=>r[10]||r[9]||'').filter(Boolean),rc={};refs.forEach(r=>{const k=String(r).toLowerCase();rc[k]=(rc[k]||0)+1});const dr=Object.values(rc).filter(c=>c>1).length;if(dr){score-=Math.min(20,dr*5);issues.push({severity:'error',message:dr+' potential duplicate(s).'})}const td=new Date(),ov=iv.filter(r=>{if(String(r[9]||'').toLowerCase()!=='unpaid')return false;const d=r[8]?new Date(r[8]):null;if(!d)return false;return(td-d)/(1000*60*60*24)>60});if(ov.length){score-=Math.min(10,ov.length*3);issues.push({severity:'warning',message:ov.length+' invoice(s) unpaid 60+ days.'})}const used=new Set(ed.map(r=>r[4]).filter(Boolean)),miss=COMMON_EXPENSE_CATS.filter(c=>!used.has(c));if(miss.length>=4){score-=Math.min(10,miss.length);issues.push({severity:'info',message:'No entries for: '+miss.join(', ')})}issues.sort((a,b)=>({error:0,warning:1,info:2})[a.severity]-({error:0,warning:1,info:2})[b.severity]);score=Math.max(0,Math.min(100,score));return{ok:true,score,issues,grade:score>=90?'A':score>=75?'B':score>=60?'C':score>=40?'D':'F',summary:issues.length===0?'Books look clean.':'Found '+issues.length+' item(s) to review.'}}

function ler(sheet,refCol){const lr=gnr(sheet)-1;if(lr<DATA_START_ROW)return new Set();const c=String.fromCharCode(64+refCol);return new Set(sheet.getRange(c+DATA_START_ROW+':'+c+lr).getValues().flat().filter(Boolean).map(r=>String(r).toLowerCase()))}
function gsr(bank,date,amount,vendor){const a=Math.abs(parseFloat(amount)||0).toFixed(2),v=String(vendor||'').substring(0,20).replace(/[^a-zA-Z0-9]/g,'').toUpperCase(),d=String(date||'').replace(/\\//g,'-');return bank+'-'+d+'-'+a+'-'+v}
function rd2(n){return Math.round((n||0)*100)/100}
function fds(d){if(!d)return'';if(d instanceof Date)return Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyy-MM-dd');return String(d)}
function gtd(ss,tab){const s=ss.getSheetByName(tab);if(!s)return[];const lr=gnr(s)-1;if(lr<DATA_START_ROW)return[];return s.getRange(DATA_START_ROW,2,lr-DATA_START_ROW+1,11).getValues()}
function gnr(sheet){const c=sheet.getRange('B'+DATA_START_ROW+':B500').getValues();for(let i=0;i<c.length;i++){if(!c[i][0])return DATA_START_ROW+i}return DATA_START_ROW+c.length}
function respond(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON)}
`;
}

function getAppsScriptManifest() {
  return JSON.stringify({
    timeZone: 'America/Toronto',
    dependencies: { enabledAdvancedServices: [] },
    exceptionLogging: 'STACKDRIVER',
    runtimeVersion: 'V8',
    oauthScopes: [
      'https://www.googleapis.com/auth/spreadsheets.currentonly',
      'https://www.googleapis.com/auth/drive.file',
      'https://www.googleapis.com/auth/script.container.ui',
    ],
    webapp: {
      executeAs: 'USER_DEPLOYING',
      access: 'ANYONE_ANONYMOUS',
    },
  });
}
