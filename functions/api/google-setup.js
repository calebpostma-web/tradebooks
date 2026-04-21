// functions/api/google-setup.js
// Handles Google OAuth token exchange and automatic sheet + script creation
// POST /api/google-setup with { code } or { action: 'create-sheet', token }

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

    // Step 1: Exchange auth code for tokens
    if (body.code) {
      return handleTokenExchange(body.code, env, headers);
    }

    // Step 2: Create sheet + deploy script using access token
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
async function handleTokenExchange(code, env, headers) {
  // ─── TEMPORARY DEBUG (remove after secret issue is resolved) ───
  const secret = env.GOOGLE_CLIENT_SECRET || '';
  const clientId = env.GOOGLE_CLIENT_ID || '';
  const debug = {
    clientIdLength: clientId.length,
    clientIdPrefix: clientId.slice(0, 20),
    clientIdEnd: clientId.slice(-20),
    secretLength: secret.length,
    secretPrefix: secret.slice(0, 7),
    secretEnd: secret.slice(-4),
    secretHasWhitespace: /\s/.test(secret),
    secretHasNonAscii: /[^\x20-\x7E]/.test(secret),
    codeLength: (code || '').length,
    codePrefix: (code || '').slice(0, 10),
  };
  console.log('[handleTokenExchange] DEBUG:', JSON.stringify(debug));

  let tokenResp, tokenRespStatus, tokenRespHeaders, tokenRaw, tokenData;
  try {
    tokenResp = await fetch('https://oauth2.googleapis.com/token', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        code,
        client_id: clientId,
        client_secret: secret,
        redirect_uri: 'https://tradebooks-bju.pages.dev/app/',
        grant_type: 'authorization_code',
      }),
    });
    tokenRespStatus = tokenResp.status;
    tokenRespHeaders = Object.fromEntries(tokenResp.headers.entries());
    tokenRaw = await tokenResp.text();
    try { tokenData = JSON.parse(tokenRaw); } catch(_) { tokenData = { error: 'parse_failed', raw: tokenRaw }; }
  } catch(fetchErr) {
    return new Response(JSON.stringify({
      ok: false,
      error: 'Network error reaching Google: ' + fetchErr.message,
      debug,
    }), { headers });
  }

  console.log('[handleTokenExchange] Google status:', tokenRespStatus);
  console.log('[handleTokenExchange] Google response:', tokenRaw);

  if (tokenData.error) {
    return new Response(JSON.stringify({
      ok: false,
      error: tokenData.error_description || tokenData.error,
      debug,
      googleStatus: tokenRespStatus,
      googleResponse: tokenData,
      googleResponseHeaders: tokenRespHeaders,
    }), { headers });
  }

  // Get user info
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
// STEP 2: Create Google Sheet + inject Apps Script + deploy
// This is the one-click magic
// ════════════════════════════════════════════════════════════════════
async function handleCreateSheet(accessToken, profile, env, headers) {
  const authHeader = { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' };

  // ── 1. Create the Google Sheet ──
  const sheetResp = await fetch('https://sheets.googleapis.com/v4/spreadsheets', {
    method: 'POST',
    headers: authHeader,
    body: JSON.stringify({
      properties: {
        title: `${profile.businessName || 'My Business'} — TradeBooks`,
      },
      sheets: [
        { properties: { title: '📊 Dashboard', index: 0 } },
        { properties: { title: '💰 Income', index: 1 } },
        { properties: { title: '💸 Expenses', index: 2 } },
        { properties: { title: '🧾 Invoices', index: 3 } },
        { properties: { title: '⚙️ Config', index: 4 } },
      ],
    }),
  });

  const sheetData = await sheetResp.json();
  if (sheetData.error) {
    return new Response(JSON.stringify({ ok: false, error: 'Sheet creation failed: ' + (sheetData.error.message || JSON.stringify(sheetData.error)) }), { headers });
  }

  const spreadsheetId = sheetData.spreadsheetId;
  const spreadsheetUrl = sheetData.spreadsheetUrl;

  // ── 2. Populate headers and formulas via batch update ──
  await populateSheet(accessToken, spreadsheetId, profile);

  // ── 3. Create Apps Script project bound to the sheet ──
  const scriptResp = await fetch('https://script.googleapis.com/v1/projects', {
    method: 'POST',
    headers: authHeader,
    body: JSON.stringify({
      title: 'TradeBooks Add-on',
      parentId: spreadsheetId,
    }),
  });

  const scriptData = await scriptResp.json();
  if (scriptData.error) {
    // Script creation failed — sheet still works, just no web app endpoint
    return new Response(JSON.stringify({
      ok: true,
      spreadsheetId,
      spreadsheetUrl,
      scriptError: scriptData.error.message || 'Could not create Apps Script',
      scriptUrl: null,
    }), { headers });
  }

  const scriptId = scriptData.scriptId;

  // ── 4. Push the Apps Script code ──
  const codeContent = getAppsScriptCode();
  const manifestContent = getAppsScriptManifest();

  const updateResp = await fetch(`https://script.googleapis.com/v1/projects/${scriptId}/content`, {
    method: 'PUT',
    headers: authHeader,
    body: JSON.stringify({
      files: [
        {
          name: 'Code',
          type: 'SERVER_JS',
          source: codeContent,
        },
        {
          name: 'appsscript',
          type: 'JSON',
          source: manifestContent,
        },
      ],
    }),
  });

  const updateData = await updateResp.json();

  // ── 5. Deploy as web app ──
  let scriptUrl = null;
  try {
    const deployResp = await fetch(`https://script.googleapis.com/v1/projects/${scriptId}/deployments`, {
      method: 'POST',
      headers: authHeader,
      body: JSON.stringify({
        versionNumber: 1,
        manifestFileName: 'appsscript',
        description: 'TradeBooks Web App',
      }),
    });

    // First we need to create a version
    const versionResp = await fetch(`https://script.googleapis.com/v1/projects/${scriptId}/versions`, {
      method: 'POST',
      headers: authHeader,
      body: JSON.stringify({ description: 'TradeBooks v1' }),
    });
    const versionData = await versionResp.json();

    if (versionData.versionNumber) {
      const deployResp2 = await fetch(`https://script.googleapis.com/v1/projects/${scriptId}/deployments`, {
        method: 'POST',
        headers: authHeader,
        body: JSON.stringify({
          versionNumber: versionData.versionNumber,
          description: 'TradeBooks Web App',
          entryPoints: [{
            entryPointType: 'WEB_APP',
            webApp: {
              access: 'ANYONE_ANONYMOUS',
              executeAs: 'USER_DEPLOYING',
            },
          }],
        }),
      });
      const deployData = await deployResp2.json();
      if (deployData.entryPoints) {
        const webApp = deployData.entryPoints.find(e => e.entryPointType === 'WEB_APP');
        if (webApp) scriptUrl = webApp.webApp.url;
      }
    }
  } catch (deployErr) {
    // Deployment failed — not critical, user can deploy manually
  }

  // ── 6. Write the script URL back to the Config tab ──
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
    ok: true,
    spreadsheetId,
    spreadsheetUrl,
    scriptId,
    scriptUrl,
  }), { headers });
}


// ════════════════════════════════════════════════════════════════════
// POPULATE SHEET — writes headers, formulas, config to all tabs
// ════════════════════════════════════════════════════════════════════
async function populateSheet(accessToken, spreadsheetId, profile) {
  const authHeader = { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' };

  const prov = profile.province || 'ON';
  const taxRate = {
    ON:0.13, BC:0.12, AB:0.05, SK:0.11, MB:0.12,
    QC:0.14975, NB:0.15, NS:0.15, PE:0.15, NL:0.15
  }[prov] || 0.13;

  // Batch update all tabs at once
  const data = [
    // ── Dashboard ──
    { range: "'📊 Dashboard'!B3", values: [['PROFIT & LOSS SUMMARY']] },
    { range: "'📊 Dashboard'!B5:B10", values: [
      ['Total Revenue (excl HST)'],['HST Collected'],
      ['Total Expenses (excl HST)'],['HST Paid — ITCs'],
      ['NET INCOME before tax'],['HST OWING / (Refund)']
    ]},
    { range: "'📊 Dashboard'!D5:D10", values: [
      ['=SUMPRODUCT((\'💰 Income\'!F12:F500<>"")*1*(\'💰 Income\'!F12:F500))'],
      ['=SUMPRODUCT((\'💰 Income\'!H12:H500<>"")*1*(\'💰 Income\'!H12:H500))'],
      ['=SUMPRODUCT((\'💸 Expenses\'!E12:E500<>"")*1*(\'💸 Expenses\'!E12:E500))'],
      ['=SUMPRODUCT((\'💸 Expenses\'!G12:G500<>"")*1*(\'💸 Expenses\'!G12:G500))'],
      ['=D5-D7'],
      ['=D6-D8']
    ]},

    // ── Expenses headers ──
    { range: "'💸 Expenses'!B11:L11", values: [[
      'Date','Vendor','Description','Amount (excl HST)','Category',
      'HST (auto)','Total (auto)','Payment Method','On AMEX?','Source','Source Ref'
    ]]},
    // HST formula (column G)
    { range: "'💸 Expenses'!G12", values: [[
      `=ARRAYFORMULA(IF(E12:E500="","",IF(REGEXMATCH(F12:F500,"Wages|Insurance|Interest & Bank|Property Tax|Bill Payment|Owner Draw|Income Received|SKIP"),0,E12:E500*${taxRate})))`
    ]]},
    // Total formula (column H)
    { range: "'💸 Expenses'!H12", values: [[
      '=ARRAYFORMULA(IF(E12:E500="","",E12:E500+G12:G500))'
    ]]},

    // ── Income headers ──
    { range: "'💰 Income'!B11:K11", values: [[
      'Date','Client/Source','Description','Invoice #','Amount (excl HST)',
      'Category','HST (auto)','Total (auto)','Payment Method','Source Ref'
    ]]},
    // HST formula (column H)
    { range: "'💰 Income'!H12", values: [[
      `=ARRAYFORMULA(IF(F12:F500="","",F12:F500*${taxRate}))`
    ]]},
    // Total formula (column I)
    { range: "'💰 Income'!I12", values: [[
      '=ARRAYFORMULA(IF(F12:F500="","",F12:F500+H12:H500))'
    ]]},

    // ── Invoices headers ──
    { range: "'🧾 Invoices'!B11:K11", values: [[
      'Invoice #','Date','Client','Description','Subtotal','HST','Total','HST?','Due Date','Status'
    ]]},

    // ── Config tab ──
    { range: "'⚙️ Config'!A1", values: [['⚙️ TRADEBOOKS CONFIGURATION']] },
    { range: "'⚙️ Config'!A2", values: [['Edit the yellow cells below. The app reads these settings automatically.']] },
    { range: "'⚙️ Config'!A4", values: [['  BUSINESS INFORMATION']] },
    { range: "'⚙️ Config'!B5:C13", values: [
      ['Business Name', profile.businessName || ''],
      ['Trading / DBA Name', profile.tradingName || ''],
      ['Owner Name', profile.ownerName || ''],
      ['City', profile.city || ''],
      ['Province', prov],
      ['HST / GST Rate', taxRate],
      ['HST / GST Number', profile.hstNumber || ''],
      ['Business Type', profile.businessType || 'sole_prop'],
      ['Fiscal Year End', profile.fiscalYearEnd || 'December 31'],
    ]},
    { range: "'⚙️ Config'!E5:F13", values: [
      ['Primary Bank', profile.primaryBank || 'BMO'],
      ['Credit Card', profile.creditCard || 'AMEX'],
      ['Starting Invoice #', profile.invoiceStart || 1001],
      ['Home Office %', profile.homeOfficePercent || 0],
      ['Email', profile.email || ''],
      ['Apps Script URL', ''],
      ['Corporate Structure', profile.structure || ''],
      ['Business Activities', profile.activities || ''],
      ['',''],
    ]},

    // Custom categories section
    { range: "'⚙️ Config'!A15", values: [['  CUSTOM EXPENSE CATEGORIES — add your own below (one per row)']] },
    { range: "'⚙️ Config'!B16", values: [['Default Categories (do not edit)']] },
    { range: "'⚙️ Config'!E16", values: [['Your Custom Categories (edit these)']] },

    // Default expense categories list
    { range: "'⚙️ Config'!B17:B48", values: [
      ['Meals & Entertainment'],['Professional Fees'],['Wages & Salaries'],['Small Tools'],
      ['Supplies'],['Uniforms'],['Dues & Memberships'],['Vehicle Repairs'],['Fuel'],
      ['Insurance'],['Interest & Bank Charges'],['Repairs & Maintenance'],['Property Tax'],
      ['Telephone'],['Utilities'],['Conferences'],['Equipment Purchase'],
      ['Inventory — Materials (COGS)'],['Advertising & Marketing'],['Subcontractors'],
      ['Office Supplies'],['Rent'],['Home Office'],['Vehicle Lease/Payments'],['Travel'],
      ['Training & Education'],['Permits & Licenses'],['Tax Payments'],['Other'],
      ['Bill Payment / Transfer'],['Owner Draw / Distribution'],['SKIP — not a business expense'],
    ]},

    // Custom income categories section
    { range: "'⚙️ Config'!A50", values: [['  CUSTOM INCOME CATEGORIES — add your own below (one per row)']] },
    { range: "'⚙️ Config'!B51", values: [['Default Categories (do not edit)']] },
    { range: "'⚙️ Config'!E51", values: [['Your Custom Categories (edit these)']] },
    { range: "'⚙️ Config'!B52:B59", values: [
      ['Consulting Revenue'],['Service Revenue'],['Sales Revenue'],
      ['Sales Revenue — Materials'],['Rental Income'],['Interest Income'],
      ['Income Received'],['Other Income'],
    ]},

    // Clients section
    { range: "'⚙️ Config'!A62", values: [['  CLIENTS — for invoice autocomplete (one per row)']] },
    { range: "'⚙️ Config'!B63:E63", values: [['Client Name','','','Default Income Category']] },

    // Employees section
    { range: "'⚙️ Config'!A75", values: [['  EMPLOYEES — for payroll and T4 generation (one per row)']] },
    { range: "'⚙️ Config'!B76:F76", values: [['Employee Name','SIN (last 3)','','Annual Salary/Wages','Status']] },
  ];

  await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values:batchUpdate`, {
    method: 'POST',
    headers: authHeader,
    body: JSON.stringify({
      valueInputOption: 'USER_ENTERED',
      data,
    }),
  });
}


// ════════════════════════════════════════════════════════════════════
// APPS SCRIPT CODE — the full script injected into the user's sheet
// This is the same Code.gs from the add-on, minified for transport
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

function doGet(e){try{const ss=SpreadsheetApp.getActiveSpreadsheet(),a=(e&&e.parameter&&e.parameter.action)||'';if(a==='summary')return respond(buildSummary(ss));if(a==='health')return respond(buildHealthCheck(ss));if(a==='config')return respond({ok:true,config:readConfig(ss)});if(a==='categories')return respond(buildCategoryLists(ss));const c=readConfig(ss);return respond({ok:true,message:'TradeBooks v4 ✓',business:c.businessName||'Not configured',province:c.province||'ON'})}catch(e){return respond({ok:false,error:e.message})}}

function doPost(e){try{const d=JSON.parse(e.postData.contents),ss=SpreadsheetApp.getActiveSpreadsheet();if(d.type==='invoice')return handleInvoice(ss,d.invoice);return handleRows(ss,d.rows||[],d.bank||'AMEX',d.source||'Import')}catch(e){return respond({ok:false,error:e.message})}}

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
    dependencies: {},
    exceptionLogging: 'STACKDRIVER',
    runtimeVersion: 'V8',
    webapp: {
      executeAs: 'USER_DEPLOYING',
      access: 'ANYONE_ANONYMOUS',
    },
  });
}
