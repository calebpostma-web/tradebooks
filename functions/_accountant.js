// ════════════════════════════════════════════════════════════════════
// ACCOUNTANT VIEW — formula-driven mirror of the bookkeeper's own layout
//
// Builds three tabs, all computed from 📒 Transactions (nobody types here):
//   🧮 Accountant – <bank>   one row per bank transaction
//   🧮 Accountant – <card>   one row per credit-card transaction
//   🧮 HST Report            Sales / HST Collected / HST Paid / Payable
//
// Column layout of an account tab (row 3 headers, row 4 totals, data from row 6):
//   Date | Name | Total | Receipt | HST on sales | <one col per revenue category>
//   | HST paid on expenses | <one col per expense category> | (no category)
//   | <one col per "transfer column"> | hidden helper columns
//
// Category columns come from the user's custom category lists (their own
// wording) and fall back to the app defaults. "Transfer columns" are the
// money-between-pockets rows (card payments, CRA, shareholder loan): they are
// Internal Transfer in the ledger and are shown here under the names the user
// chose, matched by payee. They are configured in ⚙️ Config (see TRANSFER_CFG).
//
// Idempotent: header row 3 is compared with the expected headers; the tab is
// (re)built only when they differ, e.g. after a category is added.
// ════════════════════════════════════════════════════════════════════

import { readRange, batchUpdate, spreadsheetsBatchUpdate } from './_sheets.js';

const DEFAULT_EXP_CATS = [
  'Meals & Entertainment', 'Professional Fees', 'Wages & Salaries', 'Small Tools',
  'Supplies', 'Uniforms', 'Dues & Memberships', 'Vehicle Repairs', 'Fuel',
  'Insurance', 'Interest & Bank Charges', 'Repairs & Maintenance', 'Property Tax',
  'Telephone', 'Utilities', 'Conferences', 'Equipment Purchase',
  'Inventory — Materials (COGS)', 'Advertising & Marketing', 'Subcontractors',
  'Office Supplies', 'Rent', 'Home Office', 'Vehicle Lease/Payments', 'Travel',
  'Training & Education', 'Permits & Licenses', 'Tax Payments', 'Other',
];
const DEFAULT_INC_CATS = [
  'Consulting Revenue', 'Service Revenue', 'Sales Revenue', 'Sales Revenue — Materials',
  'Rental Income', 'Interest Income', 'Income Received', 'Other Income',
];

// Transfer-column config block in ⚙️ Config. Header at A100, rows 101–110:
//   B = column name shown in the view, C = text to match in the payee (regex, case-insensitive)
const TRANSFER_CFG = { headerCell: 'A100', range: 'B101:C110', firstRow: 101 };
const DEFAULT_TRANSFERS = [
  ['Credit card payment', 'AMEX|VISA|MASTERCARD|CARD PAYMENT'],
  ['CRA payments',        'CRA|CANADA|RECEIVER GENERAL'],
  ['Shareholder / owner', 'SHAREHOLDER|OWNER|LOAN'],
];

const COLORS = {
  teal:     { red: 0.078, green: 0.278, blue: 0.247 },
  tealTint: { red: 0.949, green: 0.980, blue: 0.976 },
  white:    { red: 1, green: 1, blue: 1 },
  grey:     { red: 0.93, green: 0.93, blue: 0.93 },
};
const FMT_CURRENCY = { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00;("$"#,##0.00)' } };
const FMT_DATE     = { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } };

const LAYOUT_VERSION = 3;   // bump when row-6 formulas change so existing tabs get rebuilt
const HELPER_COUNT = 8;   // Date, Party, Total(incl HST), Category, Amount(excl), HST, SourceRef, Receipt link
const FIRST_DATA_ROW = 6;

export function colLetter(idx) {              // 0 → A, 25 → Z, 26 → AA
  let n = idx + 1, out = '';
  while (n > 0) { const r = (n - 1) % 26; out = String.fromCharCode(65 + r) + out; n = Math.floor((n - 1) / 26); }
  return out;
}
const q = t => `'${String(t).replace(/'/g, "''")}'`;
const esc = s => String(s).replace(/"/g, '""');

function findTab(sheetsByTitle, re) { return Object.values(sheetsByTitle).find(s => re.test(s.title)); }

function cleanList(list, fallback) {
  const seen = new Set();
  const out = (Array.isArray(list) ? list : []).map(c => String(c || '').trim()).filter(c => c && !seen.has(c) && seen.add(c));
  return out.length ? out : fallback;
}

// ── Read (or seed) the transfer-column config from ⚙️ Config ──
async function loadTransferColumns(env, userId, cfgTab, dryRun, changes, errors) {
  const cfgTitle = cfgTab.title;
  const hdr = await readRange(env, userId, `${q(cfgTitle)}!${TRANSFER_CFG.headerCell}`);
  const present = hdr.ok && hdr.values && hdr.values[0] && String(hdr.values[0][0] || '').includes('TRANSFER COLUMNS');
  if (!present) {
    if (dryRun) {
      changes.push(`Add a TRANSFER COLUMNS section to '${cfgTitle}' (names for card payments, CRA and shareholder rows in the Accountant View — edit them to your wording).`);
      return DEFAULT_TRANSFERS;
    }
    // ⚙️ Config is created with 100 rows; the block lives at 100–110.
    const needRows = TRANSFER_CFG.firstRow + 10;
    if ((cfgTab.gridProperties?.rowCount || 0) < needRows) {
      const grow = await spreadsheetsBatchUpdate(env, userId, [{ updateSheetProperties: {
        properties: { sheetId: cfgTab.sheetId, gridProperties: { rowCount: needRows } }, fields: 'gridProperties.rowCount' } }]);
      if (!grow.ok) { errors.push(`Could not grow '${cfgTitle}' for TRANSFER COLUMNS: ${grow.error}`); return DEFAULT_TRANSFERS; }
      cfgTab.gridProperties = { ...(cfgTab.gridProperties || {}), rowCount: needRows };
    }
    const res = await batchUpdate(env, userId, [
      { range: `${q(cfgTitle)}!${TRANSFER_CFG.headerCell}`, values: [['  TRANSFER COLUMNS — Accountant View column names for money moved between accounts (not expenses). Column C = words to match in the payee.']] },
      { range: `${q(cfgTitle)}!B${TRANSFER_CFG.firstRow}:C${TRANSFER_CFG.firstRow + DEFAULT_TRANSFERS.length - 1}`, values: DEFAULT_TRANSFERS },
    ]);
    if (!res.ok) { errors.push(`Could not write TRANSFER COLUMNS to '${cfgTitle}': ${res.error}`); }
    else changes.push(`Added TRANSFER COLUMNS section to '${cfgTitle}' — edit the names to your own wording.`);
    return DEFAULT_TRANSFERS;
  }
  const rows = await readRange(env, userId, `${q(cfgTitle)}!${TRANSFER_CFG.range}`);
  const list = (rows.ok ? rows.values : []).map(r => [String(r[0] || '').trim(), String(r[1] || '').trim()]).filter(r => r[0] && r[1]);
  return list.length ? list : DEFAULT_TRANSFERS;
}

// ── Expected header row for an account tab ──
function buildHeaders(revCats, expCats, transfers) {
  return ['Date', 'Name', 'Total', 'Receipt', 'HST on sales', ...revCats, 'HST paid on expenses', ...expCats, '(no category)', ...transfers.map(t => t[0])];
}

// ── Formulas for one account tab ──
// `sign` is +1 for a bank account (money in positive), −1 for a credit card
// (charges shown positive, payments negative — the way a card statement reads).
function buildAccountTab({ title, sheetId, account, sign, txnTitle, revCats, expCats, transfers, headers }) {
  const nVisible = headers.length;                  // headers start at column B (index 1)
  const helperStart = 1 + nVisible + 1;             // one blank column, then helpers
  const H = i => colLetter(helperStart + i);        // helper column letters
  const [hDate, hParty, hTotal, hCat, hAmt, hHst, hRef, hRcpt] = [0, 1, 2, 3, 4, 5, 6, 7].map(H);
  const rng = c => `${c}${FIRST_DATA_ROW}:${c}`;
  const T = q(txnTitle);
  const values = [];

  values.push({ range: `${q(title)}!A1`, values: [[`ACCOUNTANT VIEW — ${account}  ·  Built automatically from ${txnTitle}  ·  Do not type here — edit the ledger instead  ·  v${LAYOUT_VERSION}`]] });
  values.push({ range: `${q(title)}!B3`, values: [headers] });

  // Helper block: one FILTER, sorted by date, spills 8 columns
  values.push({ range: `${q(title)}!${hDate}5`, values: [['Date', 'Party', 'Total incl HST', 'Category', 'Amount excl HST', 'HST', 'Source Ref', 'Receipt link']] });
  values.push({ range: `${q(title)}!${hDate}${FIRST_DATA_ROW}`, values: [[
    `=IFERROR(SORT(FILTER({${T}!B12:B,${T}!C12:C,${T}!N12:N,${T}!F12:F,${T}!E12:E,${T}!H12:H,${T}!K12:K,${T}!O12:O},${T}!I12:I="${esc(account)}",${T}!B12:B<>""),1,TRUE),"")`
  ]] });

  const guard = expr => `=ARRAYFORMULA(IF(LEN(${rng(hDate)})=0,"",${expr}))`;
  const row6 = [];      // formulas for row 6, columns B..
  const row4 = [];      // totals for row 4
  headers.forEach((h, i) => {
    const col = colLetter(1 + i);
    let f, tot = `=SUM(${rng(col)})`;
    if (i === 0)      { f = guard(rng(hDate)); tot = ''; }
    else if (i === 1) { f = guard(rng(hParty)); tot = ''; }
    else if (i === 2) { f = guard(`${sign}*(${rng(hAmt)}+${rng(hHst)}*SIGN(${rng(hAmt)}))`); }
    else if (i === 3) { f = guard(`IF(LEN(${rng(hRcpt)})>0,HYPERLINK(${rng(hRcpt)},"📎 view"),"")`); tot = `=COUNTIF(${rng(col)},"📎*")`; }
    else if (h === 'HST on sales')        { f = guard(`IF(${rng(hAmt)}>0,${rng(hHst)},"")`); }
    else if (h === 'HST paid on expenses'){ f = guard(`IF(${rng(hAmt)}<0,${rng(hHst)},"")`); }
    else if (h === '(no category)')       { f = guard(`IF((${rng(hCat)}="")*(${rng(hAmt)}<>""),ABS(${rng(hAmt)}),"")`); }
    else if (revCats.includes(h))         { f = guard(`IF((${rng(hCat)}="${esc(h)}")*(${rng(hAmt)}>0),${rng(hAmt)},"")`); }
    else if (expCats.includes(h))         { f = guard(`IF((${rng(hCat)}="${esc(h)}")*(${rng(hAmt)}<0),-${rng(hAmt)},"")`); }
    else {
      const t = transfers.find(x => x[0] === h);
      f = guard(`IF((${rng(hCat)}="Internal Transfer")*REGEXMATCH(UPPER(${rng(hParty)}),"${esc(t ? t[1] : h)}"),ABS(${rng(hAmt)})+ABS(${rng(hHst)}),"")`);
    }
    row6.push(f); row4.push(tot);
  });
  values.push({ range: `${q(title)}!B${FIRST_DATA_ROW}`, values: [row6] });
  values.push({ range: `${q(title)}!A4`, values: [['TOTALS', ...row4]] });
  values.push({ range: `${q(title)}!A5`, values: [['↓ rows']] });

  // Styling
  const lastCol = helperStart + HELPER_COUNT;
  const reqs = [
    { repeatCell: { range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: lastCol },
        cell: { userEnteredFormat: { backgroundColor: COLORS.teal, textFormat: { foregroundColor: COLORS.white, bold: true, fontSize: 11 } } }, fields: 'userEnteredFormat' } },
    { repeatCell: { range: { sheetId, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 1, endColumnIndex: 1 + nVisible },
        cell: { userEnteredFormat: { backgroundColor: COLORS.teal, textFormat: { foregroundColor: COLORS.white, bold: true, fontSize: 9 }, wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat' } },
    { repeatCell: { range: { sheetId, startRowIndex: 3, endRowIndex: 4, startColumnIndex: 0, endColumnIndex: 1 + nVisible },
        cell: { userEnteredFormat: { backgroundColor: COLORS.tealTint, textFormat: { bold: true } , numberFormat: FMT_CURRENCY.numberFormat } }, fields: 'userEnteredFormat' } },
    { repeatCell: { range: { sheetId, startRowIndex: FIRST_DATA_ROW - 1, endRowIndex: 5000, startColumnIndex: 1, endColumnIndex: 2 },
        cell: { userEnteredFormat: FMT_DATE }, fields: 'userEnteredFormat.numberFormat' } },
    { repeatCell: { range: { sheetId, startRowIndex: FIRST_DATA_ROW - 1, endRowIndex: 5000, startColumnIndex: 3, endColumnIndex: 1 + nVisible },
        cell: { userEnteredFormat: FMT_CURRENCY }, fields: 'userEnteredFormat.numberFormat' } },
    { repeatCell: { range: { sheetId, startRowIndex: 4, endRowIndex: 5000, startColumnIndex: helperStart, endColumnIndex: lastCol },
        cell: { userEnteredFormat: { backgroundColor: COLORS.grey } }, fields: 'userEnteredFormat.backgroundColor' } },
    { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: helperStart, endIndex: lastCol }, properties: { hiddenByUser: true }, fields: 'hiddenByUser' } },
    { updateDimensionProperties: { range: { sheetId, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 64 }, fields: 'pixelSize' } },
    { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 260 }, fields: 'pixelSize' } },
    { updateSheetProperties: { properties: { sheetId, gridProperties: { frozenRowCount: 5, frozenColumnCount: 4 } }, fields: 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount' } },
  ];
  return { values, reqs, columnCount: lastCol + 1 };
}

// ── HST Report tab ──
function buildHstReport({ title, sheetId, txnTitle, hstTitle }) {
  const T = q(txnTitle);
  const fy = `${q(hstTitle)}!C3`;
  const values = [
    { range: `${q(title)}!A1`, values: [['HST REPORT  ·  pick a period in C3  ·  figures come from the ledger']] },
    { range: `${q(title)}!B3:C3`, values: [['Period', 'Full year']] },
    { range: `${q(title)}!B4:C6`, values: [
      ['Fiscal year start', `=${fy}`],
      ['From', `=IF(C3="Full year",C4,EDATE(C4,3*(MATCH(LEFT(C3,2),{"Q1","Q2","Q3","Q4"},0)-1)))`],
      ['To',   `=IF(C3="Full year",EDATE(C4,12)-1,EDATE(C5,3)-1)`],
    ]},
    { range: `${q(title)}!B8:C11`, values: [
      ['Sales',            `=SUMIFS(${T}!E12:E,${T}!E12:E,">0",${T}!F12:F,"<>Internal Transfer",${T}!B12:B,">="&C5,${T}!B12:B,"<="&C6)`],
      ['HST Collected',    `=SUMIFS(${T}!H12:H,${T}!E12:E,">0",${T}!F12:F,"<>Internal Transfer",${T}!B12:B,">="&C5,${T}!B12:B,"<="&C6)`],
      ['HST Paid',         `=SUMIFS(${T}!H12:H,${T}!E12:E,"<0",${T}!F12:F,"<>Internal Transfer",${T}!B12:B,">="&C5,${T}!B12:B,"<="&C6)`],
      ['(Refund)/Payable', `=C9-C10`],
    ]},
  ];
  const reqs = [
    { repeatCell: { range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 6 },
        cell: { userEnteredFormat: { backgroundColor: COLORS.teal, textFormat: { foregroundColor: COLORS.white, bold: true, fontSize: 11 } } }, fields: 'userEnteredFormat' } },
    { setDataValidation: { range: { sheetId, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 2, endColumnIndex: 3 },
        rule: { condition: { type: 'ONE_OF_LIST', values: ['Q1 Apr–Jun', 'Q2 Jul–Sep', 'Q3 Oct–Dec', 'Q4 Jan–Mar', 'Full year'].map(v => ({ userEnteredValue: v })) }, showCustomUi: true, strict: true } } },
    { repeatCell: { range: { sheetId, startRowIndex: 3, endRowIndex: 6, startColumnIndex: 2, endColumnIndex: 3 }, cell: { userEnteredFormat: FMT_DATE }, fields: 'userEnteredFormat.numberFormat' } },
    { repeatCell: { range: { sheetId, startRowIndex: 7, endRowIndex: 11, startColumnIndex: 2, endColumnIndex: 3 }, cell: { userEnteredFormat: FMT_CURRENCY }, fields: 'userEnteredFormat.numberFormat' } },
    { repeatCell: { range: { sheetId, startRowIndex: 7, endRowIndex: 11, startColumnIndex: 1, endColumnIndex: 2 }, cell: { userEnteredFormat: { textFormat: { bold: true } } }, fields: 'userEnteredFormat.textFormat' } },
    { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 1, endIndex: 3 }, properties: { pixelSize: 170 }, fields: 'pixelSize' } },
  ];
  return { values, reqs };
}

// ── Entry point (used by Migration 11 and by new-sheet setup) ──
// profile: { primaryBank, creditCard, customExpenseCats, customIncomeCats }
export async function ensureAccountantTabs(env, userId, { sheetsByTitle, profile, dryRun, changes, errors }) {
  const txnTab = findTab(sheetsByTitle, /transactions/i);
  const cfgTab = findTab(sheetsByTitle, /config/i);
  const hstTab = findTab(sheetsByTitle, /hst returns/i);
  if (!txnTab || !cfgTab || !hstTab) return;   // partial sheet — nothing to build on

  const bank = String(profile?.primaryBank || 'BMO').trim();
  const card = String(profile?.creditCard || 'AMEX').trim();
  const revCats = cleanList(profile?.customIncomeCats, DEFAULT_INC_CATS);
  const expCats = cleanList(profile?.customExpenseCats, DEFAULT_EXP_CATS).filter(c => !/internal transfer|owner draw|skip/i.test(c));
  const transfers = await loadTransferColumns(env, userId, cfgTab, dryRun, changes, errors);
  const headers = buildHeaders(revCats, expCats, transfers);

  const usedIds = new Set(Object.values(sheetsByTitle).map(s => s.sheetId));
  let nextId = 1500; const freeId = () => { while (usedIds.has(nextId)) nextId += 1; usedIds.add(nextId); return nextId; };

  const specs = [
    { title: `🧮 Accountant – ${bank}`, account: bank, sign: 1 },
    { title: `🧮 Accountant – ${card}`, account: card, sign: -1 },
  ];

  for (const spec of specs) {
    const existing = sheetsByTitle[spec.title];
    if (existing) {
      const hdr = await readRange(env, userId, `${q(spec.title)}!B3:${colLetter(headers.length)}3`);
      const current = (hdr.ok && hdr.values && hdr.values[0]) ? hdr.values[0].map(v => String(v)) : [];
      const banner = await readRange(env, userId, `${q(spec.title)}!A1`);
      const sameVersion = banner.ok && banner.values && banner.values[0] && String(banner.values[0][0] || '').endsWith(`v${LAYOUT_VERSION}`);
      if (sameVersion && current.length === headers.length && current.every((v, i) => v === headers[i])) continue;   // up to date
    }
    if (dryRun) {
      changes.push(existing
        ? `Rebuild '${spec.title}' — your category list changed or the layout was updated.`
        : `Add '${spec.title}' — your bookkeeping layout (one column per category), built automatically from the ledger.`);
      continue;
    }
    let sheetId = existing?.sheetId;
    const columnCount = 1 + headers.length + 1 + HELPER_COUNT + 1;
    if (!existing) {
      sheetId = freeId();
      const add = await spreadsheetsBatchUpdate(env, userId, [{ addSheet: { properties: { sheetId, title: spec.title, index: Object.keys(sheetsByTitle).length,
        gridProperties: { rowCount: 5000, columnCount, frozenRowCount: 5, frozenColumnCount: 4 }, tabColor: COLORS.teal } } }]);
      if (!add.ok) { errors.push(`Could not create '${spec.title}': ${add.error}`); continue; }
      sheetsByTitle[spec.title] = { sheetId, title: spec.title, gridProperties: { rowCount: 5000, columnCount } };
    } else {
      // Clear old layout (rows 1–6, all columns) and make sure the grid is wide enough
      const reqs = [{ updateCells: { range: { sheetId, startRowIndex: 0, endRowIndex: 6 }, fields: 'userEnteredValue,userEnteredFormat' }}];
      if ((existing.gridProperties?.columnCount || 0) < columnCount) {
        reqs.unshift({ updateSheetProperties: { properties: { sheetId, gridProperties: { columnCount } }, fields: 'gridProperties.columnCount' } });
      }
      // Unhide everything first; helper columns get re-hidden below
      reqs.push({ updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 0, endIndex: columnCount }, properties: { hiddenByUser: false }, fields: 'hiddenByUser' } });
      const clr = await spreadsheetsBatchUpdate(env, userId, reqs);
      if (!clr.ok) { errors.push(`Could not reset '${spec.title}': ${clr.error}`); continue; }
    }
    const built = buildAccountTab({ ...spec, sheetId, txnTitle: txnTab.title, revCats, expCats, transfers, headers });
    const w = await batchUpdate(env, userId, built.values);
    if (!w.ok) { errors.push(`Could not write '${spec.title}': ${w.error}`); continue; }
    const st = await spreadsheetsBatchUpdate(env, userId, built.reqs);
    if (!st.ok) errors.push(`'${spec.title}' written but styling failed: ${st.error}`);
    changes.push(existing ? `Rebuilt '${spec.title}' with your current categories.` : `Added '${spec.title}' — your layout, ${headers.length} columns, filled from the ledger automatically.`);
  }

  // HST Report
  const hstReportTitle = '🧮 HST Report';
  if (!sheetsByTitle[hstReportTitle]) {
    if (dryRun) { changes.push(`Add '${hstReportTitle}' — Sales / HST Collected / HST Paid / Payable for any quarter or the full year.`); return; }
    const sheetId = freeId();
    const add = await spreadsheetsBatchUpdate(env, userId, [{ addSheet: { properties: { sheetId, title: hstReportTitle, index: Object.keys(sheetsByTitle).length,
      gridProperties: { rowCount: 20, columnCount: 6 }, tabColor: COLORS.teal } } }]);
    if (!add.ok) { errors.push(`Could not create '${hstReportTitle}': ${add.error}`); return; }
    sheetsByTitle[hstReportTitle] = { sheetId, title: hstReportTitle };
    const built = buildHstReport({ title: hstReportTitle, sheetId, txnTitle: txnTab.title, hstTitle: hstTab.title });
    const w = await batchUpdate(env, userId, built.values);
    if (!w.ok) { errors.push(`Could not write '${hstReportTitle}': ${w.error}`); return; }
    const st = await spreadsheetsBatchUpdate(env, userId, built.reqs);
    if (!st.ok) errors.push(`'${hstReportTitle}' written but styling failed: ${st.error}`);
    changes.push(`Added '${hstReportTitle}'.`);
  }
}
