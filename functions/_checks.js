// ════════════════════════════════════════════════════════════════════
// ⚠️ CHECKS — "what needs my attention" list, computed from the ledger
//
// One formula-driven tab. Top: a count per rule. Below: one line per
// problem row in plain words, with a link that jumps to that row in
// 📒 Transactions. Rebuilt live by the sheet; nobody types here.
//
// Rules (per Transactions row):
//   HST      HST? = Yes but the HST amount isn't the configured rate (±2¢),
//            or HST? = No with an HST amount. Rows mentioning "Status"
//            (Status-card customers, off-reserve) may carry GST-only (5%):
//            those are accepted and listed in their own informational block.
//   Category blank
//   Date     not a real date, or outside the fiscal year (from HST Returns C3)
//   Duplicate same payee + same amount within 2 days (possible; review)
//   Account  not the configured bank or card (e.g. "Receipt", blank)
//   Amount   missing/zero on a dated row
// Plus: any period on 🏦 Account Balances whose Match column says "⚠ Off by".
//
// Idempotent via a layout version in the A1 banner.
// ════════════════════════════════════════════════════════════════════

import { readRange, batchUpdate, spreadsheetsBatchUpdate } from './_sheets.js';

export const CHECKS_TITLE = '⚠️ Checks';
const LAYOUT_VERSION = 2;
const GST_RATE = 0.05;      // federal part, for Status-card GST-only sales

const COLORS = {
  amber:     { red: 0.60, green: 0.36, blue: 0.02 },
  amberTint: { red: 0.996, green: 0.965, blue: 0.90 },
  white:     { red: 1, green: 1, blue: 1 },
  teal:      { red: 0.078, green: 0.278, blue: 0.247 },
};
const FMT_CURRENCY = { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00;("$"#,##0.00)' } };
const FMT_DATE     = { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } };

const q = t => `'${String(t).replace(/'/g, "''")}'`;
const esc = s => String(s).replace(/"/g, '""');
const findTab = (m, re) => Object.values(m).find(s => re.test(s.title));

// Rule labels — the Problem column is built from these, and the summary
// counts search for them, so keep the two in sync.
const RULES = [
  ['HST',       'HST amount doesn\'t match the rate'],
  ['HST-NO',    'HST amount entered but HST? is No'],
  ['CATEGORY',  'No category'],
  ['DATE-BAD',  'Date isn\'t a real date'],
  ['DATE-FY',   'Date is outside this fiscal year'],
  ['DUP',       'Possible duplicate (same payee, amount, within 2 days)'],
  ['ACCOUNT',   'Unknown account'],
  ['AMOUNT',    'Amount missing'],
];

function buildChecksTab({ title, sheetId, txnTitle, txnSheetId, cfgTitle, hstTitle, balTitle, bank, card }) {
  const T = q(txnTitle);
  const B = `${T}!B12:B`, C = `${T}!C12:C`, D = `${T}!D12:D`, E = `${T}!E12:E`,
        F = `${T}!F12:F`, G = `${T}!G12:G`, H = `${T}!H12:H`, I = `${T}!I12:I`;
  const rate = `${q(cfgTitle)}!$C$10`;
  const fyStart = `${q(hstTitle)}!$C$3`;
  const label = k => RULES.find(r => r[0] === k)[1];

  // Per-row expressions (all arrays over the Transactions rows)
  const hasRow   = `((LEN(${B})+LEN(${C})+LEN(${E}))>0)`;
  const notXfer  = `(${F}<>"Internal Transfer")`;
  const status   = `REGEXMATCH(UPPER(${C}&" "&${D}),"STATUS")`;
  const expHST   = `ROUND(ABS(IFERROR(${E}*1,0))*${rate},2)`;
  const gstOnly  = `ROUND(ABS(IFERROR(${E}*1,0))*${GST_RATE},2)`;
  const hstOff   = `(ABS(IFERROR(${H}*1,0)-${expHST})>0.02)`;
  const hstIsGst = `(ABS(IFERROR(${H}*1,0)-${gstOnly})<=0.02)`;   // GST-only (off-reserve Status-card sale); 0% is still flagged
  const ruleHST  = `IF(${hasRow}*${notXfer}*(${G}="Yes")*${hstOff}*NOT(${status}*${hstIsGst}),"${esc(label('HST'))}; ","")`;
  const ruleHSTn = `IF(${hasRow}*(${G}="No")*(IFERROR(${H}*1,0)<>0),"${esc(label('HST-NO'))}; ","")`;
  const ruleCat  = `IF(${hasRow}*(LEN(${F})=0),"${esc(label('CATEGORY'))}; ","")`;
  const ruleDate = `IF(${hasRow}*NOT(ISNUMBER(${B})),"${esc(label('DATE-BAD'))}; ",IF(${hasRow}*ISNUMBER(${B})*((${B}<${fyStart})+(${B}>EDATE(${fyStart},12)-1)),"${esc(label('DATE-FY'))}; ",""))`;
  const dnum     = `IF(ISNUMBER(${B}),${B},0)`;   // text dates would error inside COUNTIFS criteria
  const ruleDup  = `IF(${hasRow}*ISNUMBER(${B})*(IFERROR(${E}*1,0)<>0)*(COUNTIFS(${C},${C},${E},${E},${B},">="&(${dnum}-2),${B},"<="&(${dnum}+2))>1),"${esc(label('DUP'))}; ","")`;
  const ruleAcct = `IF(${hasRow}*(${I}<>"${esc(bank)}")*(${I}<>"${esc(card)}"),"${esc(label('ACCOUNT'))}: "&${I}&"; ","")`;
  const ruleAmt  = `IF(${hasRow}*ISNUMBER(${B})*(IFERROR(${E}*1,0)=0),"${esc(label('AMOUNT'))}; ","")`;
  const safe = r => `IFERROR(${r},"(check error); ")`;
  const flags = [ruleHST, ruleHSTn, ruleCat, ruleDate, ruleDup, ruleAcct, ruleAmt].map(safe).join('&');

  const listCore = `SORT(FILTER({ROW(${B}),${B},${C},${E},${F},${H},${I},flags},flags<>""),2,TRUE)`;
  const listFormula =
    `=LET(flags,ARRAYFORMULA(${flags}),` +
    `IFERROR(${listCore},IF(COUNTIF(flags,"?*")=0,"No problems found","CHECK ENGINE ERROR — see status cell")))`;
  const statusCell =
    `=LET(flags,ARRAYFORMULA(${flags}),x,IFERROR(${listCore},NA()),` +
    `IF(ISNA(x),IF(COUNTIF(flags,"?*")=0,"OK — nothing to list","ERROR "&IFERROR(ERROR.TYPE(${listCore}),"?")),"OK"))`;

  const statusFormula =
    `=IFERROR(SORT(FILTER({ROW(${B}),${B},${C},${E},${H},ROUND(ABS(IFERROR(${E}*1,0))*(${rate}-${GST_RATE}),2)},` +
    `ARRAYFORMULA(${hasRow}*${status}*(${G}="Yes")*${hstIsGst})),2,TRUE),"None")`;

  const gid = txnSheetId;
  const S = 2000;
  const values = [
    { range: `${q(title)}!A1`, values: [[`CHECKS  ·  what needs a look, straight from the ledger  ·  click "open" to jump to the row  ·  v${LAYOUT_VERSION}`]] },
    { range: `${q(title)}!B3:C3`, values: [['Problems found', `=IF(ISNUMBER(B14),COUNTA(B14:B),0)`]] },
    { range: `${q(title)}!B4:C11`, values: RULES.map(([, lbl]) => [lbl, `=COUNTIF($I$14:$I,"*${esc(lbl)}*")`]) },
    { range: `${q(title)}!E6:F6`, values: [['Check engine', statusCell]] },
    { range: `${q(title)}!E3:F4`, values: [
      ['Bank periods off (🏦 Account Balances)', `=IFERROR(COUNTIF(${q(balTitle)}!K12:K,"⚠*"),0)`],
      ['Status-card GST-only sales (info)', `=IF(ISNUMBER(B${S + 2}),COUNTA(B${S + 2}:B),0)`],
    ]},
    { range: `${q(title)}!B13:J13`, values: [['Row', 'Date', 'Name', 'Amount', 'Category', 'HST', 'Account', 'Problem(s)', 'Open']] },
    { range: `${q(title)}!B14`, values: [[listFormula]] },
    { range: `${q(title)}!J14`, values: [[`=ARRAYFORMULA(IF(ISNUMBER(B14:B),HYPERLINK("#gid=${gid}&range=B"&B14:B,"open"),""))`]] },
  ];

  // Status-card block goes far down so the main list has room to grow (row S)
  values.push({ range: `${q(title)}!B${S}`, values: [['STATUS-CARD SALES  ·  GST-only rows (accepted by the HST check)  ·  last column = 8% Ontario part credited at point of sale']] });
  values.push({ range: `${q(title)}!B${S + 1}:G${S + 1}`, values: [['Row', 'Date', 'Name', 'Amount', 'HST charged', '8% credited']] });
  values.push({ range: `${q(title)}!B${S + 2}`, values: [[statusFormula]] });
  values.push({ range: `${q(title)}!E${S - 2}:F${S - 2}`, values: [['Total 8% credited', `=SUM(G${S + 2}:G)`]] });

  const reqs = [
    { repeatCell: { range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 11 },
        cell: { userEnteredFormat: { backgroundColor: COLORS.amber, textFormat: { foregroundColor: COLORS.white, bold: true, fontSize: 11 } } }, fields: 'userEnteredFormat' } },
    { repeatCell: { range: { sheetId, startRowIndex: 2, endRowIndex: 11, startColumnIndex: 1, endColumnIndex: 3 },
        cell: { userEnteredFormat: { backgroundColor: COLORS.amberTint } }, fields: 'userEnteredFormat.backgroundColor' } },
    { repeatCell: { range: { sheetId, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 1, endColumnIndex: 3 },
        cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 12 } } }, fields: 'userEnteredFormat.textFormat' } },
    { repeatCell: { range: { sheetId, startRowIndex: 12, endRowIndex: 13, startColumnIndex: 1, endColumnIndex: 10 },
        cell: { userEnteredFormat: { backgroundColor: COLORS.teal, textFormat: { foregroundColor: COLORS.white, bold: true } } }, fields: 'userEnteredFormat' } },
    { repeatCell: { range: { sheetId, startRowIndex: S, endRowIndex: S + 1, startColumnIndex: 1, endColumnIndex: 7 },
        cell: { userEnteredFormat: { backgroundColor: COLORS.teal, textFormat: { foregroundColor: COLORS.white, bold: true } } }, fields: 'userEnteredFormat' } },
    { repeatCell: { range: { sheetId, startRowIndex: 13, endRowIndex: 5000, startColumnIndex: 2, endColumnIndex: 3 }, cell: { userEnteredFormat: FMT_DATE }, fields: 'userEnteredFormat.numberFormat' } },
    { repeatCell: { range: { sheetId, startRowIndex: 13, endRowIndex: 5000, startColumnIndex: 4, endColumnIndex: 5 }, cell: { userEnteredFormat: FMT_CURRENCY }, fields: 'userEnteredFormat.numberFormat' } },
    { repeatCell: { range: { sheetId, startRowIndex: 13, endRowIndex: 5000, startColumnIndex: 6, endColumnIndex: 7 }, cell: { userEnteredFormat: FMT_CURRENCY }, fields: 'userEnteredFormat.numberFormat' } },
    { repeatCell: { range: { sheetId, startRowIndex: S - 3, endRowIndex: 5000, startColumnIndex: 5, endColumnIndex: 6 }, cell: { userEnteredFormat: FMT_CURRENCY }, fields: 'userEnteredFormat.numberFormat' } },
    { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 300 }, fields: 'pixelSize' } },
    { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 220 }, fields: 'pixelSize' } },
    { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 8, endIndex: 9 }, properties: { pixelSize: 420 }, fields: 'pixelSize' } },
    { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 260 }, fields: 'pixelSize' } },
    { updateSheetProperties: { properties: { sheetId, gridProperties: { frozenRowCount: 13 } }, fields: 'gridProperties.frozenRowCount' } },
  ];
  return { values, reqs };
}

export async function ensureChecksTab(env, userId, { sheetsByTitle, profile, dryRun, changes, errors }) {
  const txnTab = findTab(sheetsByTitle, /transactions/i);
  const cfgTab = findTab(sheetsByTitle, /config/i);
  const hstTab = findTab(sheetsByTitle, /hst returns/i);
  const balTab = findTab(sheetsByTitle, /account balances/i);
  if (!txnTab || !cfgTab || !hstTab) return;

  const existing = sheetsByTitle[CHECKS_TITLE];
  if (existing) {
    const banner = await readRange(env, userId, `${q(CHECKS_TITLE)}!A1`);
    const ok = banner.ok && banner.values && banner.values[0] && String(banner.values[0][0] || '').endsWith(`v${LAYOUT_VERSION}`);
    if (ok) return;
  }
  if (dryRun) {
    changes.push(existing ? `Rebuild '${CHECKS_TITLE}' with the latest rules.` : `Add '${CHECKS_TITLE}' — a list of ledger rows that need a look (HST off, no category, bad date, possible duplicate, unknown account), with a link to each row.`);
    return;
  }
  let sheetId = existing?.sheetId;
  if (!existing) {
    const used = new Set(Object.values(sheetsByTitle).map(s => s.sheetId));
    sheetId = 1600; while (used.has(sheetId)) sheetId += 1;
    const add = await spreadsheetsBatchUpdate(env, userId, [{ addSheet: { properties: { sheetId, title: CHECKS_TITLE, index: 2,
      gridProperties: { rowCount: 5000, columnCount: 12, frozenRowCount: 13 }, tabColor: COLORS.amber } } }]);
    if (!add.ok) { errors.push(`Could not create '${CHECKS_TITLE}': ${add.error}`); return; }
    sheetsByTitle[CHECKS_TITLE] = { sheetId, title: CHECKS_TITLE };
  } else {
    const clr = await spreadsheetsBatchUpdate(env, userId, [{ updateCells: { range: { sheetId }, fields: 'userEnteredValue,userEnteredFormat' } }]);
    if (!clr.ok) { errors.push(`Could not reset '${CHECKS_TITLE}': ${clr.error}`); return; }
  }
  const built = buildChecksTab({
    title: CHECKS_TITLE, sheetId, txnTitle: txnTab.title, txnSheetId: txnTab.sheetId,
    cfgTitle: cfgTab.title, hstTitle: hstTab.title, balTitle: balTab ? balTab.title : '🏦 Account Balances',
    bank: String(profile?.primaryBank || 'BMO').trim(), card: String(profile?.creditCard || 'AMEX').trim(),
  });
  const w = await batchUpdate(env, userId, built.values);
  if (!w.ok) { errors.push(`Could not write '${CHECKS_TITLE}': ${w.error}`); return; }
  const st = await spreadsheetsBatchUpdate(env, userId, built.reqs);
  if (!st.ok) errors.push(`'${CHECKS_TITLE}' written but styling failed: ${st.error}`);
  changes.push(existing ? `Rebuilt '${CHECKS_TITLE}'.` : `Added '${CHECKS_TITLE}' — rows that need a look, with a link to each.`);
}
