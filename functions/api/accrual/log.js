// ════════════════════════════════════════════════════════════════════
// POST /api/accrual/log
//
// Records a year-end adjusting entry on the 📓 Adjusting Entries tab.
// THE linchpin endpoint for accrual-basis prep — without these adjustments
// the cash-basis books aren't ready for the T2, which means MNP has to
// make the entries themselves (= full prep, not review-only).
//
// Request body:
//   {
//     date: 'YYYY-MM-DD',         // typically the FYE date
//     type: 'Accounts Receivable (AR)' | 'Accounts Payable (AP)' |
//           'Prepaid Expense' | 'Accrued Expense' | 'Accrued Revenue' |
//           'Depreciation (CCA)' | 'Other Adjustment',
//     description: string,
//     counterparty: string,        // client / vendor name
//     netAmount: number,           // signed: positive = revenue/asset, negative = reverse
//     hst: number,                 // 0 if not HST-applicable (most accruals don't have HST)
//     effect: string,              // human description of what this changes on the books
//     linkedRef: string,           // optional invoice # or bill ref
//     notes: string,
//     fy: string,                  // e.g. 'FY2026'
//   }
// ════════════════════════════════════════════════════════════════════

import { appendRows, getSpreadsheetMetadata, resolveTabName } from '../../_sheets.js';
import { authenticateRequest, json, options } from '../../_shared.js';

const VALID_TYPES = new Set([
  'Accounts Receivable (AR)',
  'Accounts Payable (AP)',
  'Prepaid Expense',
  'Accrued Expense',
  'Accrued Revenue',
  'Depreciation (CCA)',
  'Other Adjustment',
]);

export const onRequestOptions = () => options();

export async function onRequestPost({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  const userId = auth.userId;

  let body;
  try { body = await request.json(); }
  catch { return json({ ok: false, error: 'Invalid JSON' }, 400); }

  // Validate
  const date = (body.date || '').trim();
  const type = (body.type || '').trim();
  const description = (body.description || '').trim();
  const counterparty = (body.counterparty || '').trim();
  const netAmount = Math.round((parseFloat(body.netAmount) || 0) * 100) / 100;
  const hst = Math.round((parseFloat(body.hst) || 0) * 100) / 100;
  const effect = (body.effect || '').trim();
  const linkedRef = (body.linkedRef || '').trim();
  const notes = (body.notes || '').trim();
  const fy = (body.fy || '').trim();

  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return json({ ok: false, error: 'date must be YYYY-MM-DD' }, 400);
  if (!VALID_TYPES.has(type)) return json({ ok: false, error: `Invalid type: ${type}` }, 400);
  if (!description) return json({ ok: false, error: 'description is required' }, 400);
  if (netAmount === 0) return json({ ok: false, error: 'netAmount cannot be zero' }, 400);

  // Resolve actual tab name (handles legacy emoji prefixes)
  const meta = await getSpreadsheetMetadata(env, userId);
  if (!meta.ok) return json({ ok: false, error: 'Could not read sheet: ' + meta.error });
  const adjTitle = resolveTabName(meta.sheets, 'Adjusting Entries');
  if (!adjTitle) {
    return json({
      ok: false,
      error: "'📓 Adjusting Entries' tab not found. Run Settings → Update sheet to latest schema.",
      needsMigration: true,
    });
  }

  // Append the row. Total (col H) is filled by the ARRAYFORMULA in H12 — we
  // leave it blank in the row data so the formula populates it.
  // Cols B-L: Date | Type | Description | Counterparty | Net Amount | HST | Total (formula) | Effect | Linked Ref | Notes | FY
  const row = [[
    date, type, description, counterparty, netAmount, hst, '', effect, linkedRef, notes, fy,
  ]];

  const result = await appendRows(env, userId, `'${adjTitle}'!B12:L`, row);
  if (!result.ok) return json({ ok: false, error: 'Adjusting entry write failed: ' + result.error });

  return json({
    ok: true,
    type, date, netAmount, hst,
    range: result.updates?.updatedRange,
  });
}
