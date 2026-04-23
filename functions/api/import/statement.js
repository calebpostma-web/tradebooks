// ════════════════════════════════════════════════════════════════════
// POST /api/import/statement
// Import parsed bank statement rows into Expenses + Income tabs.
// Replaces Apps Script's handleRows().
// ════════════════════════════════════════════════════════════════════

import { appendRows, writeRange, readExistingRefs, generateRef } from '../../_sheets.js';
import { authenticateRequest, json, options } from '../../_shared.js';

const EXPENSES_TAB = '💸 Expenses';
const INCOME_TAB = '💰 Income';

const DEFAULT_INCOME_CATS = new Set([
    'Consulting Revenue', 'Service Revenue', 'Sales Revenue',
    'Sales Revenue — Materials', 'Rental Income', 'Interest Income',
    'Income Received', 'Other Income',
  ]);

const DEFAULT_SKIP_CATS = new Set([
    'Bill Payment / Transfer', 'Owner Draw / Distribution', 'SKIP — not a business expense',
  ]);

export const onRequestOptions = () => options();

export async function onRequestPost({ request, env }) {
    const auth = await authenticateRequest(request, env);
    if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
    const userId = auth.userId;

  let body;
    try {
          body = await request.json();
    } catch {
          return json({ ok: false, error: 'Invalid JSON' }, 400);
    }

  const rows = body.rows || [];
    const bank = body.bank || 'AMEX';
    const source = body.source || 'Import';

  if (!rows.length) return json({ ok: false, error: 'No rows' }, 400);

  // Read existing refs for dedup
  const expRefs = await readExistingRefs(env, userId, EXPENSES_TAB, 'L');
    const incRefs = await readExistingRefs(env, userId, INCOME_TAB, 'K');

  // Build rows to write
  const expRows = [], expMeta = [];
    const incRows = [], incMeta = [];
    const duplicateDetails = [];
    const batchCounter = {};
    let skipped = 0, duplicates = 0;

  for (const row of rows) {
        const cat = row.category || '';
        if (DEFAULT_SKIP_CATS.has(cat)) {
                skipped++;
                continue;
        }

      // Batch-aware ref (handles multiple same-amount-same-vendor in one upload)
      const base = generateRef(bank, row.date, row.amount || row.net, row.vendor);
        const baseLC = base.toLowerCase();
        batchCounter[baseLC] = (batchCounter[baseLC] || 0) + 1;
        const ref = batchCounter[baseLC] > 1 ? `${base}-${batchCounter[baseLC]}` : base;
        const refLC = ref.toLowerCase();

      const netAmt = Math.abs(parseFloat(row.net) || parseFloat(row.amount) || 0);

      if (DEFAULT_INCOME_CATS.has(cat)) {
              if (incRefs.has(refLC)) {
                        duplicates++;
                        duplicateDetails.push({ date: row.date, vendor: row.vendor, amount: row.amount });
                        continue;
              }
              incRefs.add(refLC);
              incRows.push([row.date || '', row.vendor || '', row.description || '', '', netAmt, cat]);
              incMeta.push([bank, ref]);
      } else {
              if (expRefs.has(refLC)) {
                        duplicates++;
                        duplicateDetails.push({ date: row.date, vendor: row.vendor, amount: row.amount });
                        continue;
              }
              expRefs.add(refLC);
              expRows.push([row.date || '', row.vendor || '', row.description || '', netAmt, cat]);
              expMeta.push([bank, bank === 'AMEX' ? 'Yes' : 'No', source, ref]);
      }
  }

  // Write expenses — two-range approach: B-F for data, I-L for meta
  if (expRows.length) {
        const appendResult = await appendRows(env, userId, `'${EXPENSES_TAB}'!B12:F`, expRows);
        if (!appendResult.ok) return json({ ok: false, error: 'Expenses write failed: ' + appendResult.error });

      // appendResult.updates.updatedRange looks like "'💸 Expenses'!B87:F91"
      const match = /!B(\d+):F(\d+)/.exec(appendResult.updates.updatedRange);
        if (match) {
                const startRow = parseInt(match[1]);
                const metaRange = `'${EXPENSES_TAB}'!I${startRow}:L${startRow + expRows.length - 1}`;
                const metaResult = await writeRange(env, userId, metaRange, expMeta);
                if (!metaResult.ok) return json({ ok: false, error: 'Expenses meta write failed: ' + metaResult.error });
        }
  }

  // Income: B-G data, J-K meta
  if (incRows.length) {
        const appendResult = await appendRows(env, userId, `'${INCOME_TAB}'!B12:G`, incRows);
        if (!appendResult.ok) return json({ ok: false, error: 'Income write failed: ' + appendResult.error });

      const match = /!B(\d+):G(\d+)/.exec(appendResult.updates.updatedRange);
        if (match) {
                const startRow = parseInt(match[1]);
                const metaRange = `'${INCOME_TAB}'!J${startRow}:K${startRow + incRows.length - 1}`;
                const metaResult = await writeRange(env, userId, metaRange, incMeta);
                if (!metaResult.ok) return json({ ok: false, error: 'Income meta write failed: ' + metaResult.error });
        }
  }

  return json({
        ok: true,
        expenses: expRows.length,
        income: incRows.length,
        skipped,
        duplicates,
        duplicatesBlocked: duplicates,
        duplicateDetails,
        total: rows.length,
  });
}
