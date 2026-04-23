// ════════════════════════════════════════════════════════════════════
// GOOGLE SHEETS OPERATIONS
// Generic read/write/append operations on user's sheet via their access token
// ════════════════════════════════════════════════════════════════════

import { getGoogleAccessToken, getUserSheetId } from './_google.js';

const SHEETS_API = 'https://sheets.googleapis.com/v4/spreadsheets';

/**
 * Read values from a range. e.g. range = "'💸 Expenses'!B12:L500"
 */
export async function readRange(env, userId, range) {
    const sheetId = await getUserSheetId(env, userId);
    if (!sheetId) return { ok: false, error: 'No sheet connected' };

  const tok = await getGoogleAccessToken(env, userId);
    if (!tok.ok) return tok;

  const resp = await fetch(
        `${SHEETS_API}/${sheetId}/values/${encodeURIComponent(range)}`,
    { headers: { 'Authorization': `Bearer ${tok.accessToken}` } }
      );

  const data = await resp.json();
    if (data.error) return { ok: false, error: data.error.message };
    return { ok: true, values: data.values || [] };
}

/**
 * Append rows to a sheet. rows = [[col1, col2, ...], [col1, col2, ...]]
 * range = "'💸 Expenses'!B:L" or similar — Google auto-finds the next empty row
 */
export async function appendRows(env, userId, range, rows) {
    const sheetId = await getUserSheetId(env, userId);
    if (!sheetId) return { ok: false, error: 'No sheet connected' };

  const tok = await getGoogleAccessToken(env, userId);
    if (!tok.ok) return tok;

  const resp = await fetch(
        `${SHEETS_API}/${sheetId}/values/${encodeURIComponent(range)}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
    {
            method: 'POST',
            headers: {
                      'Authorization': `Bearer ${tok.accessToken}`,
                      'Content-Type': 'application/json',
            },
            body: JSON.stringify({ values: rows }),
    }
      );

  const data = await resp.json();
    if (data.error) return { ok: false, error: data.error.message };
    return { ok: true, updates: data.updates };
}

/**
 * Write values to a specific range (overwrites). Used for updating statuses etc.
 */
export async function writeRange(env, userId, range, values) {
    const sheetId = await getUserSheetId(env, userId);
    if (!sheetId) return { ok: false, error: 'No sheet connected' };

  const tok = await getGoogleAccessToken(env, userId);
    if (!tok.ok) return tok;

  const resp = await fetch(
        `${SHEETS_API}/${sheetId}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`,
    {
            method: 'PUT',
            headers: {
                      'Authorization': `Bearer ${tok.accessToken}`,
                      'Content-Type': 'application/json',
            },
            body: JSON.stringify({ values }),
    }
      );

  const data = await resp.json();
    if (data.error) return { ok: false, error: data.error.message };
    return { ok: true };
}

/**
 * Batch multiple operations in one API call. ops = [{range, values}, ...]
 */
export async function batchUpdate(env, userId, data) {
    const sheetId = await getUserSheetId(env, userId);
    if (!sheetId) return { ok: false, error: 'No sheet connected' };

  const tok = await getGoogleAccessToken(env, userId);
    if (!tok.ok) return tok;

  const resp = await fetch(
        `${SHEETS_API}/${sheetId}/values:batchUpdate`,
    {
            method: 'POST',
            headers: {
                      'Authorization': `Bearer ${tok.accessToken}`,
                      'Content-Type': 'application/json',
            },
            body: JSON.stringify({ valueInputOption: 'USER_ENTERED', data }),
    }
      );

  const result = await resp.json();
    if (result.error) return { ok: false, error: result.error.message };
    return { ok: true, result };
}

/**
 * Find the next empty row in a column (for dedup + insertion).
 */
export async function findNextEmptyRow(env, userId, tab, column = 'B', startRow = 12) {
    const range = `'${tab}'!${column}${startRow}:${column}500`;
    const result = await readRange(env, userId, range);
    if (!result.ok) return { ok: false, error: result.error };

  const values = result.values;
    for (let i = 0; i < values.length; i++) {
          if (!values[i] || !values[i][0]) return { ok: true, row: startRow + i };
    }
    return { ok: true, row: startRow + values.length };
}

/**
 * Read existing refs from a column (for dedup).
 */
export async function readExistingRefs(env, userId, tab, refColLetter, startRow = 12) {
    const range = `'${tab}'!${refColLetter}${startRow}:${refColLetter}500`;
    const result = await readRange(env, userId, range);
    if (!result.ok) return new Set();
    return new Set(
          result.values
            .flat()
            .filter(Boolean)
            .map(v => String(v).toLowerCase())
        );
}

/**
 * Fingerprint generator for dedup (same logic as Apps Script had).
 */
export function generateRef(bank, date, amount, vendor) {
    const a = Math.abs(parseFloat(amount) || 0).toFixed(2);
    const v = String(vendor || '').substring(0, 20).replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
    const d = String(date || '').replace(/\//g, '-');
    return `${bank}-${d}-${a}-${v}`;
}
