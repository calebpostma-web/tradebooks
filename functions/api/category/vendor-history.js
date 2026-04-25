// ════════════════════════════════════════════════════════════════════
// GET /api/category/vendor-history
//
// Returns a map of { normalizedVendor: { category, count, exemplars } } built
// from the user's Transactions tab. Used by the importer to skip the AI for
// vendors the user has already categorized — tradebooks gets smarter the more
// it's used.
//
// Vendor normalization: lowercase, strip non-alphanumerics, collapse spaces.
// "Tire Discounter Group I East Garafrax" → "tirediscountergroup..." but we
// keep the first 30 chars to allow fuzzy substring match on import.
//
// Response:
//   {
//     ok: true,
//     count: 156,           // total Transactions rows scanned
//     uniqueVendors: 42,    // distinct vendor keys
//     map: {
//       "tirediscountergroup": { category: "Vehicle Repairs", count: 12, exemplar: "Tire Discounter Group..." }
//     }
//   }
// ════════════════════════════════════════════════════════════════════

import { readRange, getSpreadsheetMetadata, resolveTabName } from '../../_sheets.js';
import { authenticateRequest, json, options } from '../../_shared.js';

export const onRequestOptions = () => options();

export async function onRequest({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  const userId = auth.userId;

  // Resolve the actual Transactions tab title (handles legacy emoji prefixes)
  const meta = await getSpreadsheetMetadata(env, userId);
  if (!meta.ok) return json({ ok: false, error: 'Could not read sheet: ' + meta.error });
  const txnTitle = resolveTabName(meta.sheets, 'Transactions');
  if (!txnTitle) return json({ ok: true, count: 0, uniqueVendors: 0, map: {} });

  // Read columns C (Party/Vendor) and F (Category) — these are at indices 1 and 4
  // when reading from B12:F. We only need vendor + category.
  const r = await readRange(env, userId, `'${txnTitle}'!C12:F1000`);
  if (!r.ok) return json({ ok: false, error: 'Failed to read Transactions: ' + r.error });

  // Group by normalized vendor → tally categories. Pick the most frequent
  // category as the "learned" one. Skip rows missing either field, and skip
  // anything categorized as Internal Transfer (those are bank movements, not
  // user-categorized vendors).
  const buckets = {};   // normalizedVendor → { catCounts: {cat: n}, exemplar }
  let scanned = 0;
  for (const row of (r.values || [])) {
    const vendor = (row[0] || '').toString().trim();
    const category = (row[3] || '').toString().trim();
    if (!vendor || !category) continue;
    if (category === 'Internal Transfer' || category === 'SKIP — not a business expense') continue;
    const norm = normalizeVendor(vendor);
    if (!norm) continue;
    if (!buckets[norm]) buckets[norm] = { catCounts: {}, exemplar: vendor };
    buckets[norm].catCounts[category] = (buckets[norm].catCounts[category] || 0) + 1;
    scanned++;
  }

  // For each vendor, pick the most common category. Tie-break: most recent (we
  // don't track recency here, so first-seen wins — fine for first iteration).
  const map = {};
  for (const [norm, bucket] of Object.entries(buckets)) {
    let best = null, bestN = 0;
    for (const [cat, n] of Object.entries(bucket.catCounts)) {
      if (n > bestN) { best = cat; bestN = n; }
    }
    if (best) map[norm] = { category: best, count: bestN, exemplar: bucket.exemplar };
  }

  return json({
    ok: true,
    count: scanned,
    uniqueVendors: Object.keys(map).length,
    map,
  });
}

// Normalize a vendor name for fuzzy matching.
// Strips: punctuation, location suffixes ("CHATHAM" etc), common store#s,
// and non-alphanumeric chars. Lowercases. Caps at 30 chars to avoid bloat.
// Goal: "Tire Discounter Group I East Garafrax" and "TIRE DISCOUNTER GROUP CHATHAM"
// should normalize to overlapping prefixes that can match.
function normalizeVendor(s) {
  return String(s || '')
    .toLowerCase()
    // Strip very common location/store-number patterns
    .replace(/\b(?:chatham|kent|london|toronto|ottawa|hamilton|windsor|sarnia)\b/gi, '')
    .replace(/#?\d{2,}/g, '')      // store numbers
    .replace(/[^a-z0-9]/g, '')      // collapse to alphanumeric only
    .slice(0, 30);
}
