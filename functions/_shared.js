// ════════════════════════════════════════════════════════════════════
// TradeBooks — Shared Worker Utilities
// Used by all /functions/*.js endpoints
// ════════════════════════════════════════════════════════════════════

export const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization',
  'Content-Type': 'application/json',
};

export function json(data, status = 200) {
  return new Response(JSON.stringify(data), { status, headers: CORS });
}

export function options() {
  return new Response(null, { status: 204, headers: CORS });
}

// ── JWT using HMAC-SHA256 (Web Crypto API — works in Workers) ──

async function getKey(secret) {
  const enc = new TextEncoder();
  return crypto.subtle.importKey(
    'raw', enc.encode(secret), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign', 'verify']
  );
}

function base64url(buf) {
  return btoa(String.fromCharCode(...new Uint8Array(buf)))
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

function base64urlDecode(str) {
  str = str.replace(/-/g, '+').replace(/_/g, '/');
  while (str.length % 4) str += '=';
  const bin = atob(str);
  return new Uint8Array([...bin].map(c => c.charCodeAt(0)));
}

export async function createToken(payload, secret, expiresInHours = 720) {
  // 720 hours = 30 days default
  const header = { alg: 'HS256', typ: 'JWT' };
  const now = Math.floor(Date.now() / 1000);
  const body = { ...payload, iat: now, exp: now + expiresInHours * 3600 };

  const enc = new TextEncoder();
  const headerB64 = base64url(enc.encode(JSON.stringify(header)));
  const bodyB64 = base64url(enc.encode(JSON.stringify(body)));
  const message = `${headerB64}.${bodyB64}`;

  const key = await getKey(secret);
  const sig = await crypto.subtle.sign('HMAC', key, enc.encode(message));

  return `${message}.${base64url(sig)}`;
}

export async function verifyToken(token, secret) {
  try {
    const parts = token.split('.');
    if (parts.length !== 3) return null;

    const [headerB64, bodyB64, sigB64] = parts;
    const message = `${headerB64}.${bodyB64}`;

    const key = await getKey(secret);
    const enc = new TextEncoder();
    const sig = base64urlDecode(sigB64);

    const valid = await crypto.subtle.verify('HMAC', key, sig, enc.encode(message));
    if (!valid) return null;

    const body = JSON.parse(new TextDecoder().decode(base64urlDecode(bodyB64)));

    // Check expiry
    if (body.exp && body.exp < Math.floor(Date.now() / 1000)) return null;

    return body;
  } catch {
    return null;
  }
}

// ── Password hashing using PBKDF2 (Web Crypto API) ──

export async function hashPassword(password) {
  const enc = new TextEncoder();
  const salt = crypto.getRandomValues(new Uint8Array(16));
  const key = await crypto.subtle.importKey('raw', enc.encode(password), 'PBKDF2', false, ['deriveBits']);
  const hash = await crypto.subtle.deriveBits(
    { name: 'PBKDF2', salt, iterations: 100000, hash: 'SHA-256' }, key, 256
  );
  const saltHex = [...salt].map(b => b.toString(16).padStart(2, '0')).join('');
  const hashHex = [...new Uint8Array(hash)].map(b => b.toString(16).padStart(2, '0')).join('');
  return `${saltHex}:${hashHex}`;
}

export async function verifyPassword(password, stored) {
  const [saltHex, hashHex] = stored.split(':');
  const salt = new Uint8Array(saltHex.match(/.{2}/g).map(h => parseInt(h, 16)));
  const enc = new TextEncoder();
  const key = await crypto.subtle.importKey('raw', enc.encode(password), 'PBKDF2', false, ['deriveBits']);
  const hash = await crypto.subtle.deriveBits(
    { name: 'PBKDF2', salt, iterations: 100000, hash: 'SHA-256' }, key, 256
  );
  const computedHex = [...new Uint8Array(hash)].map(b => b.toString(16).padStart(2, '0')).join('');
  return computedHex === hashHex;
}

// ── Auth middleware — extract user from JWT ──

export async function authenticateRequest(request, env) {
  const authHeader = request.headers.get('Authorization');
  if (!authHeader || !authHeader.startsWith('Bearer ')) return null;

  const token = authHeader.slice(7);
  const payload = await verifyToken(token, env.JWT_SECRET);
  if (!payload || !payload.userId) return null;

  return payload;
}

// ── Fiscal year math ────────────────────────────────────────────────
// All year-end / quarterly tooling needs a consistent way to derive the
// fiscal year window from a user's `fiscalYearEnd` profile field
// (e.g. "March 31", "December 31"). FY{N} ends on that date in calendar year N.
// e.g. Postma fiscalYearEnd="March 31", so FY2026 = Apr 1 2025 → Mar 31 2026.

const MONTHS = {
  january:1, february:2, march:3, april:4, may:5, june:6,
  july:7, august:8, september:9, october:10, november:11, december:12,
};

/**
 * Parse a fiscal-year-end string like "March 31" → { month: 3, day: 31 }.
 * Falls back to Dec 31 (calendar year) on bad input.
 */
export function parseFiscalYearEnd(s) {
  const m = String(s || '').trim().match(/^([a-zA-Z]+)\s+(\d{1,2})$/);
  if (!m) return { month: 12, day: 31 };
  const month = MONTHS[m[1].toLowerCase()] || 12;
  const day = Math.min(31, Math.max(1, parseInt(m[2], 10) || 31));
  return { month, day };
}

/**
 * Given a fiscal-year-end + optional reference date (defaults to today),
 * return the fiscal year window that the reference date falls inside.
 * Returns: { fyLabel: 'FY2026', startISO: '2025-04-01', endISO: '2026-03-31',
 *            startDate, endDate }
 */
export function fiscalYearOf(fyeString, refDate = new Date()) {
  const { month, day } = parseFiscalYearEnd(fyeString);
  // FY{N} ends on month/day of year N. The fiscal year that contains refDate
  // is the smallest N such that refDate <= end-of-FY{N}.
  const refYear = refDate.getUTCFullYear();
  // Construct the FY{refYear} end date:
  let endDate = new Date(Date.UTC(refYear, month - 1, day, 23, 59, 59));
  // If refDate is after that, we're in FY{refYear + 1}.
  let fyEndYear = refYear;
  if (refDate > endDate) {
    fyEndYear = refYear + 1;
    endDate = new Date(Date.UTC(fyEndYear, month - 1, day, 23, 59, 59));
  }
  // Start = day after previous FY end = (fyEndYear - 1, month, day) + 1 day.
  const startDate = new Date(Date.UTC(fyEndYear - 1, month - 1, day + 1, 0, 0, 0));
  return {
    fyLabel: `FY${fyEndYear}`,
    fyEndYear,
    startISO: startDate.toISOString().slice(0, 10),
    endISO: endDate.toISOString().slice(0, 10),
    startDate,
    endDate,
  };
}

/**
 * Enumerate the (year, monthIndex) pairs that compose a fiscal year window.
 * For an Apr 1 → Mar 31 FY: returns 12 entries [{year:2025, month:4}, ..., {year:2026, month:3}].
 * Used by the statement archive checklist (we expect one bank statement per month).
 */
export function fiscalYearMonths(startDate, endDate) {
  const out = [];
  const cur = new Date(Date.UTC(startDate.getUTCFullYear(), startDate.getUTCMonth(), 1));
  const lastMonth = new Date(Date.UTC(endDate.getUTCFullYear(), endDate.getUTCMonth(), 1));
  while (cur <= lastMonth) {
    out.push({ year: cur.getUTCFullYear(), month: cur.getUTCMonth() + 1 });
    cur.setUTCMonth(cur.getUTCMonth() + 1);
  }
  return out;
}

