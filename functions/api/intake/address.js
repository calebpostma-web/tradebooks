// ════════════════════════════════════════════════════════════════════
// GET  /api/intake/address   → { address, token, events:[…recent] }
// POST /api/intake/address   → (re)generate the token; body { rotate: true }
//
// The intake address is where the user forwards their website's order
// emails (Gmail filter → forward). The Worker that receives mail on the
// domain looks the token up to find the account.
// ════════════════════════════════════════════════════════════════════
import { authenticateRequest, json, options } from '../../_shared.js';

const DOMAIN = 'aibookkeeper.ca';

export const onRequestOptions = () => options();

function newToken() {
  const alphabet = 'abcdefghjkmnpqrstuvwxyz23456789';        // no 0/o/1/l/i — it gets typed into Gmail
  const bytes = crypto.getRandomValues(new Uint8Array(8));
  return Array.from(bytes, b => alphabet[b % alphabet.length]).join('');
}

async function ensureToken(env, userId, rotate = false) {
  const row = await env.DB.prepare('SELECT intake_token FROM profiles WHERE user_id = ?').bind(userId).first();
  if (row?.intake_token && !rotate) return row.intake_token;
  for (let i = 0; i < 5; i++) {
    const t = newToken();
    try {
      await env.DB.prepare('UPDATE profiles SET intake_token = ? WHERE user_id = ?').bind(t, userId).run();
      return t;
    } catch (e) { if (!/UNIQUE/i.test(String(e.message))) throw e; }
  }
  throw new Error('Could not allocate an intake token');
}

async function recentEvents(env, userId) {
  try {
    const r = await env.DB.prepare(
      'SELECT id, received_at, from_addr, subject, kind, order_no, status, detail, payload FROM intake_events WHERE user_id = ? ORDER BY id DESC LIMIT 40'
    ).bind(userId).all();
    return r.results || [];
  } catch { return []; }
}

export async function onRequestGet({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  try {
    const token = await ensureToken(env, auth.userId);
    const events = await recentEvents(env, auth.userId);
    const verify = events.find(e => e.kind === 'gmail_verify');
    return json({ ok: true, token, address: `orders-${token}@${DOMAIN}`, events: events.map(({ payload, ...e }) => e), gmailVerify: verify ? (() => { let pl = {}; try { pl = JSON.parse(verify.payload || '{}') || {}; } catch { /* legacy */ } return { code: pl.code || (/^\d{6,12}$/.test(verify.detail) ? verify.detail : ''), link: pl.link || '', at: verify.received_at }; })() : null });
  } catch (e) {
    return json({ ok: false, error: e.message }, 500);
  }
}

export async function onRequestPost({ request, env }) {
  const auth = await authenticateRequest(request, env);
  if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
  let body = {}; try { body = await request.json(); } catch { /* empty */ }
  try {
    const token = await ensureToken(env, auth.userId, !!body.rotate);
    return json({ ok: true, token, address: `orders-${token}@${DOMAIN}` });
  } catch (e) {
    return json({ ok: false, error: e.message }, 500);
  }
}
