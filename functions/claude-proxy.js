// ════════════════════════════════════════════════════════════════════
// /functions/claude-proxy.js
// Cloudflare Pages Function — proxies Claude API calls
//
// Changes from the original single-user version:
// - Validates JWT auth token (optional during migration, required for production)
// - Tracks API usage per user per day in D1
// - Rate limits: 150 AI calls per user per day
// - Passes through all Claude API features (text, image, PDF)
//
// Required env vars:
//   ANTHROPIC_API_KEY — your Anthropic API key
//   JWT_SECRET — same secret used by auth endpoints
//
// Required D1 binding:
//   DB — bound to your TradeBooks D1 database
// ════════════════════════════════════════════════════════════════════

import { CORS, json, options, verifyToken } from './_shared.js';

const DAILY_LIMIT = 150; // AI calls per user per day

export async function onRequestOptions() { return options(); }

export async function onRequestPost(context) {
  const { request, env } = context;

  const apiKey = env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return json({ error: 'API key not configured on server' }, 500);
  }

  // ── Authenticate user (extract from JWT if present) ──
  let userId = null;
  const authHeader = request.headers.get('Authorization');
  if (authHeader && authHeader.startsWith('Bearer ') && env.JWT_SECRET) {
    const token = authHeader.slice(7);
    const payload = await verifyToken(token, env.JWT_SECRET);
    if (payload && payload.userId) {
      userId = payload.userId;
    }
  }

  // If no valid auth and we're in production mode, reject
  // For now, allow unauthenticated calls during migration
  // Uncomment the next 3 lines when ready to enforce auth:
  // if (!userId) {
  //   return json({ error: 'Authentication required' }, 401);
  // }

  // ── Rate limiting (per-user, per-day) ──
  if (userId && env.DB) {
    const today = new Date().toISOString().split('T')[0];
    try {
      // Get or create today's usage record
      const usage = await env.DB.prepare(
        'SELECT call_count FROM api_usage WHERE user_id = ? AND date = ?'
      ).bind(userId, today).first();

      if (usage && usage.call_count >= DAILY_LIMIT) {
        return json({
          error: `Daily AI limit reached (${DAILY_LIMIT} calls). Resets at midnight UTC. Contact support if you need more.`
        }, 429);
      }

      // Increment usage
      await env.DB.prepare(`
        INSERT INTO api_usage (user_id, date, call_count, token_count) 
        VALUES (?, ?, 1, 0)
        ON CONFLICT(user_id, date) DO UPDATE SET 
          call_count = call_count + 1
      `).bind(userId, today).run();
    } catch (e) {
      // Don't block the request if usage tracking fails
      console.error('Usage tracking error:', e.message);
    }
  }

  // ── Proxy the request to Anthropic ──
  try {
    const body = await request.json();

    const headers = {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
    };

    // Support PDF beta header
    if (body.beta) {
      headers['anthropic-beta'] = body.beta;
      delete body.beta;
    }

    const anthropicResponse = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers,
      body: JSON.stringify({
        model: body.model || 'claude-sonnet-4-20250514',
        max_tokens: body.max_tokens || 1000,
        messages: body.messages || [],
        ...(body.system ? { system: body.system } : {}),
      })
    });

    const data = await anthropicResponse.json();

    // Track token usage if available
    if (userId && env.DB && data.usage) {
      const today = new Date().toISOString().split('T')[0];
      const tokens = (data.usage.input_tokens || 0) + (data.usage.output_tokens || 0);
      try {
        await env.DB.prepare(`
          UPDATE api_usage SET token_count = token_count + ? 
          WHERE user_id = ? AND date = ?
        `).bind(tokens, userId, today).run();
      } catch (e) {
        // Non-blocking
      }
    }

    if (!anthropicResponse.ok) {
      return json({
        error: data.error?.message || `Anthropic API error (${anthropicResponse.status})`
      }, anthropicResponse.status);
    }

    return json(data);

  } catch (err) {
    return json({ error: 'Proxy error: ' + err.message }, 500);
  }
}
