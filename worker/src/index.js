/**
 * MacroFlow AI proxy.
 *
 * Forwards VBA-generation requests to Anthropic Claude Sonnet 5.
 * The API key lives only in this Worker (wrangler secret ANTHROPIC_API_KEY).
 * The desktop app never receives the key.
 */

const ANTHROPIC_MESSAGES_URL = 'https://api.anthropic.com/v1/messages';
const RATE_LIMIT_CACHE_HOST = 'https://macroflow-ai-rate-limit.local';

function jsonResponse(status, body, extraHeaders = {}) {
  return new Response(JSON.stringify(body), {
    status,
    headers: {
      'Content-Type': 'application/json',
      'Cache-Control': 'no-store',
      ...extraHeaders
    }
  });
}

function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Max-Age': '86400'
  };
}

function getClientIp(request) {
  return (
    request.headers.get('CF-Connecting-IP') ||
    request.headers.get('X-Forwarded-For')?.split(',')[0]?.trim() ||
    'unknown'
  );
}

function parsePositiveInt(value, fallback) {
  const parsed = Number.parseInt(String(value || ''), 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

async function checkRateLimit(request, limitPerHour) {
  const ip = getClientIp(request);
  const hourBucket = Math.floor(Date.now() / (60 * 60 * 1000));
  const cacheKey = new Request(`${RATE_LIMIT_CACHE_HOST}/${encodeURIComponent(ip)}/${hourBucket}`);

  const cache = caches.default;
  const cached = await cache.match(cacheKey);
  let count = 0;
  if (cached) {
    const text = await cached.text();
    count = Number.parseInt(text, 10) || 0;
  }

  if (count >= limitPerHour) {
    return { allowed: false, remaining: 0, count };
  }

  count += 1;
  await cache.put(
    cacheKey,
    new Response(String(count), {
      headers: { 'Cache-Control': 'max-age=3600' }
    })
  );

  return { allowed: true, remaining: Math.max(0, limitPerHour - count), count };
}

function isValidMessages(messages) {
  if (!Array.isArray(messages) || messages.length === 0 || messages.length > 32) {
    return false;
  }
  return messages.every((message) => {
    if (!message || typeof message !== 'object') return false;
    if (message.role !== 'user' && message.role !== 'assistant') return false;
    return typeof message.content === 'string';
  });
}

export default {
  async fetch(request, env) {
    const headers = corsHeaders();

    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers });
    }

    const url = new URL(request.url);
    if (request.method !== 'POST' || url.pathname !== '/v1/messages') {
      return jsonResponse(404, { error: { message: 'Not found.' } }, headers);
    }

    if (!env.ANTHROPIC_API_KEY) {
      return jsonResponse(500, { error: { message: 'AI proxy is not configured.' } }, headers);
    }

    const maxBodyBytes = parsePositiveInt(env.MAX_BODY_BYTES, 200000);
    const contentLength = Number.parseInt(request.headers.get('Content-Length') || '0', 10);
    if (Number.isFinite(contentLength) && contentLength > maxBodyBytes) {
      return jsonResponse(413, { error: { message: 'Request body is too large.' } }, headers);
    }

    const rateLimit = parsePositiveInt(env.RATE_LIMIT_PER_HOUR, 20);
    const rate = await checkRateLimit(request, rateLimit);
    if (!rate.allowed) {
      return jsonResponse(
        429,
        { error: { message: 'Rate limit exceeded. Try again later.' } },
        { ...headers, 'Retry-After': '3600' }
      );
    }

    let body;
    try {
      body = await request.json();
    } catch {
      return jsonResponse(400, { error: { message: 'Invalid JSON body.' } }, headers);
    }

    const serialized = JSON.stringify(body || {});
    if (serialized.length > maxBodyBytes) {
      return jsonResponse(413, { error: { message: 'Request body is too large.' } }, headers);
    }

    const system = typeof body.system === 'string' ? body.system : '';
    const messages = body.messages;
    if (!system || !isValidMessages(messages)) {
      return jsonResponse(400, { error: { message: 'system and messages are required.' } }, headers);
    }

    const maxTokens = Math.min(
      Math.max(parsePositiveInt(body.max_tokens, 1024), 16),
      8192
    );
    const stream = Boolean(body.stream);
    const model = String(env.ANTHROPIC_MODEL || 'claude-sonnet-5').trim();

    const anthropicPayload = {
      model,
      max_tokens: maxTokens,
      system,
      messages,
      stream
    };

    const anthropicResponse = await fetch(ANTHROPIC_MESSAGES_URL, {
      method: 'POST',
      headers: {
        'x-api-key': env.ANTHROPIC_API_KEY,
        'anthropic-version': String(env.ANTHROPIC_VERSION || '2023-06-01'),
        'content-type': 'application/json'
      },
      body: JSON.stringify(anthropicPayload)
    });

    if (stream) {
      return new Response(anthropicResponse.body, {
        status: anthropicResponse.status,
        headers: {
          ...headers,
          'Content-Type': anthropicResponse.headers.get('Content-Type') || 'text/event-stream',
          'Cache-Control': 'no-store'
        }
      });
    }

    const responseText = await anthropicResponse.text();
    return new Response(responseText, {
      status: anthropicResponse.status,
      headers: {
        ...headers,
        'Content-Type': 'application/json',
        'Cache-Control': 'no-store'
      }
    });
  }
};
