# MacroFlow AI proxy

Cloudflare Worker that holds the Anthropic API key and forwards Create-tab requests to Claude Sonnet 5.

## Deploy

```bash
cd worker
npx wrangler login
npx wrangler secret put ANTHROPIC_API_KEY
npx wrangler deploy
```

The Worker URL is typically:

`https://macroflow-ai.<your-account>.workers.dev`

Default desktop app URL: `https://macroflow-ai.macroflowai.workers.dev`.
Override with `MACROFLOW_AI_URL` if your Worker hostname differs.

## Rate limit

20 requests per client IP per hour. Raise or lower `RATE_LIMIT_PER_HOUR` in `wrangler.toml`.
