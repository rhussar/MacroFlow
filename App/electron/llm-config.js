/**
 * Cloud AI configuration for VBA generation (Claude Sonnet 5 via Cloudflare Worker).
 */

const AI_PROVIDER = 'anthropic';
const AI_MODEL = String(process.env.MACROFLOW_LLM_MODEL || 'claude-sonnet-5').trim();
const AI_PROXY_URL = String(
  process.env.MACROFLOW_AI_URL || 'https://macroflow-ai.macroflowai.workers.dev'
).trim().replace(/\/+$/, '');
const LOCAL_AI_CONTEXT_LENGTH = 8192;
const REQUEST_TIMEOUT_MS = 60000;
const MAX_CURRENT_CODE_CHARS = 8000;
const MAX_PROMPT_CHARS = 3000;
const MAX_WORKBOOK_CONTEXT_CHARS = 2200;
const MAX_COMPLETION_TOKENS = 2048;
const TEMPERATURE = 0.2;

function getCloudAiStatus() {
  return {
    success: true,
    provider: AI_PROVIDER,
    model: AI_MODEL,
    ready: true,
    needsSetup: false,
    setupInProgress: false,
    removeInProgress: false,
    runtimeInstalled: true,
    serverReachable: true,
    modelInstalled: true,
    stage: 'ready',
    statusText: 'Claude Sonnet 5 is ready.',
    progress: null,
    lastError: ''
  };
}

module.exports = {
  AI_PROVIDER,
  AI_MODEL,
  AI_PROXY_URL,
  LOCAL_AI_CONTEXT_LENGTH,
  REQUEST_TIMEOUT_MS,
  MAX_CURRENT_CODE_CHARS,
  MAX_PROMPT_CHARS,
  MAX_WORKBOOK_CONTEXT_CHARS,
  MAX_COMPLETION_TOKENS,
  TEMPERATURE,
  getCloudAiStatus
};
