/**
 * Local AI configuration for VBA generation.
 */

function parseBooleanEnv(name, fallback = false) {
  const rawValue = String(process.env[name] || '').trim().toLowerCase();
  if (!rawValue) {
    return fallback;
  }
  return ['1', 'true', 'yes', 'on'].includes(rawValue);
}

function parseNumberEnv(name, fallback) {
  const rawValue = Number(process.env[name]);
  if (!Number.isFinite(rawValue) || rawValue <= 0) {
    return fallback;
  }
  return Math.round(rawValue);
}

const LOCAL_AI_PROVIDER = 'ollama';
const LOCAL_AI_MODEL = String(process.env.MACROFLOW_LLM_MODEL || 'qwen2.5-coder:3b').trim();
const LOCAL_AI_BASE_URL = String(process.env.MACROFLOW_LLM_BASE_URL || 'http://127.0.0.1:11434').trim();
const LOCAL_AI_CONTEXT_LENGTH = parseNumberEnv('MACROFLOW_LLM_CONTEXT_LENGTH', 8192);
const ALLOW_EXTERNAL_OLLAMA = parseBooleanEnv('MACROFLOW_ALLOW_EXTERNAL_OLLAMA', false);
const OLLAMA_RUNTIME_VERSION = String(process.env.MACROFLOW_OLLAMA_RUNTIME_VERSION || '0.18.2').trim();
const OLLAMA_RELEASE_BASE_URL = `https://github.com/ollama/ollama/releases/download/v${OLLAMA_RUNTIME_VERSION}`;
const OLLAMA_WINDOWS_ZIP_URL = String(
  process.env.MACROFLOW_OLLAMA_WINDOWS_ZIP_URL || `${OLLAMA_RELEASE_BASE_URL}/ollama-windows-amd64.zip`
).trim();
const OLLAMA_WINDOWS_ZIP_SHA256 = String(
  process.env.MACROFLOW_OLLAMA_WINDOWS_ZIP_SHA256 || '7752b897afbe1852d6eb210e1c69263cfeefd39f8595f7489c0e2ceefecab919'
).trim().toLowerCase();
const OLLAMA_WINDOWS_ROCM_ZIP_URL = String(
  process.env.MACROFLOW_OLLAMA_WINDOWS_ROCM_ZIP_URL || `${OLLAMA_RELEASE_BASE_URL}/ollama-windows-amd64-rocm.zip`
).trim();
const OLLAMA_WINDOWS_ROCM_ZIP_SHA256 = String(
  process.env.MACROFLOW_OLLAMA_WINDOWS_ROCM_ZIP_SHA256 || '6e2100abfa541b0232fecabb6ad20eac74b22550cacd05e40ce2a8e47a63cf4c'
).trim().toLowerCase();
const REQUEST_TIMEOUT_MS = 30000;
const HEALTHCHECK_TIMEOUT_MS = 1500;
const SETUP_SERVER_TIMEOUT_MS = 60 * 1000;
const MAX_CURRENT_CODE_CHARS = 8000;
const MAX_PROMPT_CHARS = 3000;
const MAX_WORKBOOK_CONTEXT_CHARS = 2200;
const MAX_COMPLETION_TOKENS = 900;
const TEMPERATURE = 0.2;

module.exports = {
  LOCAL_AI_PROVIDER,
  LOCAL_AI_MODEL,
  LOCAL_AI_BASE_URL,
  LOCAL_AI_CONTEXT_LENGTH,
  ALLOW_EXTERNAL_OLLAMA,
  OLLAMA_RUNTIME_VERSION,
  OLLAMA_WINDOWS_ZIP_URL,
  OLLAMA_WINDOWS_ZIP_SHA256,
  OLLAMA_WINDOWS_ROCM_ZIP_URL,
  OLLAMA_WINDOWS_ROCM_ZIP_SHA256,
  REQUEST_TIMEOUT_MS,
  HEALTHCHECK_TIMEOUT_MS,
  SETUP_SERVER_TIMEOUT_MS,
  MAX_CURRENT_CODE_CHARS,
  MAX_PROMPT_CHARS,
  MAX_WORKBOOK_CONTEXT_CHARS,
  MAX_COMPLETION_TOKENS,
  TEMPERATURE
};
