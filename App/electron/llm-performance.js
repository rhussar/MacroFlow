const os = require('node:os');

const {
  LOCAL_AI_CONTEXT_LENGTH,
  REQUEST_TIMEOUT_MS,
  MAX_CURRENT_CODE_CHARS,
  MAX_PROMPT_CHARS,
  MAX_WORKBOOK_CONTEXT_CHARS,
  MAX_COMPLETION_TOKENS
} = require('./llm-config');

const PROFILE_NAMES = new Set(['low_resource', 'balanced', 'standard']);
const CONTEXT_MODES = new Set(['minimal', 'reduced', 'full']);

const PROFILE_PRESETS = {
  low_resource: {
    contextLength: Math.min(LOCAL_AI_CONTEXT_LENGTH, 2048),
    requestTimeoutMs: Math.max(REQUEST_TIMEOUT_MS, 90000),
    contextFetchTimeoutMs: 1000,
    maxPromptChars: Math.min(MAX_PROMPT_CHARS, 1600),
    maxCurrentCodeChars: Math.min(MAX_CURRENT_CODE_CHARS, 2500),
    maxWorkbookContextChars: Math.min(MAX_WORKBOOK_CONTEXT_CHARS, 700),
    maxCompletionTokens: Math.min(MAX_COMPLETION_TOKENS, 400),
    workbookContextMode: 'minimal',
    includeWorkbookContextForAsk: false,
    enableWorkbookContext: true,
    ollamaMaxLoadedModels: 1,
    ollamaNumParallel: 1,
    ollamaKeepAlive: '30s'
  },
  balanced: {
    contextLength: Math.min(LOCAL_AI_CONTEXT_LENGTH, 4096),
    requestTimeoutMs: Math.max(REQUEST_TIMEOUT_MS, 60000),
    contextFetchTimeoutMs: 1400,
    maxPromptChars: Math.min(MAX_PROMPT_CHARS, 2400),
    maxCurrentCodeChars: Math.min(MAX_CURRENT_CODE_CHARS, 4500),
    maxWorkbookContextChars: Math.min(MAX_WORKBOOK_CONTEXT_CHARS, 1200),
    maxCompletionTokens: Math.min(MAX_COMPLETION_TOKENS, 550),
    workbookContextMode: 'minimal',
    includeWorkbookContextForAsk: false,
    enableWorkbookContext: true,
    ollamaMaxLoadedModels: 1,
    ollamaNumParallel: 1,
    ollamaKeepAlive: '2m'
  },
  standard: {
    contextLength: Math.min(Math.max(LOCAL_AI_CONTEXT_LENGTH, 8192), 8192),
    requestTimeoutMs: Math.max(REQUEST_TIMEOUT_MS, 45000),
    contextFetchTimeoutMs: 2200,
    maxPromptChars: MAX_PROMPT_CHARS,
    maxCurrentCodeChars: MAX_CURRENT_CODE_CHARS,
    maxWorkbookContextChars: MAX_WORKBOOK_CONTEXT_CHARS,
    maxCompletionTokens: MAX_COMPLETION_TOKENS,
    workbookContextMode: 'reduced',
    includeWorkbookContextForAsk: false,
    enableWorkbookContext: true,
    ollamaMaxLoadedModels: 1,
    ollamaNumParallel: 1,
    ollamaKeepAlive: '5m'
  }
};

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

function parseEnumEnv(name, allowedValues, fallback) {
  const rawValue = String(process.env[name] || '').trim().toLowerCase();
  if (allowedValues.has(rawValue)) {
    return rawValue;
  }
  return fallback;
}

function parseStringEnv(name, fallback) {
  const rawValue = String(process.env[name] || '').trim();
  return rawValue || fallback;
}

function toMemoryGb(totalMemoryBytes) {
  return Number((Number(totalMemoryBytes || 0) / (1024 ** 3)).toFixed(1));
}

function detectHardwareTier(options = {}) {
  const totalMemoryBytes = Number(options.totalMemoryBytes || os.totalmem());
  const freeMemoryBytes = Number(options.freeMemoryBytes || os.freemem());
  const cpuCount = Number(options.cpuCount || os.cpus()?.length || 0);
  const totalMemoryGb = toMemoryGb(totalMemoryBytes);
  const freeMemoryGb = toMemoryGb(freeMemoryBytes);
  const memoryPressureRatio = totalMemoryBytes > 0
    ? Number((freeMemoryBytes / totalMemoryBytes).toFixed(2))
    : 1;

  if (
    totalMemoryGb <= 8
    || freeMemoryGb <= 3
    || cpuCount <= 4
    || (totalMemoryGb <= 16 && (freeMemoryGb <= 6 || memoryPressureRatio <= 0.35))
  ) {
    return {
      tier: 'low_resource',
      totalMemoryGb,
      freeMemoryGb,
      cpuCount,
      memoryPressureRatio
    };
  }

  if (
    totalMemoryGb <= 16
    || freeMemoryGb <= 8
    || cpuCount <= 8
    || memoryPressureRatio <= 0.5
  ) {
    return {
      tier: 'balanced',
      totalMemoryGb,
      freeMemoryGb,
      cpuCount,
      memoryPressureRatio
    };
  }

  return {
    tier: 'standard',
    totalMemoryGb,
    freeMemoryGb,
    cpuCount,
    memoryPressureRatio
  };
}

function resolveLlmPerformanceProfile(options = {}) {
  const hardware = detectHardwareTier(options);
  const selectedProfileName = parseEnumEnv('MACROFLOW_LLM_PROFILE', PROFILE_NAMES, hardware.tier);
  const preset = PROFILE_PRESETS[selectedProfileName] || PROFILE_PRESETS.balanced;

  return {
    name: selectedProfileName,
    hardwareTier: hardware.tier,
    totalMemoryGb: hardware.totalMemoryGb,
    freeMemoryGb: hardware.freeMemoryGb,
    cpuCount: hardware.cpuCount,
    memoryPressureRatio: hardware.memoryPressureRatio,
    contextLength: parseNumberEnv('MACROFLOW_LLM_CONTEXT_LENGTH', preset.contextLength),
    requestTimeoutMs: parseNumberEnv('MACROFLOW_LLM_REQUEST_TIMEOUT_MS', preset.requestTimeoutMs),
    contextFetchTimeoutMs: parseNumberEnv('MACROFLOW_LLM_CONTEXT_FETCH_TIMEOUT_MS', preset.contextFetchTimeoutMs),
    maxPromptChars: parseNumberEnv('MACROFLOW_LLM_MAX_PROMPT_CHARS', preset.maxPromptChars),
    maxCurrentCodeChars: parseNumberEnv('MACROFLOW_LLM_MAX_CURRENT_CODE_CHARS', preset.maxCurrentCodeChars),
    maxWorkbookContextChars: parseNumberEnv('MACROFLOW_LLM_MAX_WORKBOOK_CONTEXT_CHARS', preset.maxWorkbookContextChars),
    maxCompletionTokens: parseNumberEnv('MACROFLOW_LLM_MAX_COMPLETION_TOKENS', preset.maxCompletionTokens),
    workbookContextMode: parseEnumEnv(
      'MACROFLOW_LLM_WORKBOOK_CONTEXT_MODE',
      CONTEXT_MODES,
      preset.workbookContextMode
    ),
    includeWorkbookContextForAsk: parseBooleanEnv(
      'MACROFLOW_LLM_CONTEXT_FOR_ASK',
      preset.includeWorkbookContextForAsk
    ),
    enableWorkbookContext: parseBooleanEnv(
      'MACROFLOW_LLM_ENABLE_WORKBOOK_CONTEXT',
      preset.enableWorkbookContext
    ),
    ollamaMaxLoadedModels: parseNumberEnv(
      'MACROFLOW_OLLAMA_MAX_LOADED_MODELS',
      preset.ollamaMaxLoadedModels
    ),
    ollamaNumParallel: parseNumberEnv(
      'MACROFLOW_OLLAMA_NUM_PARALLEL',
      preset.ollamaNumParallel
    ),
    ollamaKeepAlive: parseStringEnv(
      'MACROFLOW_OLLAMA_KEEP_ALIVE',
      preset.ollamaKeepAlive
    )
  };
}

module.exports = {
  PROFILE_PRESETS,
  detectHardwareTier,
  resolveLlmPerformanceProfile
};
