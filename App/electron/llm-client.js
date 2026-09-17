const http = require('node:http');
const https = require('node:https');
const logger = require('./logger');
const {
  AI_MODEL,
  AI_PROVIDER,
  AI_PROXY_URL,
  TEMPERATURE
} = require('./llm-config');
const { requestText } = require('./http-client');
const { normalizeAiIntent, buildUserPrompt, getSystemPrompt } = require('./llm-prompts');
const { resolveWorkbookPromptContext } = require('./llm-context');
const { resolveLlmPerformanceProfile } = require('./llm-performance');
const { matchHardcodedResponse } = require('./llm-hardcoded');

function toSafeString(value) {
  return String(value || '').trim();
}

function normalizeLineEndings(value) {
  return String(value || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
}

function truncate(value, maxChars) {
  const normalized = String(value || '');
  if (normalized.length <= maxChars) {
    return normalized;
  }
  return normalized.slice(0, maxChars);
}

function extractTextContent(content) {
  if (typeof content === 'string') {
    return content;
  }
  if (Array.isArray(content)) {
    return content
      .map((part) => {
        if (typeof part === 'string') {
          return part;
        }
        if (part && typeof part === 'object' && typeof part.text === 'string') {
          return part.text;
        }
        return '';
      })
      .join('\n');
  }
  return '';
}

function extractAnthropicText(parsedBody) {
  if (Array.isArray(parsedBody?.content)) {
    return extractTextContent(parsedBody.content);
  }
  return extractTextContent(parsedBody?.choices?.[0]?.message?.content);
}

function extractCodeBlock(text) {
  const normalized = normalizeLineEndings(text);
  const fencedMatch = normalized.match(/```(?:vba|vb)?\s*([\s\S]*?)```/i);
  if (fencedMatch && fencedMatch[1]) {
    return normalizeLineEndings(fencedMatch[1]).trim();
  }
  return normalized.trim();
}

function hasVbaProcedure(code) {
  return /\b(Sub|Function)\b/i.test(String(code || ''));
}

function getErrorMessageFromResponse(responseBody) {
  const apiMessage = responseBody?.error?.message || responseBody?.error?.type;
  if (apiMessage && typeof apiMessage === 'string') {
    return apiMessage.trim();
  }
  return '';
}

function mapGenerateError(error, statusCode = 0) {
  const rawMessage = String(error?.message || '').trim();
  const messageLower = rawMessage.toLowerCase();

  if (statusCode === 429 || messageLower.includes('rate limit')) {
    return { reason: 'AI_RATE_LIMITED', message: 'AI rate limit reached. Please try again later.' };
  }
  if (statusCode === 401 || statusCode === 403) {
    return { reason: 'AI_UNAUTHORIZED', message: 'AI proxy rejected the request. Check the Cloudflare Worker configuration.' };
  }
  if (messageLower.includes('timeout') || messageLower.includes('timed out') || error?.code === 'ETIMEDOUT') {
    return { reason: 'AI_TIMEOUT', message: 'AI request timed out. Please try again.' };
  }
  if (error?.code === 'ENOTFOUND' || error?.code === 'ECONNRESET' || error?.code === 'ECONNREFUSED') {
    return { reason: 'AI_NOT_READY', message: 'Could not reach the AI service. Check your internet connection and try again.' };
  }
  if (statusCode >= 400) {
    return { reason: 'AI_INVALID_RESPONSE', message: rawMessage || `AI request failed (HTTP ${statusCode}).` };
  }
  return { reason: 'AI_INVALID_RESPONSE', message: rawMessage || 'AI returned an invalid response.' };
}

function createRequestId() {
  return `ai_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`;
}

function normalizeContextPayload(contextPayload) {
  if (typeof contextPayload === 'string') {
    return { text: contextPayload, meta: {} };
  }
  if (contextPayload && typeof contextPayload === 'object') {
    return {
      text: toSafeString(contextPayload.text),
      meta: contextPayload.meta && typeof contextPayload.meta === 'object' ? contextPayload.meta : {}
    };
  }
  return { text: '', meta: {} };
}

function buildDiagnostics({ performanceProfile, contextMeta, workbookContextChars, requestDurationMs = 0, totalDurationMs = 0 }) {
  return {
    profile: performanceProfile.name,
    hardwareTier: performanceProfile.hardwareTier,
    totalMemoryGb: performanceProfile.totalMemoryGb,
    freeMemoryGb: performanceProfile.freeMemoryGb,
    cpuCount: performanceProfile.cpuCount,
    memoryPressureRatio: performanceProfile.memoryPressureRatio,
    timings: {
      workbookContextMs: Number(contextMeta?.durationMs) || 0,
      llmRequestMs: Number(requestDurationMs) || 0,
      totalMs: Number(totalDurationMs) || 0
    },
    context: {
      included: Boolean(contextMeta?.included),
      chars: Number(workbookContextChars) || 0,
      cacheHit: Boolean(contextMeta?.cacheHit),
      timedOut: Boolean(contextMeta?.timedOut),
      mode: toSafeString(contextMeta?.mode),
      skippedReason: toSafeString(contextMeta?.skippedReason)
    }
  };
}

function extractUsage(parsedBody) {
  const usage = parsedBody?.usage;
  if (!usage || typeof usage !== 'object') {
    return undefined;
  }

  const promptTokens = Number(usage.input_tokens ?? usage.prompt_tokens) || 0;
  const completionTokens = Number(usage.output_tokens ?? usage.completion_tokens) || 0;
  const totalTokens = Number(usage.total_tokens) || (promptTokens + completionTokens);
  if (!promptTokens && !completionTokens && !totalTokens) {
    return undefined;
  }
  return { promptTokens, completionTokens, totalTokens };
}

async function prepareGeneration(
  { prompt = '', intent = '', workbookName = '', workbookPath = '', moduleName = '', sheetName = '', currentCode = '', includeCurrentCode = false } = {},
  dependencies = {}
) {
  const requestId = createRequestId();
  const startedAt = Date.now();
  const performanceProfile =
    (typeof dependencies.performanceProfile === 'function'
      ? dependencies.performanceProfile()
      : dependencies.performanceProfile)
    || resolveLlmPerformanceProfile();
  const promptText = truncate(String(prompt || ''), performanceProfile.maxPromptChars);
  const normalizedIntent = normalizeAiIntent(intent, { includeCurrentCode: Boolean(includeCurrentCode) });
  const shouldIncludeCurrentCode = normalizedIntent === 'edit'
    ? Boolean(includeCurrentCode || toSafeString(currentCode))
    : Boolean(includeCurrentCode);

  if (!promptText.trim()) {
    return { earlyReturn: { success: false, reason: 'AI_INVALID_PROMPT', message: 'Prompt is required to generate VBA.' } };
  }

  const selectedModel = toSafeString(dependencies.model) || AI_MODEL;
  const contextImpl = typeof dependencies.contextImpl === 'function'
    ? dependencies.contextImpl
    : resolveWorkbookPromptContext;
  let workbookContext = '';
  let contextMeta = {};

  try {
    const contextPayload = await Promise.resolve(
      contextImpl({ intent: normalizedIntent, workbookName, workbookPath, moduleName, sheetName }, { performanceProfile })
    );
    const normalized = normalizeContextPayload(contextPayload);
    workbookContext = normalized.text;
    contextMeta = normalized.meta;
  } catch (error) {
    logger.warn('[AI] workbook context unavailable', {
      requestId, intent: normalizedIntent, workbookName, workbookPath,
      message: String(error?.message || 'Unknown workbook context error')
    });
  }

  const payload = {
    system: getSystemPrompt(normalizedIntent, { includeCurrentCode: shouldIncludeCurrentCode }),
    messages: [
      {
        role: 'user',
        content: buildUserPrompt({
          intent: normalizedIntent,
          prompt: promptText,
          workbookName,
          moduleName,
          workbookContext,
          currentCode,
          includeCurrentCode: shouldIncludeCurrentCode,
          limits: performanceProfile
        })
      }
    ],
    max_tokens: normalizedIntent === 'ask'
      ? Math.min(performanceProfile.maxCompletionTokens, 80)
      : performanceProfile.maxCompletionTokens,
    temperature: TEMPERATURE
  };

  const baseUrl = toSafeString(dependencies.baseUrl || AI_PROXY_URL).replace(/\/+$/, '');

  logger.debug('[AI] generate-vba request', {
    requestId, provider: AI_PROVIDER, model: selectedModel,
    intent: normalizedIntent, promptChars: promptText.length,
    requestTimeoutMs: performanceProfile.requestTimeoutMs,
    maxCompletionTokens: performanceProfile.maxCompletionTokens,
    performanceProfile: performanceProfile.name,
    includeCurrentCode: shouldIncludeCurrentCode,
    currentCodeChars: shouldIncludeCurrentCode ? String(currentCode || '').length : 0,
    workbookContextChars: String(workbookContext || '').length,
    workbookContextMeta: contextMeta
  });

  return {
    requestId, startedAt, performanceProfile, normalizedIntent, shouldIncludeCurrentCode,
    selectedModel, payload, baseUrl, workbookContext, contextMeta
  };
}

function buildResult({ rawContent, normalizedIntent, selectedModel, requestId, performanceProfile, contextMeta, workbookContext, startedAt, requestStartedAt, usage }) {
  const normalizedContent = normalizeLineEndings(rawContent).trim();
  const diagnostics = buildDiagnostics({
    performanceProfile, contextMeta,
    workbookContextChars: String(workbookContext || '').length,
    requestDurationMs: Date.now() - requestStartedAt,
    totalDurationMs: Date.now() - startedAt
  });

  if (normalizedIntent === 'ask') {
    if (!normalizedContent) {
      return { success: false, reason: 'AI_INVALID_RESPONSE', message: 'AI returned an empty answer.', diagnostics };
    }
    logger.info('[AI] generate-vba success', { requestId, model: selectedModel, intent: normalizedIntent, durationMs: diagnostics.timings.totalMs, diagnostics });
    return { success: true, content: normalizedContent, intent: normalizedIntent, model: selectedModel, usage, diagnostics };
  }

  const code = extractCodeBlock(rawContent);
  if (!code || !hasVbaProcedure(code)) {
    return { success: false, reason: 'AI_INVALID_RESPONSE', message: 'AI did not return valid VBA code.', diagnostics };
  }

  logger.info('[AI] generate-vba success', { requestId, model: selectedModel, intent: normalizedIntent, durationMs: diagnostics.timings.totalMs, diagnostics });
  return { success: true, code, intent: normalizedIntent, model: selectedModel, usage, diagnostics };
}

function extractStreamDelta(parsed) {
  if (parsed?.type === 'content_block_delta' && parsed?.delta?.type === 'text_delta') {
    return String(parsed.delta.text || '');
  }
  if (parsed?.delta?.text) {
    return String(parsed.delta.text);
  }
  return String(parsed?.choices?.[0]?.delta?.content || '');
}

async function generateVba(params = {}, dependencies = {}) {
  const hardcoded = await matchHardcodedResponse(params);
  if (hardcoded) return hardcoded;

  const prep = await prepareGeneration(params, dependencies);
  if (prep.earlyReturn) return prep.earlyReturn;

  const { requestId, startedAt, performanceProfile, normalizedIntent, selectedModel, payload, baseUrl, workbookContext, contextMeta } = prep;
  const requestImpl = typeof dependencies.requestImpl === 'function' ? dependencies.requestImpl : requestText;
  const requestStartedAt = Date.now();

  try {
    const url = `${baseUrl}/v1/messages`;
    const body = JSON.stringify(payload);

    const response = await requestImpl({
      url, timeoutMs: performanceProfile.requestTimeoutMs,
      headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) },
      body, method: 'POST'
    });

    const statusCode = Number(response?.statusCode) || 0;
    const bodyText = String(response?.bodyText || '').trim();
    let parsedBody = {};

    try {
      parsedBody = bodyText ? JSON.parse(bodyText) : {};
    } catch {
      throw new Error('AI_INVALID_RESPONSE: AI returned invalid JSON.');
    }

    if (statusCode >= 400) {
      const apiMessage = getErrorMessageFromResponse(parsedBody);
      const mapped = mapGenerateError(new Error(apiMessage || `AI request failed (HTTP ${statusCode}).`), statusCode);
      const diagnostics = buildDiagnostics({ performanceProfile, contextMeta, workbookContextChars: String(workbookContext || '').length, requestDurationMs: Date.now() - requestStartedAt, totalDurationMs: Date.now() - startedAt });
      return { success: false, reason: mapped.reason, message: mapped.message, diagnostics };
    }

    const rawContent = extractAnthropicText(parsedBody);
    return buildResult({
      rawContent,
      normalizedIntent,
      selectedModel,
      requestId,
      performanceProfile,
      contextMeta,
      workbookContext,
      startedAt,
      requestStartedAt,
      usage: extractUsage(parsedBody)
    });
  } catch (error) {
    const mapped = mapGenerateError(error);
    const diagnostics = buildDiagnostics({ performanceProfile, contextMeta, workbookContextChars: String(workbookContext || '').length, requestDurationMs: Date.now() - requestStartedAt, totalDurationMs: Date.now() - startedAt });
    logger.warn('[AI] generate-vba failed', { requestId, reason: mapped.reason, durationMs: diagnostics.timings.totalMs, message: mapped.message, diagnostics });
    return { success: false, reason: mapped.reason, message: mapped.message, diagnostics };
  }
}

async function generateVbaStream(params = {}, dependencies = {}, onToken) {
  const hardcoded = await matchHardcodedResponse(params, onToken);
  if (hardcoded) return hardcoded;

  const prep = await prepareGeneration(params, dependencies);
  if (prep.earlyReturn) return prep.earlyReturn;

  const { requestId, startedAt, performanceProfile, normalizedIntent, selectedModel, payload, baseUrl, workbookContext, contextMeta } = prep;
  const requestStartedAt = Date.now();

  try {
    const url = `${baseUrl}/v1/messages`;
    const streamPayload = { ...payload, stream: true };
    const body = JSON.stringify(streamPayload);
    let fullContent = '';
    let usage;

    await new Promise((resolve, reject) => {
      const transport = url.startsWith('https:') ? https : http;
      const request = transport.request(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) }
      }, (response) => {
        const statusCode = Number(response.statusCode) || 0;
        if (statusCode >= 400) {
          let errorText = '';
          response.setEncoding('utf8');
          response.on('data', (chunk) => { errorText += chunk; });
          response.on('end', () => {
            let parsed = {};
            try {
              parsed = errorText ? JSON.parse(errorText) : {};
            } catch {
              parsed = {};
            }
            const apiMessage = getErrorMessageFromResponse(parsed);
            const mapped = mapGenerateError(new Error(apiMessage || errorText || `HTTP ${statusCode}`), statusCode);
            reject(Object.assign(new Error(mapped.message), { mapped, statusCode }));
          });
          return;
        }

        response.setEncoding('utf8');
        let buffer = '';

        response.on('data', (chunk) => {
          buffer += chunk;
          const lines = buffer.split('\n');
          buffer = lines.pop() || '';

          for (const line of lines) {
            const trimmed = line.trim();
            if (!trimmed || trimmed === 'data: [DONE]' || trimmed.startsWith('event:')) continue;
            const jsonStr = trimmed.startsWith('data: ') ? trimmed.slice(6) : trimmed;
            try {
              const parsed = JSON.parse(jsonStr);
              const delta = extractStreamDelta(parsed);
              if (delta) {
                fullContent += delta;
                if (typeof onToken === 'function') onToken(delta);
              }
              const nextUsage = extractUsage(parsed);
              if (nextUsage) {
                usage = nextUsage;
              }
            } catch {
              // Ignore malformed SSE chunks.
            }
          }
        });

        response.on('end', resolve);
      });

      request.setTimeout(performanceProfile.requestTimeoutMs, () => {
        request.destroy(new Error('request timed out'));
      });
      request.on('error', reject);
      request.write(body);
      request.end();
    });

    return buildResult({
      rawContent: fullContent,
      normalizedIntent,
      selectedModel,
      requestId,
      performanceProfile,
      contextMeta,
      workbookContext,
      startedAt,
      requestStartedAt,
      usage
    });
  } catch (error) {
    const mapped = error?.mapped || mapGenerateError(error, error?.statusCode);
    const diagnostics = buildDiagnostics({ performanceProfile, contextMeta, workbookContextChars: String(workbookContext || '').length, requestDurationMs: Date.now() - requestStartedAt, totalDurationMs: Date.now() - startedAt });
    logger.warn('[AI] generate-vba-stream failed', { requestId, reason: mapped.reason, durationMs: diagnostics.timings.totalMs, message: mapped.message, diagnostics });
    return { success: false, reason: mapped.reason, message: mapped.message, diagnostics };
  }
}

module.exports = {
  generateVba,
  generateVbaStream
};
