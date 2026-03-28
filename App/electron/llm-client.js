const logger = require('./logger');
const {
  LOCAL_AI_MODEL,
  LOCAL_AI_BASE_URL,
  TEMPERATURE
} = require('./llm-config');
const { requestText } = require('./http-client');
const localAiManager = require('./local-ai-manager');
const { normalizeAiIntent, buildUserPrompt, getSystemPrompt } = require('./llm-prompts');
const { resolveWorkbookPromptContext } = require('./llm-context');
const { resolveLlmPerformanceProfile } = require('./llm-performance');

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
  const apiMessage = responseBody?.error?.message;
  if (apiMessage && typeof apiMessage === 'string') {
    return apiMessage.trim();
  }
  return '';
}

function mapGenerateError(error, statusCode = 0) {
  const rawMessage = String(error?.message || '').trim();
  const messageLower = rawMessage.toLowerCase();

  if (messageLower.includes('model') && messageLower.includes('not found')) {
    return {
      reason: 'AI_MODEL_MISSING',
      message: 'The local AI model is not installed yet.'
    };
  }

  if (statusCode === 404) {
    return {
      reason: 'AI_MODEL_MISSING',
      message: rawMessage || 'The local AI model is not installed yet.'
    };
  }

  if (messageLower.includes('timeout') || messageLower.includes('timed out') || error?.code === 'ETIMEDOUT') {
    return {
      reason: 'AI_TIMEOUT',
      message: 'Local AI timed out. Please try again.'
    };
  }

  if (error?.code === 'ENOTFOUND' || error?.code === 'ECONNRESET' || error?.code === 'ECONNREFUSED') {
    return {
      reason: 'AI_NOT_READY',
      message: 'Local AI is not running yet. Finish setup and retry.'
    };
  }

  if (statusCode >= 400) {
    return {
      reason: 'AI_INVALID_RESPONSE',
      message: rawMessage || `Local AI request failed (HTTP ${statusCode}).`
    };
  }

  return {
    reason: 'AI_INVALID_RESPONSE',
    message: rawMessage || 'Local AI returned an invalid response.'
  };
}

function createRequestId() {
  return `ai_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`;
}

function normalizeContextPayload(contextPayload) {
  if (typeof contextPayload === 'string') {
    return {
      text: contextPayload,
      meta: {}
    };
  }

  if (contextPayload && typeof contextPayload === 'object') {
    return {
      text: toSafeString(contextPayload.text),
      meta: contextPayload.meta && typeof contextPayload.meta === 'object' ? contextPayload.meta : {}
    };
  }

  return {
    text: '',
    meta: {}
  };
}

function buildDiagnostics({
  performanceProfile,
  contextMeta,
  workbookContextChars,
  requestDurationMs = 0,
  totalDurationMs = 0
}) {
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

async function generateVba(
  {
    prompt = '',
    intent = '',
    workbookName = '',
    workbookPath = '',
    moduleName = '',
    sheetName = '',
    currentCode = '',
    includeCurrentCode = false
  } = {},
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
    return {
      success: false,
      reason: 'AI_INVALID_PROMPT',
      message: 'Prompt is required to generate VBA.'
    };
  }

  const statusImpl = typeof dependencies.statusImpl === 'function'
    ? dependencies.statusImpl
    : localAiManager.getStatus;
  const aiStatus = await Promise.resolve(statusImpl());

  if (!aiStatus?.runtimeInstalled) {
    return {
      success: false,
      reason: 'AI_RUNTIME_MISSING',
      message: 'Local AI runtime is not installed yet.'
    };
  }

  if (!aiStatus?.serverReachable) {
    return {
      success: false,
      reason: 'AI_NOT_READY',
      message: 'Local AI is not running yet. Finish setup and retry.'
    };
  }

  if (!aiStatus?.modelInstalled) {
    return {
      success: false,
      reason: 'AI_MODEL_MISSING',
      message: 'The local AI model is not installed yet.'
    };
  }

  const selectedModel = toSafeString(aiStatus?.model) || LOCAL_AI_MODEL;
  const contextImpl = typeof dependencies.contextImpl === 'function'
    ? dependencies.contextImpl
    : resolveWorkbookPromptContext;
  let workbookContext = '';
  let contextMeta = {};

  try {
    const contextPayload = await Promise.resolve(
      contextImpl({
        intent: normalizedIntent,
        workbookName,
        workbookPath,
        moduleName,
        sheetName
      }, {
        performanceProfile
      })
    );
    const normalizedContextPayload = normalizeContextPayload(contextPayload);
    workbookContext = normalizedContextPayload.text;
    contextMeta = normalizedContextPayload.meta;
  } catch (error) {
    logger.warn('[AI] workbook context unavailable', {
      requestId,
      intent: normalizedIntent,
      workbookName,
      workbookPath,
      message: String(error?.message || 'Unknown workbook context error')
    });
  }

  const payload = {
    model: selectedModel,
    temperature: TEMPERATURE,
    max_tokens: performanceProfile.maxCompletionTokens,
    messages: [
      {
        role: 'system',
        content: getSystemPrompt(normalizedIntent, {
          includeCurrentCode: shouldIncludeCurrentCode
        })
      },
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
    ]
  };

  logger.debug('[AI] generate-vba request', {
    requestId,
    provider: aiStatus?.provider || 'ollama',
    model: selectedModel,
    intent: normalizedIntent,
    promptChars: promptText.length,
    requestTimeoutMs: performanceProfile.requestTimeoutMs,
    maxCompletionTokens: performanceProfile.maxCompletionTokens,
    performanceProfile: performanceProfile.name,
    includeCurrentCode: shouldIncludeCurrentCode,
    currentCodeChars: shouldIncludeCurrentCode ? String(currentCode || '').length : 0,
    workbookContextChars: String(workbookContext || '').length,
    workbookContextMeta: contextMeta
  });

  const requestImpl = typeof dependencies.requestImpl === 'function'
    ? dependencies.requestImpl
    : requestText;
  let requestStartedAt = 0;

  try {
    const baseUrl = String(aiStatus?.baseUrl || LOCAL_AI_BASE_URL).trim();
    const url = `${baseUrl.replace(/\/+$/, '')}/v1/chat/completions`;
    const body = JSON.stringify(payload);
    requestStartedAt = Date.now();

    const response = await requestImpl({
      url,
      timeoutMs: performanceProfile.requestTimeoutMs,
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(body)
      },
      body,
      method: 'POST'
    });

    const statusCode = Number(response?.statusCode) || 0;
    const bodyText = String(response?.bodyText || '').trim();
    let parsedBody = {};

    try {
      parsedBody = bodyText ? JSON.parse(bodyText) : {};
    } catch {
      throw new Error('AI_INVALID_RESPONSE: Local AI returned invalid JSON.');
    }

    if (statusCode >= 400) {
      const apiMessage = getErrorMessageFromResponse(parsedBody);
      const mapped = mapGenerateError(new Error(apiMessage || `Local AI request failed (HTTP ${statusCode}).`), statusCode);
      const diagnostics = buildDiagnostics({
        performanceProfile,
        contextMeta,
        workbookContextChars: String(workbookContext || '').length,
        requestDurationMs: Date.now() - requestStartedAt,
        totalDurationMs: Date.now() - startedAt
      });
      return {
        success: false,
        reason: mapped.reason,
        message: mapped.message,
        diagnostics
      };
    }

    const rawContent = extractTextContent(parsedBody?.choices?.[0]?.message?.content);
    const normalizedContent = normalizeLineEndings(rawContent).trim();
    const requestDurationMs = Date.now() - requestStartedAt;
    const diagnostics = buildDiagnostics({
      performanceProfile,
      contextMeta,
      workbookContextChars: String(workbookContext || '').length,
      requestDurationMs,
      totalDurationMs: Date.now() - startedAt
    });

    if (normalizedIntent === 'ask') {
      if (!normalizedContent) {
        return {
          success: false,
          reason: 'AI_INVALID_RESPONSE',
          message: 'Local AI returned an empty answer.',
          diagnostics
        };
      }

      const usage = parsedBody?.usage && typeof parsedBody.usage === 'object'
        ? {
            promptTokens: Number(parsedBody.usage.prompt_tokens) || 0,
            completionTokens: Number(parsedBody.usage.completion_tokens) || 0,
            totalTokens: Number(parsedBody.usage.total_tokens) || 0
          }
        : undefined;

      logger.info('[AI] generate-vba success', {
        requestId,
        model: selectedModel,
        intent: normalizedIntent,
        statusCode,
        durationMs: diagnostics.timings.totalMs,
        usage,
        diagnostics
      });

      return {
        success: true,
        content: normalizedContent,
        intent: normalizedIntent,
        model: selectedModel,
        usage,
        diagnostics
      };
    }

    const code = extractCodeBlock(rawContent);

    if (!code || !hasVbaProcedure(code)) {
      return {
        success: false,
        reason: 'AI_INVALID_RESPONSE',
        message: 'Local AI did not return valid VBA code.',
        diagnostics
      };
    }

    const usage = parsedBody?.usage && typeof parsedBody.usage === 'object'
      ? {
          promptTokens: Number(parsedBody.usage.prompt_tokens) || 0,
          completionTokens: Number(parsedBody.usage.completion_tokens) || 0,
          totalTokens: Number(parsedBody.usage.total_tokens) || 0
        }
      : undefined;

    logger.info('[AI] generate-vba success', {
      requestId,
      model: selectedModel,
      intent: normalizedIntent,
      statusCode,
      durationMs: diagnostics.timings.totalMs,
      usage,
      diagnostics
    });

    return {
      success: true,
      code,
      intent: normalizedIntent,
      model: selectedModel,
      usage,
      diagnostics
    };
  } catch (error) {
    const mapped = mapGenerateError(error);
    const diagnostics = buildDiagnostics({
      performanceProfile,
      contextMeta,
      workbookContextChars: String(workbookContext || '').length,
      requestDurationMs: requestStartedAt > 0 ? Date.now() - requestStartedAt : 0,
      totalDurationMs: Date.now() - startedAt
    });
    logger.warn('[AI] generate-vba failed', {
      requestId,
      reason: mapped.reason,
      durationMs: diagnostics.timings.totalMs,
      message: mapped.message,
      diagnostics
    });
    return {
      success: false,
      reason: mapped.reason,
      message: mapped.message,
      diagnostics
    };
  }
}

module.exports = {
  generateVba
};
