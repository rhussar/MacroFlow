const logger = require('./logger');
const {
  LOCAL_AI_MODEL,
  LOCAL_AI_BASE_URL,
  REQUEST_TIMEOUT_MS,
  MAX_PROMPT_CHARS,
  MAX_COMPLETION_TOKENS,
  TEMPERATURE
} = require('./llm-config');
const { requestText } = require('./http-client');
const localAiManager = require('./local-ai-manager');
const { normalizeAiIntent, buildUserPrompt, getSystemPrompt } = require('./llm-prompts');
const { resolveWorkbookPromptContext } = require('./llm-context');

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
  const promptText = truncate(String(prompt || ''), MAX_PROMPT_CHARS);
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

  try {
    workbookContext = await Promise.resolve(
      contextImpl({
        intent: normalizedIntent,
        workbookName,
        workbookPath,
        moduleName,
        sheetName
      })
    );
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
    max_tokens: MAX_COMPLETION_TOKENS,
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
          includeCurrentCode: shouldIncludeCurrentCode
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
    includeCurrentCode: shouldIncludeCurrentCode,
    currentCodeChars: shouldIncludeCurrentCode ? String(currentCode || '').length : 0,
    workbookContextChars: String(workbookContext || '').length
  });

  const requestImpl = typeof dependencies.requestImpl === 'function'
    ? dependencies.requestImpl
    : requestText;

  try {
    const url = `${LOCAL_AI_BASE_URL.replace(/\/+$/, '')}/v1/chat/completions`;
    const body = JSON.stringify(payload);

    const response = await requestImpl({
      url,
      timeoutMs: REQUEST_TIMEOUT_MS,
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
      return {
        success: false,
        reason: mapped.reason,
        message: mapped.message
      };
    }

    const rawContent = extractTextContent(parsedBody?.choices?.[0]?.message?.content);
    const normalizedContent = normalizeLineEndings(rawContent).trim();

    if (normalizedIntent === 'ask') {
      if (!normalizedContent) {
        return {
          success: false,
          reason: 'AI_INVALID_RESPONSE',
          message: 'Local AI returned an empty answer.'
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
        durationMs: Date.now() - startedAt,
        usage
      });

      return {
        success: true,
        content: normalizedContent,
        intent: normalizedIntent,
        model: selectedModel,
        usage
      };
    }

    const code = extractCodeBlock(rawContent);

    if (!code || !hasVbaProcedure(code)) {
      return {
        success: false,
        reason: 'AI_INVALID_RESPONSE',
        message: 'Local AI did not return valid VBA code.'
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
      durationMs: Date.now() - startedAt,
      usage
    });

    return {
      success: true,
      code,
      intent: normalizedIntent,
      model: selectedModel,
      usage
    };
  } catch (error) {
    const mapped = mapGenerateError(error);
    logger.warn('[AI] generate-vba failed', {
      requestId,
      reason: mapped.reason,
      durationMs: Date.now() - startedAt,
      message: mapped.message
    });
    return {
      success: false,
      reason: mapped.reason,
      message: mapped.message
    };
  }
}

module.exports = {
  generateVba
};
