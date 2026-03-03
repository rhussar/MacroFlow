const https = require('node:https');
const logger = require('./logger');
const {
  OPENAI_MODEL,
  OPENAI_API_BASE,
  SHARED_OPENAI_API_KEY,
  REQUEST_TIMEOUT_MS,
  MAX_CURRENT_CODE_CHARS,
  MAX_PROMPT_CHARS,
  MAX_COMPLETION_TOKENS,
  TEMPERATURE
} = require('./ai-config');

const SYSTEM_PROMPT = [
  'You are an expert VBA engineer for Microsoft Excel.',
  'Return only VBA module code, no markdown, no prose.',
  'Prefer complete module output suitable for replacing the full module.',
  'Include "Option Explicit" unless the user explicitly asks not to.',
  'Preserve intent from the prompt and adapt existing module code when provided.'
].join(' ');

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

  if (statusCode === 401 || statusCode === 403) {
    return {
      reason: 'AI_AUTH_FAILED',
      message: 'OpenAI authentication failed. Please verify API key configuration.'
    };
  }

  if (statusCode === 429) {
    return {
      reason: 'AI_RATE_LIMITED',
      message: 'OpenAI rate limit reached. Please wait and try again.'
    };
  }

  if (messageLower.includes('timeout') || messageLower.includes('timed out') || error?.code === 'ETIMEDOUT') {
    return {
      reason: 'AI_TIMEOUT',
      message: 'OpenAI request timed out. Please try again.'
    };
  }

  if (error?.code === 'ENOTFOUND' || error?.code === 'ECONNRESET' || error?.code === 'ECONNREFUSED') {
    return {
      reason: 'AI_NETWORK_ERROR',
      message: 'Network error while contacting OpenAI. Check connection and retry.'
    };
  }

  if (statusCode >= 400) {
    return {
      reason: 'AI_INVALID_RESPONSE',
      message: rawMessage || `OpenAI request failed (HTTP ${statusCode}).`
    };
  }

  return {
    reason: 'AI_INVALID_RESPONSE',
    message: rawMessage || 'OpenAI returned an invalid response.'
  };
}

function buildUserPrompt({ prompt, workbookName, moduleName, currentCode, includeCurrentCode }) {
  const workbookLabel = toSafeString(workbookName) || 'Unknown Workbook';
  const moduleLabel = toSafeString(moduleName) || 'Unknown Module';
  const trimmedPrompt = truncate(String(prompt || ''), MAX_PROMPT_CHARS);
  const lines = [
    `Workbook: ${workbookLabel}`,
    `Module: ${moduleLabel}`,
    '',
    'Task:',
    trimmedPrompt
  ];

  if (includeCurrentCode) {
    const trimmedCurrentCode = truncate(String(currentCode || ''), MAX_CURRENT_CODE_CHARS);
    const currentCodeBlock = trimmedCurrentCode ? trimmedCurrentCode : "'(no existing code provided)'";
    lines.push('', 'Current module code:', currentCodeBlock);
  }

  return lines.join('\n');
}

function createRequestId() {
  return `ai_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`;
}

function postJsonRequest({ url, headers, body, timeoutMs }) {
  return new Promise((resolve, reject) => {
    const request = https.request(url, { method: 'POST', headers }, (response) => {
      let raw = '';
      response.setEncoding('utf8');
      response.on('data', (chunk) => {
        raw += chunk;
      });
      response.on('end', () => {
        resolve({
          statusCode: Number(response.statusCode) || 0,
          bodyText: raw
        });
      });
    });

    request.setTimeout(timeoutMs, () => {
      request.destroy(new Error('AI_TIMEOUT: request timed out'));
    });

    request.on('error', (error) => {
      reject(error);
    });

    request.write(body);
    request.end();
  });
}

async function generateVba(
  {
    prompt = '',
    workbookName = '',
    moduleName = '',
    currentCode = '',
    includeCurrentCode = false
  } = {},
  dependencies = {}
) {
  const requestId = createRequestId();
  const startedAt = Date.now();
  const promptText = truncate(String(prompt || ''), MAX_PROMPT_CHARS);

  if (!promptText.trim()) {
    return {
      success: false,
      reason: 'AI_INVALID_PROMPT',
      message: 'Prompt is required to generate VBA.'
    };
  }

  const apiKey = toSafeString(SHARED_OPENAI_API_KEY);
  if (!apiKey) {
    return {
      success: false,
      reason: 'AI_AUTH_FAILED',
      message: 'OpenAI API key is not configured.'
    };
  }

  const payload = {
    model: OPENAI_MODEL,
    temperature: TEMPERATURE,
    max_tokens: MAX_COMPLETION_TOKENS,
    messages: [
      {
        role: 'system',
        content: SYSTEM_PROMPT
      },
      {
        role: 'user',
        content: buildUserPrompt({
          prompt: promptText,
          workbookName,
          moduleName,
          currentCode,
          includeCurrentCode: Boolean(includeCurrentCode)
        })
      }
    ]
  };

  logger.debug('[AI] generate-vba request', {
    requestId,
    model: OPENAI_MODEL,
    promptChars: promptText.length,
    includeCurrentCode: Boolean(includeCurrentCode),
    currentCodeChars: Boolean(includeCurrentCode) ? String(currentCode || '').length : 0
  });

  const requestImpl = typeof dependencies.requestImpl === 'function'
    ? dependencies.requestImpl
    : postJsonRequest;

  try {
    const url = `${OPENAI_API_BASE.replace(/\/+$/, '')}/chat/completions`;
    const body = JSON.stringify(payload);

    const response = await requestImpl({
      url,
      timeoutMs: REQUEST_TIMEOUT_MS,
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(body),
        Authorization: `Bearer ${apiKey}`
      },
      body
    });

    const statusCode = Number(response?.statusCode) || 0;
    const bodyText = String(response?.bodyText || '').trim();
    let parsedBody = {};

    try {
      parsedBody = bodyText ? JSON.parse(bodyText) : {};
    } catch {
      throw new Error('AI_INVALID_RESPONSE: OpenAI returned invalid JSON.');
    }

    if (statusCode >= 400) {
      const apiMessage = getErrorMessageFromResponse(parsedBody);
      const mapped = mapGenerateError(new Error(apiMessage || `OpenAI request failed (HTTP ${statusCode}).`), statusCode);
      return {
        success: false,
        reason: mapped.reason,
        message: mapped.message
      };
    }

    const rawContent = extractTextContent(parsedBody?.choices?.[0]?.message?.content);
    const code = extractCodeBlock(rawContent);

    if (!code || !hasVbaProcedure(code)) {
      return {
        success: false,
        reason: 'AI_INVALID_RESPONSE',
        message: 'OpenAI did not return valid VBA code.'
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
      model: OPENAI_MODEL,
      statusCode,
      durationMs: Date.now() - startedAt,
      usage
    });

    return {
      success: true,
      code,
      model: OPENAI_MODEL,
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
