const assert = require('node:assert/strict');
const test = require('node:test');

const { generateVba } = require('./llm-client');
const { resolveLlmPerformanceProfile } = require('./llm-performance');

const TEST_PERFORMANCE_PROFILE = resolveLlmPerformanceProfile({
  totalMemoryBytes: 24 * 1024 * 1024 * 1024,
  cpuCount: 12
});

function anthropicBody(text, usage) {
  return JSON.stringify({
    content: [{ type: 'text', text }],
    usage: usage || {
      input_tokens: 123,
      output_tokens: 44
    }
  });
}

function createDeps(overrides = {}) {
  return {
    contextImpl: async () => '',
    performanceProfile: TEST_PERFORMANCE_PROFILE,
    ...overrides
  };
}

test('generateVba returns success for valid plain VBA output', async () => {
  const result = await generateVba(
    {
      prompt: 'Create a hello world macro',
      workbookName: 'Book1.xlsm',
      moduleName: 'Module1',
      currentCode: 'Option Explicit'
    },
    createDeps({
      requestImpl: async () => ({
        statusCode: 200,
        bodyText: anthropicBody('Option Explicit\n\nPublic Sub Hello()\n    MsgBox "hi"\nEnd Sub')
      })
    })
  );

  assert.equal(result.success, true);
  assert.match(result.code, /Sub Hello/i);
  assert.equal(result.model, 'claude-sonnet-5');
  assert.equal(result.usage?.totalTokens, 167);
  assert.equal(result.diagnostics?.profile, TEST_PERFORMANCE_PROFILE.name);
});

test('generateVba extracts fenced VBA code blocks', async () => {
  const result = await generateVba(
    { prompt: 'make a macro', moduleName: 'Module1' },
    createDeps({
      requestImpl: async () => ({
        statusCode: 200,
        bodyText: anthropicBody('```vba\nOption Explicit\nSub RunA()\nEnd Sub\n```')
      })
    })
  );

  assert.equal(result.success, true);
  assert.equal(result.code, 'Option Explicit\nSub RunA()\nEnd Sub');
});

test('generateVba rejects empty prompt before request', async () => {
  let requestCalls = 0;
  const result = await generateVba(
    { prompt: '   ' },
    createDeps({
      requestImpl: async () => {
        requestCalls += 1;
        return { statusCode: 200, bodyText: '{}' };
      }
    })
  );

  assert.equal(result.success, false);
  assert.equal(result.reason, 'AI_INVALID_PROMPT');
  assert.equal(requestCalls, 0);
});

test('generateVba maps malformed content to AI_INVALID_RESPONSE', async () => {
  const result = await generateVba(
    { prompt: 'do stuff' },
    createDeps({
      requestImpl: async () => ({
        statusCode: 200,
        bodyText: anthropicBody('No procedures here')
      })
    })
  );

  assert.equal(result.success, false);
  assert.equal(result.reason, 'AI_INVALID_RESPONSE');
});

test('generateVba maps rate-limit, timeout, and connection failures', async () => {
  const limited = await generateVba(
    { prompt: 'x' },
    createDeps({
      requestImpl: async () => ({
        statusCode: 429,
        bodyText: JSON.stringify({ error: { message: 'Rate limit exceeded. Try again later.' } })
      })
    })
  );
  assert.equal(limited.success, false);
  assert.equal(limited.reason, 'AI_RATE_LIMITED');

  const timeout = await generateVba(
    { prompt: 'x' },
    createDeps({
      requestImpl: async () => {
        const error = new Error('socket timed out');
        error.code = 'ETIMEDOUT';
        throw error;
      }
    })
  );
  assert.equal(timeout.success, false);
  assert.equal(timeout.reason, 'AI_TIMEOUT');

  const notReady = await generateVba(
    { prompt: 'x' },
    createDeps({
      requestImpl: async () => {
        const error = new Error('connect ECONNREFUSED 127.0.0.1:8787');
        error.code = 'ECONNREFUSED';
        throw error;
      }
    })
  );
  assert.equal(notReady.success, false);
  assert.equal(notReady.reason, 'AI_NOT_READY');
});

test('generateVba truncates prompt and omits current code by default', async () => {
  const longPrompt = 'P'.repeat(TEST_PERFORMANCE_PROFILE.maxPromptChars + 150);
  const longCurrentCode = 'C'.repeat(TEST_PERFORMANCE_PROFILE.maxCurrentCodeChars + 500);
  let taskText = '';
  let userContent = '';

  const result = await generateVba(
    {
      prompt: longPrompt,
      workbookName: 'Book1.xlsm',
      moduleName: 'Module1',
      currentCode: longCurrentCode
    },
    createDeps({
      requestImpl: async ({ body }) => {
        const payload = JSON.parse(String(body || '{}'));
        userContent = String(payload?.messages?.[0]?.content || '');
        const taskMatch = userContent.match(/Task:\n([\s\S]*)$/);
        taskText = taskMatch ? taskMatch[1] : '';

        return {
          statusCode: 200,
          bodyText: anthropicBody('Sub RunA()\nEnd Sub')
        };
      }
    })
  );

  assert.equal(result.success, true);
  assert.equal(taskText.length, TEST_PERFORMANCE_PROFILE.maxPromptChars);
  assert.equal(userContent.includes('Current module code:'), false);
});

test('generateVba includes and truncates current code when opted in', async () => {
  const longCurrentCode = 'C'.repeat(TEST_PERFORMANCE_PROFILE.maxCurrentCodeChars + 500);
  let codeText = '';

  const result = await generateVba(
    {
      prompt: 'Generate a macro',
      workbookName: 'Book1.xlsm',
      moduleName: 'Module1',
      currentCode: longCurrentCode,
      includeCurrentCode: true
    },
    createDeps({
      requestImpl: async ({ body }) => {
        const payload = JSON.parse(String(body || '{}'));
        const userContent = String(payload?.messages?.[0]?.content || '');
        const codeMatch = userContent.match(/Current module code:\n([\s\S]*)$/);
        codeText = codeMatch ? codeMatch[1] : '';

        return {
          statusCode: 200,
          bodyText: anthropicBody('Sub RunA()\nEnd Sub')
        };
      }
    })
  );

  assert.equal(result.success, true);
  assert.equal(codeText.length, TEST_PERFORMANCE_PROFILE.maxCurrentCodeChars);
});

test('generateVba uses performance profile max token budget', async () => {
  let observedMaxTokens = 0;

  const result = await generateVba(
    {
      prompt: 'Generate a macro',
      workbookName: 'Book1.xlsm',
      moduleName: 'Module1'
    },
    createDeps({
      requestImpl: async ({ body }) => {
        const payload = JSON.parse(String(body || '{}'));
        observedMaxTokens = Number(payload?.max_tokens) || 0;
        return {
          statusCode: 200,
          bodyText: anthropicBody('Sub RunA()\nEnd Sub')
        };
      }
    })
  );

  assert.equal(result.success, true);
  assert.equal(observedMaxTokens, TEST_PERFORMANCE_PROFILE.maxCompletionTokens);
});

test('generateVba includes workbook context when available', async () => {
  let userContent = '';

  const result = await generateVba(
    {
      prompt: 'Create a subtotal macro',
      intent: 'create',
      workbookName: 'Book1.xlsm',
      workbookPath: 'C:\\Temp\\Book1.xlsm',
      moduleName: 'Module1'
    },
    createDeps({
      contextImpl: async () => 'Active sheet: Sales\nHeaders: A1=Region, B1=Amount',
      requestImpl: async ({ body }) => {
        const payload = JSON.parse(String(body || '{}'));
        userContent = String(payload?.messages?.[0]?.content || '');
        return {
          statusCode: 200,
          bodyText: anthropicBody('Option Explicit\nSub RunA()\nEnd Sub')
        };
      }
    })
  );

  assert.equal(result.success, true);
  assert.match(userContent, /Workbook context:/);
  assert.match(userContent, /Active sheet: Sales/);
});

test('generateVba returns prose content for ask intent', async () => {
  const result = await generateVba(
    {
      prompt: 'What does Option Explicit do?',
      intent: 'ask',
      workbookName: 'Book1.xlsm',
      moduleName: 'Module1'
    },
    createDeps({
      requestImpl: async () => ({
        statusCode: 200,
        bodyText: anthropicBody('It forces variable declarations before use.')
      })
    })
  );

  assert.equal(result.success, true);
  assert.equal(result.intent, 'ask');
  assert.equal(result.content, 'It forces variable declarations before use.');
});
