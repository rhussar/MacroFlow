const assert = require('node:assert/strict');
const test = require('node:test');

const { generateVba } = require('./llm-client');
const { resolveLlmPerformanceProfile } = require('./llm-performance');

const TEST_PERFORMANCE_PROFILE = resolveLlmPerformanceProfile({
  totalMemoryBytes: 24 * 1024 * 1024 * 1024,
  cpuCount: 12
});

function readyStatus(overrides = {}) {
  return {
    success: true,
    provider: 'ollama',
    model: 'qwen2.5-coder:3b',
    ready: true,
    needsSetup: false,
    setupInProgress: false,
    runtimeInstalled: true,
    serverReachable: true,
    modelInstalled: true,
    stage: 'ready',
    statusText: 'Local AI is ready.',
    ...overrides
  };
}

function createDeps(overrides = {}) {
  return {
    statusImpl: async () => readyStatus(),
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
        bodyText: JSON.stringify({
          choices: [
            {
              message: {
                content: 'Option Explicit\n\nPublic Sub Hello()\n    MsgBox "hi"\nEnd Sub'
              }
            }
          ],
          usage: {
            prompt_tokens: 123,
            completion_tokens: 44,
            total_tokens: 167
          }
        })
      })
    })
  );

  assert.equal(result.success, true);
  assert.match(result.code, /Sub Hello/i);
  assert.equal(result.model, 'qwen2.5-coder:3b');
  assert.equal(result.usage?.totalTokens, 167);
  assert.equal(result.diagnostics?.profile, TEST_PERFORMANCE_PROFILE.name);
});

test('generateVba extracts fenced VBA code blocks', async () => {
  const result = await generateVba(
    { prompt: 'make a macro', moduleName: 'Module1' },
    createDeps({
      requestImpl: async () => ({
        statusCode: 200,
        bodyText: JSON.stringify({
          choices: [
            {
              message: {
                content: '```vba\nOption Explicit\nSub RunA()\nEnd Sub\n```'
              }
            }
          ]
        })
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

test('generateVba rejects when local AI runtime or model is missing', async () => {
  const runtimeMissing = await generateVba(
    { prompt: 'x' },
    createDeps({
      statusImpl: async () => readyStatus({ runtimeInstalled: false, ready: false, needsSetup: true })
    })
  );
  assert.equal(runtimeMissing.success, false);
  assert.equal(runtimeMissing.reason, 'AI_RUNTIME_MISSING');

  const modelMissing = await generateVba(
    { prompt: 'x' },
    createDeps({
      statusImpl: async () => readyStatus({ ready: false, modelInstalled: false, needsSetup: true })
    })
  );
  assert.equal(modelMissing.success, false);
  assert.equal(modelMissing.reason, 'AI_MODEL_MISSING');
});

test('generateVba maps malformed content to AI_INVALID_RESPONSE', async () => {
  const result = await generateVba(
    { prompt: 'do stuff' },
    createDeps({
      requestImpl: async () => ({
        statusCode: 200,
        bodyText: JSON.stringify({
          choices: [{ message: { content: 'No procedures here' } }]
        })
      })
    })
  );

  assert.equal(result.success, false);
  assert.equal(result.reason, 'AI_INVALID_RESPONSE');
});

test('generateVba maps model-missing, timeout, and connection failures', async () => {
  const missing = await generateVba(
    { prompt: 'x' },
    createDeps({
      requestImpl: async () => ({
        statusCode: 404,
        bodyText: JSON.stringify({ error: { message: 'model not found' } })
      })
    })
  );
  assert.equal(missing.success, false);
  assert.equal(missing.reason, 'AI_MODEL_MISSING');

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
        const error = new Error('connect ECONNREFUSED 127.0.0.1:11434');
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
        userContent = String(payload?.messages?.[1]?.content || '');
        const taskMatch = userContent.match(/Task:\n([\s\S]*)$/);
        taskText = taskMatch ? taskMatch[1] : '';

        return {
          statusCode: 200,
          bodyText: JSON.stringify({
            choices: [{ message: { content: 'Sub RunA()\nEnd Sub' } }]
          })
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
        const userContent = String(payload?.messages?.[1]?.content || '');
        const codeMatch = userContent.match(/Current module code:\n([\s\S]*)$/);
        codeText = codeMatch ? codeMatch[1] : '';

        return {
          statusCode: 200,
          bodyText: JSON.stringify({
            choices: [{ message: { content: 'Sub RunA()\nEnd Sub' } }]
          })
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
          bodyText: JSON.stringify({
            choices: [{ message: { content: 'Sub RunA()\nEnd Sub' } }]
          })
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
        userContent = String(payload?.messages?.[1]?.content || '');
        return {
          statusCode: 200,
          bodyText: JSON.stringify({
            choices: [{ message: { content: 'Option Explicit\nSub RunA()\nEnd Sub' } }]
          })
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
        bodyText: JSON.stringify({
          choices: [{ message: { content: 'It forces variable declarations before use.' } }]
        })
      })
    })
  );

  assert.equal(result.success, true);
  assert.equal(result.intent, 'ask');
  assert.equal(result.content, 'It forces variable declarations before use.');
});
