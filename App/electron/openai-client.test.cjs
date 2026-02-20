const assert = require('node:assert/strict');
const test = require('node:test');

const { generateVba } = require('./openai-client');
const { MAX_CURRENT_CODE_CHARS, MAX_PROMPT_CHARS } = require('./ai-config');

test('generateVba returns success for valid plain VBA output', async () => {
  const result = await generateVba(
    {
      prompt: 'Create a hello world macro',
      workbookName: 'Book1.xlsm',
      moduleName: 'Module1',
      currentCode: 'Option Explicit'
    },
    {
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
    }
  );

  assert.equal(result.success, true);
  assert.match(result.code, /Sub Hello/i);
  assert.equal(result.model, 'gpt-4.1-mini');
  assert.equal(result.usage?.totalTokens, 167);
});

test('generateVba extracts fenced VBA code blocks', async () => {
  const result = await generateVba(
    { prompt: 'make a macro', moduleName: 'Module1' },
    {
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
    }
  );

  assert.equal(result.success, true);
  assert.equal(result.code, 'Option Explicit\nSub RunA()\nEnd Sub');
});

test('generateVba rejects empty prompt before request', async () => {
  let requestCalls = 0;
  const result = await generateVba(
    { prompt: '   ' },
    {
      requestImpl: async () => {
        requestCalls += 1;
        return { statusCode: 200, bodyText: '{}' };
      }
    }
  );

  assert.equal(result.success, false);
  assert.equal(result.reason, 'AI_INVALID_PROMPT');
  assert.equal(requestCalls, 0);
});

test('generateVba maps malformed content to AI_INVALID_RESPONSE', async () => {
  const result = await generateVba(
    { prompt: 'do stuff' },
    {
      requestImpl: async () => ({
        statusCode: 200,
        bodyText: JSON.stringify({
          choices: [{ message: { content: 'No procedures here' } }]
        })
      })
    }
  );

  assert.equal(result.success, false);
  assert.equal(result.reason, 'AI_INVALID_RESPONSE');
});

test('generateVba maps HTTP auth and rate-limit errors', async () => {
  const auth = await generateVba(
    { prompt: 'x' },
    {
      requestImpl: async () => ({
        statusCode: 401,
        bodyText: JSON.stringify({ error: { message: 'bad key' } })
      })
    }
  );
  assert.equal(auth.success, false);
  assert.equal(auth.reason, 'AI_AUTH_FAILED');

  const limited = await generateVba(
    { prompt: 'x' },
    {
      requestImpl: async () => ({
        statusCode: 429,
        bodyText: JSON.stringify({ error: { message: 'too many requests' } })
      })
    }
  );
  assert.equal(limited.success, false);
  assert.equal(limited.reason, 'AI_RATE_LIMITED');
});

test('generateVba maps timeout and network failures', async () => {
  const timeout = await generateVba(
    { prompt: 'x' },
    {
      requestImpl: async () => {
        const error = new Error('socket timed out');
        error.code = 'ETIMEDOUT';
        throw error;
      }
    }
  );
  assert.equal(timeout.success, false);
  assert.equal(timeout.reason, 'AI_TIMEOUT');

  const network = await generateVba(
    { prompt: 'x' },
    {
      requestImpl: async () => {
        const error = new Error('getaddrinfo ENOTFOUND api.openai.com');
        error.code = 'ENOTFOUND';
        throw error;
      }
    }
  );
  assert.equal(network.success, false);
  assert.equal(network.reason, 'AI_NETWORK_ERROR');
});

test('generateVba truncates prompt and current code before request', async () => {
  const longPrompt = 'P'.repeat(MAX_PROMPT_CHARS + 150);
  const longCurrentCode = 'C'.repeat(MAX_CURRENT_CODE_CHARS + 500);
  let taskText = '';
  let codeText = '';

  const result = await generateVba(
    {
      prompt: longPrompt,
      workbookName: 'Book1.xlsm',
      moduleName: 'Module1',
      currentCode: longCurrentCode
    },
    {
      requestImpl: async ({ body }) => {
        const payload = JSON.parse(String(body || '{}'));
        const userContent = String(payload?.messages?.[1]?.content || '');
        const taskMatch = userContent.match(/Task:\n([\s\S]*?)\n\nCurrent module code:/);
        const codeMatch = userContent.match(/Current module code:\n([\s\S]*)$/);
        taskText = taskMatch ? taskMatch[1] : '';
        codeText = codeMatch ? codeMatch[1] : '';

        return {
          statusCode: 200,
          bodyText: JSON.stringify({
            choices: [{ message: { content: 'Sub RunA()\nEnd Sub' } }]
          })
        };
      }
    }
  );

  assert.equal(result.success, true);
  assert.equal(taskText.length, MAX_PROMPT_CHARS);
  assert.equal(codeText.length, MAX_CURRENT_CODE_CHARS);
});

