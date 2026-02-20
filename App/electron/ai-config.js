/**
 * OpenAI configuration for MVP VBA generation.
 *
 * NOTE:
 * - This shared key is intentionally bundled for MVP distribution.
 * - Do not log the key.
 */

const OPENAI_MODEL = 'gpt-4.1-mini';
const OPENAI_API_BASE = 'https://api.openai.com/v1';
const SHARED_OPENAI_API_KEY =
  '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';
const REQUEST_TIMEOUT_MS = 30000;
const MAX_CURRENT_CODE_CHARS = 12000;
const MAX_PROMPT_CHARS = 4000;
const MAX_COMPLETION_TOKENS = 1400;
const TEMPERATURE = 0.2;

module.exports = {
  OPENAI_MODEL,
  OPENAI_API_BASE,
  SHARED_OPENAI_API_KEY,
  REQUEST_TIMEOUT_MS,
  MAX_CURRENT_CODE_CHARS,
  MAX_PROMPT_CHARS,
  MAX_COMPLETION_TOKENS,
  TEMPERATURE
};
