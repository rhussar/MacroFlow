const {
  MAX_CURRENT_CODE_CHARS,
  MAX_PROMPT_CHARS,
  MAX_WORKBOOK_CONTEXT_CHARS
} = require('./llm-config');

const AI_INTENTS = new Set(['ask', 'create', 'edit']);

const SYSTEM_PROMPTS = {
  ask: [
    'You are a VBA assistant. STRICT RULES:',
    '1. Maximum 3 sentences. No exceptions.',
    '2. Never repeat or explain code that the user can already see.',
    '3. No code blocks unless the user explicitly asks for code.',
    '4. No numbered lists, no step-by-step breakdowns, no headings.',
    '5. Answer like a knowledgeable coworker in a quick chat — brief and direct.'
  ].join(' '),
  create: [
    'You are an expert VBA engineer for Microsoft Excel.',
    'Return only VBA module code, no markdown, no prose.',
    'Generate a complete module suitable for insertion into the target workbook.',
    'Include "Option Explicit" unless the user explicitly asks not to.',
    'Use workbook and worksheet context when provided.',
    'Avoid placeholders and produce runnable VBA.'
  ].join(' '),
  edit: [
    'You are an expert VBA engineer for Microsoft Excel.',
    'Return only VBA module code, no markdown, no prose.',
    'You are editing an existing module.',
    'Preserve the current module intent unless the user asks to replace it.',
    'Return a complete replacement module.',
    'Include "Option Explicit" unless the user explicitly asks not to.',
    'Use workbook and worksheet context when provided.'
  ].join(' ')
};

function toSafeString(value) {
  return String(value || '').trim();
}

function truncate(value, maxChars) {
  const normalized = String(value || '');
  if (normalized.length <= maxChars) {
    return normalized;
  }
  return normalized.slice(0, maxChars);
}

function normalizeAiIntent(value, options = {}) {
  const normalized = toSafeString(value).toLowerCase();
  if (AI_INTENTS.has(normalized)) {
    return normalized;
  }

  return options.includeCurrentCode ? 'edit' : 'create';
}

function buildUserPrompt({
  intent = 'create',
  prompt = '',
  workbookName = '',
  moduleName = '',
  currentCode = '',
  includeCurrentCode = false,
  workbookContext = '',
  limits = {}
}) {
  const normalizedIntent = normalizeAiIntent(intent, { includeCurrentCode });
  const workbookLabel = toSafeString(workbookName) || 'Unknown Workbook';
  const moduleLabel = toSafeString(moduleName) || 'Unknown Module';
  const maxPromptChars = Number(limits.maxPromptChars) || MAX_PROMPT_CHARS;
  const maxCurrentCodeChars = Number(limits.maxCurrentCodeChars) || MAX_CURRENT_CODE_CHARS;
  const maxWorkbookContextChars = Number(limits.maxWorkbookContextChars) || MAX_WORKBOOK_CONTEXT_CHARS;
  const trimmedPrompt = truncate(prompt, maxPromptChars);
  const lines = [
    `Intent: ${normalizedIntent}`,
    `Workbook: ${workbookLabel}`,
    `Module: ${moduleLabel}`
  ];

  if (toSafeString(workbookContext)) {
    lines.push('', 'Workbook context:', truncate(workbookContext, maxWorkbookContextChars));
  }

  lines.push('', normalizedIntent === 'ask' ? 'Question:' : 'Task:', trimmedPrompt);

  if ((normalizedIntent === 'edit' || normalizedIntent === 'ask') && includeCurrentCode) {
    const trimmedCurrentCode = truncate(currentCode, maxCurrentCodeChars);
    const currentCodeBlock = trimmedCurrentCode ? trimmedCurrentCode : "'(no existing code provided)'";
    lines.push('', 'Current module code:', currentCodeBlock);
  }

  return lines.join('\n');
}

function getSystemPrompt(intent, options = {}) {
  return SYSTEM_PROMPTS[normalizeAiIntent(intent, options)] || SYSTEM_PROMPTS.create;
}

module.exports = {
  normalizeAiIntent,
  buildUserPrompt,
  getSystemPrompt
};
