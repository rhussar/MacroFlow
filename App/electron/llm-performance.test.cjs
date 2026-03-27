const assert = require('node:assert/strict');
const test = require('node:test');

const { detectHardwareTier, resolveLlmPerformanceProfile } = require('./llm-performance');

test('detectHardwareTier classifies weaker laptops as low_resource', () => {
  const result = detectHardwareTier({
    totalMemoryBytes: 8 * 1024 * 1024 * 1024,
    cpuCount: 4
  });

  assert.equal(result.tier, 'low_resource');
});

test('detectHardwareTier falls back to low_resource when free memory is tight', () => {
  const result = detectHardwareTier({
    totalMemoryBytes: 16 * 1024 * 1024 * 1024,
    freeMemoryBytes: 4 * 1024 * 1024 * 1024,
    cpuCount: 8
  });

  assert.equal(result.tier, 'low_resource');
});

test('detectHardwareTier classifies mid-range machines as balanced', () => {
  const result = detectHardwareTier({
    totalMemoryBytes: 16 * 1024 * 1024 * 1024,
    freeMemoryBytes: 9 * 1024 * 1024 * 1024,
    cpuCount: 8
  });

  assert.equal(result.tier, 'balanced');
});

test('detectHardwareTier classifies stronger machines as standard', () => {
  const result = detectHardwareTier({
    totalMemoryBytes: 32 * 1024 * 1024 * 1024,
    freeMemoryBytes: 20 * 1024 * 1024 * 1024,
    cpuCount: 12
  });

  assert.equal(result.tier, 'standard');
});

test('resolveLlmPerformanceProfile applies profile-specific budgets', () => {
  const lowResourceProfile = resolveLlmPerformanceProfile({
    totalMemoryBytes: 8 * 1024 * 1024 * 1024,
    cpuCount: 4
  });
  const standardProfile = resolveLlmPerformanceProfile({
    totalMemoryBytes: 32 * 1024 * 1024 * 1024,
    freeMemoryBytes: 20 * 1024 * 1024 * 1024,
    cpuCount: 12
  });

  assert.equal(lowResourceProfile.name, 'low_resource');
  assert.equal(standardProfile.name, 'standard');
  assert.ok(lowResourceProfile.contextLength < standardProfile.contextLength);
  assert.ok(lowResourceProfile.maxCompletionTokens < standardProfile.maxCompletionTokens);
  assert.ok(lowResourceProfile.maxWorkbookContextChars < standardProfile.maxWorkbookContextChars);
  assert.equal(lowResourceProfile.ollamaMaxLoadedModels, 1);
  assert.equal(lowResourceProfile.ollamaNumParallel, 1);
});
