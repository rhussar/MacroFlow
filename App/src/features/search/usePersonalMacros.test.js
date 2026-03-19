import assert from 'node:assert/strict';
import test from 'node:test';
import {
  normalizePersonalForegroundRefreshPolicy,
  resolvePersonalCacheSignature,
  shouldRefreshPersonalOnForeground,
  shouldDeferPersonalInitialFetch,
  shouldUsePersonalCache
} from './usePersonalMacros.js';

test('shouldUsePersonalCache returns true for matching signature within TTL', () => {
  const now = 20_000;
  const result = shouldUsePersonalCache({
    cachedData: { status: 'ready', macros: [] },
    cachedAt: 10_500,
    cachedSignature: 'a|b|c',
    cachedIncludeShortcutAudit: false,
    nextSignature: 'a|b|c',
    nextIncludeShortcutAudit: false,
    now,
    ttlMs: 10_000
  });

  assert.equal(result, true);
});

test('shouldUsePersonalCache returns false when signature changes', () => {
  const result = shouldUsePersonalCache({
    cachedData: { status: 'ready', macros: [] },
    cachedAt: 1_000,
    cachedSignature: 'a|b',
    cachedIncludeShortcutAudit: false,
    nextSignature: 'a|b|c',
    nextIncludeShortcutAudit: false,
    now: 1_500,
    ttlMs: 10_000
  });

  assert.equal(result, false);
});

test('shouldUsePersonalCache can keep a fresh personal snapshot when signature changes are non-authoritative', () => {
  const result = shouldUsePersonalCache({
    cachedData: { status: 'ready', macros: [] },
    cachedAt: 1_000,
    cachedSignature: 'a|b',
    cachedIncludeShortcutAudit: false,
    nextSignature: 'a|b|c',
    nextIncludeShortcutAudit: false,
    invalidateOnSignatureChange: false,
    now: 1_500,
    ttlMs: 10_000
  });

  assert.equal(result, true);
});

test('shouldUsePersonalCache treats missing next signature as a warm remount and reuses cached signature', () => {
  const result = shouldUsePersonalCache({
    cachedData: { status: 'ready', macros: [] },
    cachedAt: 5_000,
    cachedSignature: 'active|personal',
    cachedIncludeShortcutAudit: false,
    nextSignature: '',
    nextIncludeShortcutAudit: false,
    now: 7_500,
    ttlMs: 10_000
  });

  assert.equal(result, true);
  assert.equal(resolvePersonalCacheSignature('', 'active|personal'), 'active|personal');
});

test('shouldUsePersonalCache returns false when TTL expires', () => {
  const result = shouldUsePersonalCache({
    cachedData: { status: 'ready', macros: [] },
    cachedAt: 1_000,
    cachedSignature: 'a|b',
    cachedIncludeShortcutAudit: false,
    nextSignature: 'a|b',
    nextIncludeShortcutAudit: false,
    now: 12_000,
    ttlMs: 10_000
  });

  assert.equal(result, false);
});

test('shouldUsePersonalCache returns false when shortcut-audit mode changes', () => {
  const result = shouldUsePersonalCache({
    cachedData: { status: 'ready', macros: [] },
    cachedAt: 2_000,
    cachedSignature: 'a|b',
    cachedIncludeShortcutAudit: true,
    nextSignature: 'a|b',
    nextIncludeShortcutAudit: false,
    now: 2_500,
    ttlMs: 10_000
  });

  assert.equal(result, false);
});

test('shouldDeferPersonalInitialFetch defers only on first transition to ready', () => {
  const deferFirstReady = shouldDeferPersonalInitialFetch({
    status: 'ready',
    previousStatus: 'idle',
    hasDeferredInitialFetch: false
  });
  assert.equal(deferFirstReady, true);

  const noDeferAfterAlreadyDeferred = shouldDeferPersonalInitialFetch({
    status: 'ready',
    previousStatus: 'idle',
    hasDeferredInitialFetch: true
  });
  assert.equal(noDeferAfterAlreadyDeferred, false);

  const noDeferWhenAlreadyReady = shouldDeferPersonalInitialFetch({
    status: 'ready',
    previousStatus: 'ready',
    hasDeferredInitialFetch: false
  });
  assert.equal(noDeferWhenAlreadyReady, false);
});

test('normalizePersonalForegroundRefreshPolicy accepts known policies and falls back safely', () => {
  assert.equal(normalizePersonalForegroundRefreshPolicy('always'), 'always');
  assert.equal(normalizePersonalForegroundRefreshPolicy('stale'), 'stale');
  assert.equal(normalizePersonalForegroundRefreshPolicy('never'), 'never');
  assert.equal(normalizePersonalForegroundRefreshPolicy('unexpected', 'always'), 'always');
  assert.equal(normalizePersonalForegroundRefreshPolicy('', 'unexpected'), 'stale');
});

test('shouldRefreshPersonalOnForeground honors always and stale policies', () => {
  const commonArgs = {
    cachedData: { status: 'ready', macros: [] },
    cachedAt: 1_000,
    cachedSignature: 'a|b',
    cachedIncludeShortcutAudit: false,
    nextSignature: 'a|b',
    nextIncludeShortcutAudit: false,
    now: 5_000,
    ttlMs: 10_000
  };

  assert.equal(
    shouldRefreshPersonalOnForeground({
      ...commonArgs,
      policy: 'always'
    }),
    true
  );

  assert.equal(
    shouldRefreshPersonalOnForeground({
      ...commonArgs,
      policy: 'stale'
    }),
    false
  );

  assert.equal(
    shouldRefreshPersonalOnForeground({
      ...commonArgs,
      policy: 'stale',
      now: 20_000
    }),
    true
  );
});

test('shouldRefreshPersonalOnForeground refreshes when shortcut-audit mode changes under stale policy', () => {
  const result = shouldRefreshPersonalOnForeground({
    policy: 'stale',
    cachedData: { status: 'ready', macros: [] },
    cachedAt: 1_000,
    cachedSignature: 'a|b',
    cachedIncludeShortcutAudit: true,
    nextSignature: 'a|b',
    nextIncludeShortcutAudit: false,
    now: 5_000,
    ttlMs: 10_000
  });

  assert.equal(result, true);
});

test('shouldRefreshPersonalOnForeground never refreshes when policy is never', () => {
  const result = shouldRefreshPersonalOnForeground({
    policy: 'never',
    cachedData: null,
    cachedAt: 0,
    cachedSignature: '',
    nextSignature: '',
    now: 10_000
  });

  assert.equal(result, false);
});
