import assert from 'node:assert/strict';
import test from 'node:test';
import {
  resolvePersonalCacheSignature,
  shouldDeferPersonalInitialFetch,
  shouldUsePersonalCache
} from './usePersonalMacros.js';

test('shouldUsePersonalCache returns true for matching signature within TTL', () => {
  const now = 20_000;
  const result = shouldUsePersonalCache({
    cachedData: { status: 'ready', macros: [] },
    cachedAt: 10_500,
    cachedSignature: 'a|b|c',
    nextSignature: 'a|b|c',
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
    nextSignature: 'a|b|c',
    now: 1_500,
    ttlMs: 10_000
  });

  assert.equal(result, false);
});

test('shouldUsePersonalCache treats missing next signature as a warm remount and reuses cached signature', () => {
  const result = shouldUsePersonalCache({
    cachedData: { status: 'ready', macros: [] },
    cachedAt: 5_000,
    cachedSignature: 'active|personal',
    nextSignature: '',
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
    nextSignature: 'a|b',
    now: 12_000,
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
