import assert from 'node:assert/strict';
import test from 'node:test';
import {
  isTerminalConnectionStatus,
  shouldAttemptPausedReconnect
} from './useSearchData.js';
import { SEARCH_FOCUS_REFRESH_COOLDOWN_MS } from './search-constants.js';

test('isTerminalConnectionStatus identifies no_excel and no_workbook only', () => {
  assert.equal(isTerminalConnectionStatus('no_excel'), true);
  assert.equal(isTerminalConnectionStatus('no_workbook'), true);
  assert.equal(isTerminalConnectionStatus('ready'), false);
  assert.equal(isTerminalConnectionStatus('multi_instance'), false);
});

test('shouldAttemptPausedReconnect ignores interval trigger while paused', () => {
  const now = 10_000;
  const result = shouldAttemptPausedReconnect({
    isPaused: true,
    trigger: 'interval',
    now,
    lastResumeAttemptAt: now - SEARCH_FOCUS_REFRESH_COOLDOWN_MS - 1,
    inFlight: false
  });

  assert.equal(result, false);
});

test('shouldAttemptPausedReconnect allows focus-triggered reconnect after cooldown', () => {
  const now = 10_000;
  const result = shouldAttemptPausedReconnect({
    isPaused: true,
    trigger: 'focus',
    now,
    lastResumeAttemptAt: now - SEARCH_FOCUS_REFRESH_COOLDOWN_MS,
    inFlight: false
  });

  assert.equal(result, true);
});

test('shouldAttemptPausedReconnect blocks reconnect storms via inFlight and cooldown checks', () => {
  const now = 10_000;
  const inFlightBlocked = shouldAttemptPausedReconnect({
    isPaused: true,
    trigger: 'visibility',
    now,
    lastResumeAttemptAt: now - SEARCH_FOCUS_REFRESH_COOLDOWN_MS - 1,
    inFlight: true
  });
  assert.equal(inFlightBlocked, false);

  const cooldownBlocked = shouldAttemptPausedReconnect({
    isPaused: true,
    trigger: 'focus',
    now,
    lastResumeAttemptAt: now - SEARCH_FOCUS_REFRESH_COOLDOWN_MS + 1,
    inFlight: false
  });
  assert.equal(cooldownBlocked, false);
});

