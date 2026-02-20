import assert from 'node:assert/strict';
import test from 'node:test';
import {
  isTerminalConnectionStatus,
  shouldSkipForegroundRefresh,
  shouldAttemptPausedReconnect
} from './useSearchData.js';
import {
  SEARCH_FOCUS_REFRESH_COOLDOWN_MS,
  SEARCH_MODE_ENTRY_QUIET_MS
} from './search-constants.js';

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

test('shouldSkipForegroundRefresh blocks focus/visibility/helper during mode-entry quiet window', () => {
  const now = 10_000;
  const modeEntryAt = now - SEARCH_MODE_ENTRY_QUIET_MS + 1;

  assert.equal(
    shouldSkipForegroundRefresh({
      trigger: 'focus',
      now,
      modeEntryAt,
      lastForegroundRefreshAt: 0
    }),
    true
  );
  assert.equal(
    shouldSkipForegroundRefresh({
      trigger: 'visibility',
      now,
      modeEntryAt,
      lastForegroundRefreshAt: 0
    }),
    true
  );
  assert.equal(
    shouldSkipForegroundRefresh({
      trigger: 'helper',
      now,
      modeEntryAt,
      lastForegroundRefreshAt: 0
    }),
    true
  );
});

test('shouldSkipForegroundRefresh allows foreground refresh after quiet window and cooldown', () => {
  const now = 10_000;
  const modeEntryAt = now - SEARCH_MODE_ENTRY_QUIET_MS - 1;
  const lastForegroundRefreshAt = now - SEARCH_FOCUS_REFRESH_COOLDOWN_MS - 1;

  const result = shouldSkipForegroundRefresh({
    trigger: 'focus',
    now,
    modeEntryAt,
    lastForegroundRefreshAt
  });

  assert.equal(result, false);
});

test('shouldSkipForegroundRefresh respects shared min-gap guard and ignores interval', () => {
  const now = 10_000;
  const modeEntryAt = now - SEARCH_MODE_ENTRY_QUIET_MS - 1;
  const tooRecent = shouldSkipForegroundRefresh({
    trigger: 'visibility',
    now,
    modeEntryAt,
    lastForegroundRefreshAt: now - SEARCH_FOCUS_REFRESH_COOLDOWN_MS + 1
  });
  assert.equal(tooRecent, true);

  const intervalResult = shouldSkipForegroundRefresh({
    trigger: 'interval',
    now,
    modeEntryAt,
    lastForegroundRefreshAt: now
  });
  assert.equal(intervalResult, false);
});
