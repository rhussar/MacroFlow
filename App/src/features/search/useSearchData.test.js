import assert from 'node:assert/strict';
import test from 'node:test';
import {
  isTerminalConnectionStatus,
  inferPauseReasonCodeFromResult,
  getNextPausedReconnectDelayMs,
  shouldSkipForegroundRefresh,
  shouldAttemptPausedReconnect
} from './useSearchData.js';
import {
  SEARCH_FOCUS_REFRESH_COOLDOWN_MS,
  SEARCH_MODE_ENTRY_QUIET_MS,
  SEARCH_PAUSED_RECONNECT_INITIAL_DELAY_MS,
  SEARCH_PAUSED_RECONNECT_MAX_DELAY_MS
} from './search-constants.js';

test('isTerminalConnectionStatus identifies paused terminal states only', () => {
  assert.equal(isTerminalConnectionStatus('no_excel'), true);
  assert.equal(isTerminalConnectionStatus('no_workbook'), true);
  assert.equal(isTerminalConnectionStatus('excel_background'), true);
  assert.equal(isTerminalConnectionStatus('ready'), false);
  assert.equal(isTerminalConnectionStatus('multi_instance'), false);
});

test('inferPauseReasonCodeFromResult prefers reasonCode and parses message fallback', () => {
  assert.equal(
    inferPauseReasonCodeFromResult({ reasonCode: 'NO_VISIBLE_WINDOWS' }, 'NO_EXCEL'),
    'NO_VISIBLE_WINDOWS'
  );
  assert.equal(
    inferPauseReasonCodeFromResult({ message: 'NO_WORKBOOK: workbook not available' }, 'NO_EXCEL'),
    'NO_WORKBOOK'
  );
  assert.equal(
    inferPauseReasonCodeFromResult({ message: 'NO_EXCEL: not running' }, 'NO_VISIBLE_WINDOWS'),
    'NO_EXCEL'
  );
  assert.equal(
    inferPauseReasonCodeFromResult({ message: 'unknown failure' }, 'NO_EXCEL'),
    'NO_EXCEL'
  );
});

test('shouldAttemptPausedReconnect waits until next attempt timestamp', () => {
  const now = 10_000;
  const result = shouldAttemptPausedReconnect({
    isPaused: true,
    now,
    nextAttemptAt: now + 1,
    inFlight: false
  });

  assert.equal(result, false);
});

test('shouldAttemptPausedReconnect allows reconnect when delay has elapsed', () => {
  const now = 10_000;
  const result = shouldAttemptPausedReconnect({
    isPaused: true,
    now,
    nextAttemptAt: now,
    inFlight: false
  });

  assert.equal(result, true);
});

test('shouldAttemptPausedReconnect blocks reconnect storms via inFlight and paused checks', () => {
  const now = 10_000;
  const inFlightBlocked = shouldAttemptPausedReconnect({
    isPaused: true,
    now,
    nextAttemptAt: now - 1,
    inFlight: true
  });
  assert.equal(inFlightBlocked, false);

  const notPaused = shouldAttemptPausedReconnect({
    isPaused: false,
    now,
    nextAttemptAt: now - 1,
    inFlight: false
  });
  assert.equal(notPaused, false);
});

test('getNextPausedReconnectDelayMs applies multiplier and max cap', () => {
  const first = getNextPausedReconnectDelayMs({
    currentDelayMs: SEARCH_PAUSED_RECONNECT_INITIAL_DELAY_MS
  });
  assert.equal(first, 1440);

  const second = getNextPausedReconnectDelayMs({
    currentDelayMs: first
  });
  assert.equal(second, 2592);

  const capped = getNextPausedReconnectDelayMs({
    currentDelayMs: SEARCH_PAUSED_RECONNECT_MAX_DELAY_MS
  });
  assert.equal(capped, SEARCH_PAUSED_RECONNECT_MAX_DELAY_MS);
});

test('getNextPausedReconnectDelayMs uses initial delay fallback for invalid current values', () => {
  const fallback = getNextPausedReconnectDelayMs({
    currentDelayMs: 0
  });
  assert.equal(fallback >= SEARCH_PAUSED_RECONNECT_INITIAL_DELAY_MS, true);
});

test('shouldAttemptPausedReconnect ignores cooldown arg and relies on nextAttemptAt', () => {
  const now = 10_000;
  const allowed = shouldAttemptPausedReconnect({
    isPaused: true,
    now,
    nextAttemptAt: now - 1,
    inFlight: false,
    cooldownMs: SEARCH_FOCUS_REFRESH_COOLDOWN_MS + 1000
  });
  assert.equal(allowed, true);
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
