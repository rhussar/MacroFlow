import assert from 'node:assert/strict';
import test from 'node:test';
import {
  isTerminalConnectionStatus,
  isSearchDataMode,
  inferPauseReasonCodeFromResult,
  mapPauseReasonCodeToSearchStatus,
  getNextPausedReconnectDelayMs,
  queueSearchLoadRequestState,
  shouldRefreshOnModeEntry,
  shouldSkipForegroundRefresh,
  shouldAttemptPausedReconnect
} from './useSearchData.js';
import {
  SEARCH_FOCUS_REFRESH_COOLDOWN_MS,
  SEARCH_MODE_ENTRY_QUIET_MS,
  SEARCH_PERIODIC_DEEP_REFRESH_STALE_MS,
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

test('isSearchDataMode limits warm-refresh behavior to shortcuts and files tabs', () => {
  assert.equal(isSearchDataMode('search'), true);
  assert.equal(isSearchDataMode('explorer'), true);
  assert.equal(isSearchDataMode('build'), false);
  assert.equal(isSearchDataMode('shortcuts'), false);
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

test('mapPauseReasonCodeToSearchStatus maps pause reason to UI status safely', () => {
  assert.equal(mapPauseReasonCodeToSearchStatus('NO_VISIBLE_WINDOWS', 'error'), 'excel_background');
  assert.equal(mapPauseReasonCodeToSearchStatus('NO_WORKBOOK', 'error'), 'no_workbook');
  assert.equal(mapPauseReasonCodeToSearchStatus('NO_EXCEL', 'error'), 'no_excel');
  assert.equal(mapPauseReasonCodeToSearchStatus('UNKNOWN_REASON', 'error'), 'error');
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

test('queueSearchLoadRequestState queues a first search-load request with its silent mode', () => {
  const queued = queueSearchLoadRequestState({
    hasQueuedLoad: false,
    queuedSilent: true,
    nextSilent: false
  });

  assert.deepEqual(queued, {
    hasQueuedLoad: true,
    queuedSilent: false
  });
});

test('queueSearchLoadRequestState preserves non-silent priority across overlapping requests', () => {
  const queued = queueSearchLoadRequestState({
    hasQueuedLoad: true,
    queuedSilent: false,
    nextSilent: true
  });

  assert.deepEqual(queued, {
    hasQueuedLoad: true,
    queuedSilent: false
  });
});

test('shouldRefreshOnModeEntry refreshes the first active entry and build-to-search transitions', () => {
  const now = 10_000;

  assert.equal(
    shouldRefreshOnModeEntry({
      mode: 'search',
      previousMode: null,
      status: 'idle',
      now
    }),
    true
  );

  assert.equal(
    shouldRefreshOnModeEntry({
      mode: 'search',
      previousMode: 'build',
      status: 'ready',
      lastFullRefreshAt: now,
      now
    }),
    true
  );
});

test('shouldRefreshOnModeEntry skips warm search-files tab switches and same-mode reruns', () => {
  const now = 10_000;

  assert.equal(
    shouldRefreshOnModeEntry({
      mode: 'explorer',
      previousMode: 'search',
      status: 'ready',
      lastFullRefreshAt: now - 1,
      now
    }),
    false
  );

  assert.equal(
    shouldRefreshOnModeEntry({
      mode: 'search',
      previousMode: 'search',
      status: 'ready',
      lastFullRefreshAt: now - 1,
      now
    }),
    false
  );
});

test('shouldRefreshOnModeEntry reloads active-tab transitions when paused, invalid, or stale', () => {
  const now = 10_000;

  assert.equal(
    shouldRefreshOnModeEntry({
      mode: 'search',
      previousMode: 'explorer',
      status: 'ready',
      isPaused: true,
      lastFullRefreshAt: now - 1,
      now
    }),
    true
  );

  assert.equal(
    shouldRefreshOnModeEntry({
      mode: 'search',
      previousMode: 'explorer',
      status: 'error',
      lastFullRefreshAt: now - 1,
      now
    }),
    true
  );

  assert.equal(
    shouldRefreshOnModeEntry({
      mode: 'search',
      previousMode: 'explorer',
      status: 'ready',
      lastFullRefreshAt: now - SEARCH_PERIODIC_DEEP_REFRESH_STALE_MS - 1,
      now
    }),
    true
  );
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
