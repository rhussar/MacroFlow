import { useState, useRef, useCallback, useEffect } from 'react';
import {
  mapSearchError,
  normalizeMacros,
  normalizeModules,
  normalizeWorkbook
} from '../../lib/search-data.js';
import {
  INITIAL_SEARCH_DATA,
  SEARCH_FOCUS_REFRESH_COOLDOWN_MS,
  SEARCH_MODE_ENTRY_QUIET_MS,
  SEARCH_HELPER_EVENT_COOLDOWN_MS,
  SEARCH_HELPER_EVENT_DEBOUNCE_MS,
  SEARCH_FULL_REFRESH_STALE_MS,
  SEARCH_PERIODIC_DEEP_REFRESH_STALE_MS,
  SEARCH_PERIODIC_REFRESH_MS,
  SEARCH_PAUSED_RECONNECT_TICK_MS,
  SEARCH_PAUSED_RECONNECT_INITIAL_DELAY_MS,
  SEARCH_PAUSED_RECONNECT_MULTIPLIER,
  SEARCH_PAUSED_RECONNECT_MAX_DELAY_MS
} from './search-constants.js';

export function isTerminalConnectionStatus(status) {
  return status === 'no_excel' || status === 'no_workbook' || status === 'excel_background';
}

export function inferPauseReasonCodeFromResult(result, fallback = 'NO_EXCEL') {
  const normalizedFallback = String(fallback || '').trim().toUpperCase() || 'NO_EXCEL';
  const reasonCandidate = String(result?.reasonCode || result?.code || '').trim().toUpperCase();
  if (reasonCandidate === 'NO_EXCEL' || reasonCandidate === 'NO_WORKBOOK' || reasonCandidate === 'NO_VISIBLE_WINDOWS') {
    return reasonCandidate;
  }

  const message = String(result?.message || result?.error || '').trim().toUpperCase();
  if (message.includes('NO_VISIBLE_WINDOWS')) {
    return 'NO_VISIBLE_WINDOWS';
  }
  if (message.includes('NO_WORKBOOK')) {
    return 'NO_WORKBOOK';
  }
  if (message.includes('NO_EXCEL')) {
    return 'NO_EXCEL';
  }
  return normalizedFallback;
}

export function mapPauseReasonCodeToSearchStatus(reasonCode, fallback = 'no_excel') {
  const normalizedReason = String(reasonCode || '').trim().toUpperCase();
  if (normalizedReason === 'NO_VISIBLE_WINDOWS') {
    return 'excel_background';
  }
  if (normalizedReason === 'NO_WORKBOOK') {
    return 'no_workbook';
  }
  if (normalizedReason === 'NO_EXCEL') {
    return 'no_excel';
  }
  return String(fallback || 'no_excel');
}

export function getNextPausedReconnectDelayMs({
  currentDelayMs,
  initialDelayMs = SEARCH_PAUSED_RECONNECT_INITIAL_DELAY_MS,
  multiplier = SEARCH_PAUSED_RECONNECT_MULTIPLIER,
  maxDelayMs = SEARCH_PAUSED_RECONNECT_MAX_DELAY_MS
}) {
  const normalizedInitial = Number(initialDelayMs);
  const normalizedMax = Number(maxDelayMs);
  const normalizedMultiplier = Number(multiplier);
  const normalizedCurrent = Number(currentDelayMs);

  const baseDelay = Number.isFinite(normalizedCurrent) && normalizedCurrent > 0
    ? normalizedCurrent
    : normalizedInitial;
  const nextDelay = Math.max(
    normalizedInitial,
    Math.round(baseDelay * normalizedMultiplier)
  );
  return Math.min(nextDelay, normalizedMax);
}

export function shouldAttemptPausedReconnect({
  isPaused,
  now,
  nextAttemptAt = 0,
  inFlight
}) {
  if (!isPaused || inFlight) {
    return false;
  }

  return now >= Number(nextAttemptAt || 0);
}

export function shouldSkipForegroundRefresh({
  trigger,
  now,
  modeEntryAt,
  lastForegroundRefreshAt,
  quietWindowMs = SEARCH_MODE_ENTRY_QUIET_MS,
  minGapMs = SEARCH_FOCUS_REFRESH_COOLDOWN_MS
}) {
  const isForegroundTrigger = trigger === 'focus' || trigger === 'visibility' || trigger === 'helper';
  if (!isForegroundTrigger) {
    return false;
  }

  if (now - modeEntryAt < quietWindowMs) {
    return true;
  }

  if (now - lastForegroundRefreshAt < minGapMs) {
    return true;
  }

  return false;
}

export function isSearchDataMode(mode) {
  return mode === 'search' || mode === 'explorer';
}

export function shouldRefreshOnModeEntry({
  mode,
  previousMode,
  status,
  isPaused = false,
  lastFullRefreshAt = 0,
  now = Date.now(),
  staleThresholdMs = SEARCH_PERIODIC_DEEP_REFRESH_STALE_MS
}) {
  if (!isSearchDataMode(mode)) {
    return false;
  }

  if (previousMode === mode) {
    return false;
  }

  if (!isSearchDataMode(previousMode)) {
    return true;
  }

  if (isPaused || status !== 'ready') {
    return true;
  }

  return now - Number(lastFullRefreshAt || 0) > staleThresholdMs;
}

export function useSearchData({ mode, runState, macroRunInFlightRef, shortcutSaveInFlightRef }) {
  const [searchData, setSearchData] = useState(INITIAL_SEARCH_DATA);
  const searchRequestSequence = useRef(0);
  const searchLoadInFlight = useRef(false);
  const workbookPingInFlight = useRef(false);
  const pollingPausedRef = useRef(false);
  const pollingPausedReasonRef = useRef('');
  const resumeAttemptInFlightRef = useRef(false);
  const lastResumeAttemptAtRef = useRef(0);
  const pausedReconnectDelayMsRef = useRef(SEARCH_PAUSED_RECONNECT_INITIAL_DELAY_MS);
  const pausedReconnectNextAttemptAtRef = useRef(0);
  const modeEntryAtRef = useRef(0);
  const lastForegroundRefreshAtRef = useRef(0);
  const helperRefreshDebounceTimerRef = useRef(null);
  const helperRefreshInFlightRef = useRef(false);
  const lastHelperRefreshAtRef = useRef(0);
  const lastFullSearchRefreshAt = useRef(0);
  const lastWorkbookSignature = useRef('');
  const resolveInstanceAttempted = useRef(false);
  const previousModeRef = useRef(null);
  const searchStatusRef = useRef(INITIAL_SEARCH_DATA.status);
  const refreshSearchOnForegroundRef = useRef(null);

  searchStatusRef.current = searchData.status;

  const resetPausedReconnectBackoff = () => {
    pausedReconnectDelayMsRef.current = SEARCH_PAUSED_RECONNECT_INITIAL_DELAY_MS;
    pausedReconnectNextAttemptAtRef.current = 0;
  };

  const schedulePausedReconnectBackoff = (now = Date.now()) => {
    const nextDelayMs = getNextPausedReconnectDelayMs({
      currentDelayMs: pausedReconnectDelayMsRef.current
    });
    pausedReconnectDelayMsRef.current = nextDelayMs;
    pausedReconnectNextAttemptAtRef.current = now + nextDelayMs;
  };

  const setPausedState = (reasonCode = 'NO_EXCEL') => {
    const normalizedReason = String(reasonCode || '').trim().toUpperCase() || 'NO_EXCEL';
    const wasPaused = pollingPausedRef.current;
    const previousReason = String(pollingPausedReasonRef.current || '').trim().toUpperCase();

    pollingPausedRef.current = true;
    pollingPausedReasonRef.current = normalizedReason;
    if (!wasPaused || previousReason !== normalizedReason) {
      pausedReconnectDelayMsRef.current = SEARCH_PAUSED_RECONNECT_INITIAL_DELAY_MS;
      pausedReconnectNextAttemptAtRef.current = Date.now() + SEARCH_PAUSED_RECONNECT_INITIAL_DELAY_MS;
    }
  };

  const clearPausedState = () => {
    pollingPausedRef.current = false;
    pollingPausedReasonRef.current = '';
    resetPausedReconnectBackoff();
  };

  const applyPausedReconnectFailure = (resultLike) => {
    const reasonCode = inferPauseReasonCodeFromResult(resultLike, 'NO_EXCEL');
    setPausedState(reasonCode);
    schedulePausedReconnectBackoff(Date.now());

    const fallbackMessage = `${reasonCode}: Unable to reconnect to Excel.`;
    const rawMessage = String(resultLike?.message || resultLike?.error || fallbackMessage);
    const mappedError = mapSearchError(rawMessage);
    const status = mapPauseReasonCodeToSearchStatus(reasonCode, mappedError.status);

    lastWorkbookSignature.current = '';
    setSearchData({
      status,
      workbook: null,
      modules: [],
      macros: [],
      shortcutAudit: null,
      error: {
        code: reasonCode,
        message: mappedError.message
      }
    });
  };

  const loadSearchData = useCallback(async ({ silent = false } = {}) => {
    if (searchLoadInFlight.current) {
      return;
    }

    searchLoadInFlight.current = true;
    const requestId = ++searchRequestSequence.current;

    if (!silent) {
      setSearchData((prev) => ({
        ...prev,
        status: 'loading',
        error: null
      }));
    }

    try {
      const workbookApi = window.excel?.workbook?.info;
      const workbookContextApi = window.excel?.workbook?.context;
      const modulesApi = window.excel?.vba?.modules;
      const proceduresApi = window.excel?.vba?.procedures;
      const fetchBundle = async () => {
        if (typeof workbookContextApi === 'function') {
          const contextResult = await workbookContextApi();

          return {
            workbookResult: contextResult?.success
              ? {
                  success: true,
                  name: contextResult?.workbook?.name,
                  path: contextResult?.workbook?.path,
                  activeSheet: contextResult?.workbook?.activeSheet,
                  sheets: contextResult?.workbook?.sheets
                }
              : {
                  success: false,
                  message: contextResult?.message || 'Unable to load workbook context.'
                },
            modulesResult: contextResult?.success
              ? {
                  success: true,
                  workbook: contextResult?.workbook || null,
                  modules: Array.isArray(contextResult?.modules) ? contextResult.modules : []
                }
              : {
                  success: false,
                  modules: [],
                  message: contextResult?.message || 'Unable to load workbook modules.'
                },
            proceduresResult: contextResult?.success
              ? {
                  success: true,
                  workbook: contextResult?.workbook || null,
                  procedures: Array.isArray(contextResult?.procedures) ? contextResult.procedures : []
                }
              : {
                  success: false,
                  procedures: [],
                  message: contextResult?.message || 'Unable to load workbook procedures.'
                },
            shortcutAuditResult: contextResult?.success
              ? contextResult?.shortcutAudit || { success: true, shortcuts: [], unmapped: [] }
              : { success: false, shortcuts: [], unmapped: [] },
            vbaLocked: contextResult?.vbaLocked === true
          };
        }

        const [workbookResult, modulesResult, proceduresResult] = await Promise.all([
          workbookApi(),
          modulesApi(),
          proceduresApi()
        ]);
        return {
          workbookResult,
          modulesResult,
          proceduresResult,
          shortcutAuditResult: null
        };
      };

      if ((!workbookApi || !modulesApi || !proceduresApi) && typeof workbookContextApi !== 'function') {
        const mappedError = mapSearchError('NO_EXCEL: Excel bridge API is unavailable.');
        if (requestId !== searchRequestSequence.current) {
          return;
        }
        if (isTerminalConnectionStatus(mappedError.status)) {
          setPausedState(mappedError.code);
        }

        lastWorkbookSignature.current = '';
        setSearchData({
          status: mappedError.status,
          workbook: null,
          modules: [],
          macros: [],
          shortcutAudit: null,
          error: {
            code: mappedError.code,
            message: mappedError.message
          }
        });
        return;
      }

      let {
        workbookResult,
        modulesResult,
        proceduresResult,
        shortcutAuditResult,
        vbaLocked: fetchVbaLocked
      } = await fetchBundle();

      if (requestId !== searchRequestSequence.current) {
        return;
      }

      let failedResults = [workbookResult, modulesResult, proceduresResult].filter(
        (result) => !result?.success
      );

      if (failedResults.length > 0) {
        let failureMessage = failedResults
          .map((result) => result?.message)
          .filter(Boolean)
          .join(' | ');
        let mappedError = mapSearchError(failureMessage);

        // One-shot: when multi_instance is first detected, try Solution 1
        // (C# helper enumerates instances and activates the correct one)
        if (mappedError.status === 'multi_instance' && !resolveInstanceAttempted.current) {
          resolveInstanceAttempted.current = true;
          try {
            const resolved = await window.excel?.resolveInstance();
            if (resolved?.resolved) {
              ({
                workbookResult,
                modulesResult,
                proceduresResult,
                shortcutAuditResult
              } = await fetchBundle());
              if (requestId !== searchRequestSequence.current) {
                return;
              }
              failedResults = [workbookResult, modulesResult, proceduresResult].filter(
                (result) => !result?.success
              );
              if (failedResults.length > 0) {
                failureMessage = failedResults
                  .map((result) => result?.message)
                  .filter(Boolean)
                  .join(' | ');
                mappedError = mapSearchError(failureMessage);
              }
            }
          } catch {
            // Resolution failed - fall through to show multi_instance UI.
          }
        }

        if (failedResults.length > 0) {
          if (isTerminalConnectionStatus(mappedError.status)) {
            setPausedState(mappedError.code);
          }

          lastWorkbookSignature.current = '';
          setSearchData({
            status: mappedError.status,
            workbook: null,
            modules: [],
            macros: [],
            shortcutAudit: null,
            error: {
              code: mappedError.code,
              message: mappedError.message
            }
          });
          return;
        }
      }

      const fallbackWorkbook = modulesResult?.workbook || proceduresResult?.workbook || null;
      const workbook = normalizeWorkbook(workbookResult, fallbackWorkbook);
      const modules = normalizeModules(modulesResult?.modules, workbook);
      const macros = normalizeMacros(proceduresResult?.procedures);
      const shortcutAudit = shortcutAuditResult && typeof shortcutAuditResult === 'object'
        ? shortcutAuditResult
        : null;
      lastWorkbookSignature.current = `${workbook?.path || ''}::${workbook?.name || ''}`;
      lastFullSearchRefreshAt.current = Date.now();
      resolveInstanceAttempted.current = false;
      clearPausedState();

      setSearchData({
        status: 'ready',
        workbook,
        modules,
        macros,
        shortcutAudit,
        vbaLocked: fetchVbaLocked === true,
        error: null
      });
    } catch (error) {
      if (requestId !== searchRequestSequence.current) {
        return;
      }
      const mappedError = mapSearchError(error?.message);
      if (isTerminalConnectionStatus(mappedError.status)) {
        setPausedState(mappedError.code);
      }
      lastWorkbookSignature.current = '';
      setSearchData({
        status: mappedError.status,
        workbook: null,
        modules: [],
        macros: [],
        shortcutAudit: null,
        error: {
          code: mappedError.code,
          message: mappedError.message
        }
      });
    } finally {
      searchLoadInFlight.current = false;
    }
  }, []);

  const refreshSearchOnForeground = useCallback(async ({ trigger = 'interval' } = {}) => {
    if (!isSearchDataMode(mode) || runState === 'running' || Boolean(macroRunInFlightRef?.current) || Boolean(shortcutSaveInFlightRef?.current)) {
      return;
    }
    if (trigger === 'paused-loop' && !pollingPausedRef.current) {
      return;
    }

    const now = Date.now();
    if (shouldSkipForegroundRefresh({
      trigger,
      now,
      modeEntryAt: modeEntryAtRef.current,
      lastForegroundRefreshAt: lastForegroundRefreshAtRef.current
    })) {
      return;
    }
    if (trigger === 'focus' || trigger === 'visibility' || trigger === 'helper') {
      lastForegroundRefreshAtRef.current = now;
    }

    if (pollingPausedRef.current) {
      if (!shouldAttemptPausedReconnect({
        isPaused: pollingPausedRef.current,
        now,
        nextAttemptAt: pausedReconnectNextAttemptAtRef.current,
        inFlight: resumeAttemptInFlightRef.current
      })) {
        return;
      }

      resumeAttemptInFlightRef.current = true;
      lastResumeAttemptAtRef.current = now;
      try {
        const reconnect = await window.excel?.reconnect?.();
        if (reconnect?.success) {
          clearPausedState();
          await loadSearchData({ silent: true });
        } else {
          applyPausedReconnectFailure(reconnect);
        }
      } catch (error) {
        applyPausedReconnectFailure(error);
      } finally {
        resumeAttemptInFlightRef.current = false;
      }
      return;
    }

    // When in multi_instance state, try Solution 1 (helper resolution) first,
    // then fall back to reconnect (cache clear) for Solution 2
    if (searchData.status === 'multi_instance') {
      try {
        const resolved = await window.excel?.resolveInstance();
        if (!resolved?.resolved) {
          await window.excel?.reconnect();
        }
      } catch {
        try { await window.excel?.reconnect(); } catch { /* non-fatal */ }
      }
    }

    // Helper events already represent a meaningful Excel context transition,
    // so skip the lightweight ping and perform one direct refresh.
    if (trigger === 'helper') {
      await loadSearchData({ silent: true });
      return;
    }

    const workbookApi = window.excel?.workbook?.info;
    if (!workbookApi || workbookPingInFlight.current) {
      await loadSearchData({ silent: true });
      return;
    }

    workbookPingInFlight.current = true;
    try {
      const ping = await workbookApi();
      if (!ping?.success) {
        await loadSearchData({ silent: true });
        return;
      }

      const workbookSignature = `${String(ping?.path || '').trim()}::${String(ping?.name || '').trim()}`;
      const workbookChanged = workbookSignature !== lastWorkbookSignature.current;
      const staleThresholdMs = trigger === 'interval'
        ? SEARCH_PERIODIC_DEEP_REFRESH_STALE_MS
        : SEARCH_FULL_REFRESH_STALE_MS;
      const stale = Date.now() - lastFullSearchRefreshAt.current > staleThresholdMs;
      const shouldRefreshFull = searchData.status !== 'ready' || workbookChanged || stale;

      if (shouldRefreshFull) {
        await loadSearchData({ silent: true });
      }
    } catch (error) {
      await loadSearchData({ silent: true });
    } finally {
      workbookPingInFlight.current = false;
    }
  }, [loadSearchData, macroRunInFlightRef, shortcutSaveInFlightRef, mode, runState, searchData.status]);

  useEffect(() => {
    refreshSearchOnForegroundRef.current = refreshSearchOnForeground;
  }, [refreshSearchOnForeground]);

  useEffect(() => {
    const previousMode = previousModeRef.current;
    const didModeChange = previousMode !== mode;
    previousModeRef.current = mode;

    if (!isSearchDataMode(mode) || runState === 'running') {
      return undefined;
    }

    const now = Date.now();
    if (didModeChange || modeEntryAtRef.current === 0) {
      modeEntryAtRef.current = now;
      lastForegroundRefreshAtRef.current = 0;
    }

    if (shouldRefreshOnModeEntry({
      mode,
      previousMode,
      status: searchStatusRef.current,
      isPaused: pollingPausedRef.current,
      lastFullRefreshAt: lastFullSearchRefreshAt.current,
      now
    })) {
      void loadSearchData({ silent: true });
    }

    const handleFocus = () => {
      void refreshSearchOnForegroundRef.current?.({ trigger: 'focus' });
    };
    const handleVisibilityChange = () => {
      if (document.visibilityState === 'visible') {
        void refreshSearchOnForegroundRef.current?.({ trigger: 'visibility' });
      }
    };

    window.addEventListener('focus', handleFocus);
    document.addEventListener('visibilitychange', handleVisibilityChange);

    const unsubscribeForegroundChange = window.excel?.events?.onForegroundChanged?.((payload) => {
      const excelActive = Boolean(payload?.excelActive);
      if (excelActive) {
        return;
      }

      if (document.visibilityState !== 'visible') {
        return;
      }

      if (typeof document.hasFocus === 'function' && !document.hasFocus()) {
        return;
      }

      const now = Date.now();
      if (now - lastHelperRefreshAtRef.current < SEARCH_HELPER_EVENT_COOLDOWN_MS) {
        return;
      }
      lastHelperRefreshAtRef.current = now;

      if (helperRefreshDebounceTimerRef.current) {
        clearTimeout(helperRefreshDebounceTimerRef.current);
        helperRefreshDebounceTimerRef.current = null;
      }

      helperRefreshDebounceTimerRef.current = window.setTimeout(() => {
        helperRefreshDebounceTimerRef.current = null;
        if (helperRefreshInFlightRef.current) {
          return;
        }
        helperRefreshInFlightRef.current = true;
        Promise.resolve(refreshSearchOnForegroundRef.current?.({ trigger: 'helper' })).finally(() => {
          helperRefreshInFlightRef.current = false;
        });
      }, SEARCH_HELPER_EVENT_DEBOUNCE_MS);
    });

    const periodicId = setInterval(() => {
      void refreshSearchOnForegroundRef.current?.({ trigger: 'interval' });
    }, SEARCH_PERIODIC_REFRESH_MS);
    const pausedReconnectId = setInterval(() => {
      void refreshSearchOnForegroundRef.current?.({ trigger: 'paused-loop' });
    }, SEARCH_PAUSED_RECONNECT_TICK_MS);

    return () => {
      window.removeEventListener('focus', handleFocus);
      document.removeEventListener('visibilitychange', handleVisibilityChange);
      clearInterval(periodicId);
      clearInterval(pausedReconnectId);
      if (helperRefreshDebounceTimerRef.current) {
        clearTimeout(helperRefreshDebounceTimerRef.current);
        helperRefreshDebounceTimerRef.current = null;
      }
      if (typeof unsubscribeForegroundChange === 'function') {
        unsubscribeForegroundChange();
      }
    };
  }, [loadSearchData, mode, runState]);

  return {
    searchData,
    loadSearchData
  };
}

