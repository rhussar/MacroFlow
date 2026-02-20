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
  SEARCH_PERIODIC_REFRESH_MS
} from './search-constants.js';

export function isTerminalConnectionStatus(status) {
  return status === 'no_excel' || status === 'no_workbook';
}

export function shouldAttemptPausedReconnect({
  isPaused,
  trigger,
  now,
  lastResumeAttemptAt,
  inFlight,
  cooldownMs = SEARCH_FOCUS_REFRESH_COOLDOWN_MS
}) {
  if (!isPaused || inFlight) {
    return false;
  }

  if (trigger !== 'focus' && trigger !== 'visibility') {
    return false;
  }

  return now - lastResumeAttemptAt >= cooldownMs;
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

export function useSearchData({ mode, runState, macroRunInFlightRef, shortcutSaveInFlightRef }) {
  const [searchData, setSearchData] = useState(INITIAL_SEARCH_DATA);
  const searchRequestSequence = useRef(0);
  const searchLoadInFlight = useRef(false);
  const workbookPingInFlight = useRef(false);
  const pollingPausedRef = useRef(false);
  const pollingPausedReasonRef = useRef('');
  const resumeAttemptInFlightRef = useRef(false);
  const lastResumeAttemptAtRef = useRef(0);
  const modeEntryAtRef = useRef(0);
  const lastForegroundRefreshAtRef = useRef(0);
  const helperRefreshDebounceTimerRef = useRef(null);
  const helperRefreshInFlightRef = useRef(false);
  const lastHelperRefreshAtRef = useRef(0);
  const lastFullSearchRefreshAt = useRef(0);
  const lastWorkbookSignature = useRef('');
  const resolveInstanceAttempted = useRef(false);

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
              : { success: false, shortcuts: [], unmapped: [] }
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
        if (isTerminalConnectionStatus(mappedError.status) && !pollingPausedRef.current) {
          pollingPausedRef.current = true;
          pollingPausedReasonRef.current = String(mappedError.code || '').toUpperCase() || 'NO_EXCEL';
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
        shortcutAuditResult
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
          if (isTerminalConnectionStatus(mappedError.status) && !pollingPausedRef.current) {
            pollingPausedRef.current = true;
            pollingPausedReasonRef.current = String(mappedError.code || '').toUpperCase() || 'NO_EXCEL';
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
      pollingPausedRef.current = false;
      pollingPausedReasonRef.current = '';

      setSearchData({
        status: 'ready',
        workbook,
        modules,
        macros,
        shortcutAudit,
        error: null
      });
    } catch (error) {
      if (requestId !== searchRequestSequence.current) {
        return;
      }
      const mappedError = mapSearchError(error?.message);
      if (isTerminalConnectionStatus(mappedError.status) && !pollingPausedRef.current) {
        pollingPausedRef.current = true;
        pollingPausedReasonRef.current = String(mappedError.code || '').toUpperCase() || 'NO_EXCEL';
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
    if ((mode !== 'search' && mode !== 'explorer') || runState === 'running' || Boolean(macroRunInFlightRef?.current) || Boolean(shortcutSaveInFlightRef?.current)) {
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
        trigger,
        now,
        lastResumeAttemptAt: lastResumeAttemptAtRef.current,
        inFlight: resumeAttemptInFlightRef.current
      })) {
        return;
      }

      resumeAttemptInFlightRef.current = true;
      lastResumeAttemptAtRef.current = now;
      try {
        const reconnect = await window.excel?.reconnect?.();
        if (reconnect?.success) {
          pollingPausedRef.current = false;
          pollingPausedReasonRef.current = '';
          await loadSearchData({ silent: true });
        }
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
      const stale = Date.now() - lastFullSearchRefreshAt.current > SEARCH_FULL_REFRESH_STALE_MS;
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
    if ((mode !== 'search' && mode !== 'explorer') || runState === 'running') {
      return undefined;
    }

    modeEntryAtRef.current = Date.now();
    lastForegroundRefreshAtRef.current = 0;

    // Entering Search should always perform one full refresh.
    loadSearchData({ silent: true });

    const handleFocus = () => {
      refreshSearchOnForeground({ trigger: 'focus' });
    };
    const handleVisibilityChange = () => {
      if (document.visibilityState === 'visible') {
        refreshSearchOnForeground({ trigger: 'visibility' });
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
        Promise.resolve(refreshSearchOnForeground({ trigger: 'helper' })).finally(() => {
          helperRefreshInFlightRef.current = false;
        });
      }, SEARCH_HELPER_EVENT_DEBOUNCE_MS);
    });

    const periodicId = setInterval(() => {
      refreshSearchOnForeground({ trigger: 'interval' });
    }, SEARCH_PERIODIC_REFRESH_MS);

    return () => {
      window.removeEventListener('focus', handleFocus);
      document.removeEventListener('visibilitychange', handleVisibilityChange);
      clearInterval(periodicId);
      if (helperRefreshDebounceTimerRef.current) {
        clearTimeout(helperRefreshDebounceTimerRef.current);
        helperRefreshDebounceTimerRef.current = null;
      }
      if (typeof unsubscribeForegroundChange === 'function') {
        unsubscribeForegroundChange();
      }
    };
  }, [loadSearchData, mode, refreshSearchOnForeground, runState]);

  return {
    searchData,
    loadSearchData
  };
}

