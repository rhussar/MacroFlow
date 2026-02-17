import { useState, useRef, useCallback, useEffect } from 'react';
import {
  mapSearchError,
  normalizeMacros,
  normalizeModules,
  normalizeWorkbook
} from '../../lib/search-data';
import {
  INITIAL_SEARCH_DATA,
  SEARCH_FOCUS_REFRESH_COOLDOWN_MS,
  SEARCH_FULL_REFRESH_STALE_MS,
  SEARCH_PERIODIC_REFRESH_MS
} from './search-constants';

export function useSearchData({ mode, runState, macroRunInFlightRef, shortcutSaveInFlightRef }) {
  const [searchData, setSearchData] = useState(INITIAL_SEARCH_DATA);
  const searchRequestSequence = useRef(0);
  const searchLoadInFlight = useRef(false);
  const workbookPingInFlight = useRef(false);
  const lastFocusRefreshAttemptAt = useRef(0);
  const lastFullSearchRefreshAt = useRef(0);
  const lastWorkbookSignature = useRef('');

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
      const modulesApi = window.excel?.vba?.modules;
      const proceduresApi = window.excel?.vba?.procedures;

      if (!workbookApi || !modulesApi || !proceduresApi) {
        const mappedError = mapSearchError('NO_EXCEL: Excel bridge API is unavailable.');
        if (requestId !== searchRequestSequence.current) {
          return;
        }
        lastWorkbookSignature.current = '';
        setSearchData({
          status: mappedError.status,
          workbook: null,
          modules: [],
          macros: [],
          error: {
            code: mappedError.code,
            message: mappedError.message
          }
        });
        return;
      }

      const fetchStart = performance.now();
      const [workbookResult, modulesResult, proceduresResult] = await Promise.all([
        workbookApi(),
        modulesApi(),
        proceduresApi()
      ]);
      console.log('[SearchData] fetch completed in %dms', Math.round(performance.now() - fetchStart));

      if (requestId !== searchRequestSequence.current) {
        return;
      }

      const failedResults = [workbookResult, modulesResult, proceduresResult].filter(
        (result) => !result?.success
      );

      if (failedResults.length > 0) {
        const failureMessage = failedResults
          .map((result) => result?.message)
          .filter(Boolean)
          .join(' | ');
        const mappedError = mapSearchError(failureMessage);
        lastWorkbookSignature.current = '';
        setSearchData({
          status: mappedError.status,
          workbook: null,
          modules: [],
          macros: [],
          error: {
            code: mappedError.code,
            message: mappedError.message
          }
        });
        return;
      }

      const fallbackWorkbook = modulesResult?.workbook || proceduresResult?.workbook || null;
      const workbook = normalizeWorkbook(workbookResult, fallbackWorkbook);
      const modules = normalizeModules(modulesResult?.modules, workbook);
      const macros = normalizeMacros(proceduresResult?.procedures);
      lastWorkbookSignature.current = `${workbook?.path || ''}::${workbook?.name || ''}`;
      lastFullSearchRefreshAt.current = Date.now();

      setSearchData({
        status: 'ready',
        workbook,
        modules,
        macros,
        error: null
      });
    } catch (error) {
      if (requestId !== searchRequestSequence.current) {
        return;
      }
      const mappedError = mapSearchError(error?.message);
      lastWorkbookSignature.current = '';
      setSearchData({
        status: mappedError.status,
        workbook: null,
        modules: [],
        macros: [],
        error: {
          code: mappedError.code,
          message: mappedError.message
        }
      });
    } finally {
      searchLoadInFlight.current = false;
    }
  }, []);

  const refreshSearchOnForeground = useCallback(async () => {
    if ((mode !== 'search' && mode !== 'explorer') || runState === 'running' || Boolean(macroRunInFlightRef?.current) || Boolean(shortcutSaveInFlightRef?.current)) {
      return;
    }

    const now = Date.now();
    if (now - lastFocusRefreshAttemptAt.current < SEARCH_FOCUS_REFRESH_COOLDOWN_MS) {
      return;
    }
    lastFocusRefreshAttemptAt.current = now;

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

    // Entering Search should always perform one full refresh.
    loadSearchData({ silent: true });

    const handleFocus = () => {
      refreshSearchOnForeground();
    };
    const handleVisibilityChange = () => {
      if (document.visibilityState === 'visible') {
        refreshSearchOnForeground();
      }
    };

    window.addEventListener('focus', handleFocus);
    document.addEventListener('visibilitychange', handleVisibilityChange);

    const periodicId = setInterval(() => {
      refreshSearchOnForeground();
    }, SEARCH_PERIODIC_REFRESH_MS);

    return () => {
      window.removeEventListener('focus', handleFocus);
      document.removeEventListener('visibilitychange', handleVisibilityChange);
      clearInterval(periodicId);
    };
  }, [loadSearchData, mode, refreshSearchOnForeground, runState]);

  return {
    searchData,
    loadSearchData
  };
}
