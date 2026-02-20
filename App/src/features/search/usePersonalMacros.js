import { useCallback, useEffect, useRef, useState } from 'react';
import { normalizeMacros } from '../../lib/search-data.js';
import { PERSONAL_INITIAL_DEFER_MS, PERSONAL_REFRESH_TTL_MS } from './search-constants.js';

export const PERSONAL_WORKBOOK_NAME = 'PERSONAL.XLSB';

const INITIAL_PERSONAL_MACROS_STATE = {
  status: 'idle',
  macros: [],
  workbookFound: false,
  workbook: null,
  fileExists: false,
  workbookPath: '',
  error: null
};

export function shouldUsePersonalCache({
  cachedData,
  cachedAt,
  cachedSignature,
  nextSignature,
  now,
  ttlMs = PERSONAL_REFRESH_TTL_MS
}) {
  if (!cachedData) {
    return false;
  }

  if (String(cachedSignature || '') !== String(nextSignature || '')) {
    return false;
  }

  const ageMs = Number(now) - Number(cachedAt);
  return Number.isFinite(ageMs) && ageMs >= 0 && ageMs < ttlMs;
}

export function shouldDeferPersonalInitialFetch({
  status,
  previousStatus,
  hasDeferredInitialFetch
}) {
  return status === 'ready' && previousStatus !== 'ready' && !hasDeferredInitialFetch;
}

export function usePersonalMacros(searchData, workbookListSignature = '') {
  const [personalState, setPersonalState] = useState(INITIAL_PERSONAL_MACROS_STATE);
  const [refreshTick, setRefreshTick] = useState(0);
  const requestSequence = useRef(0);
  const cacheRef = useRef({
    fetchedAt: 0,
    workbookListSignature: '',
    data: null
  });
  const cacheTtlTimerRef = useRef(null);
  const previousStatusRef = useRef('idle');
  const hasDeferredInitialFetchRef = useRef(false);
  const refresh = useCallback(() => {
    setRefreshTick((previous) => previous + 1);
  }, []);

  useEffect(() => {
    if (cacheTtlTimerRef.current) {
      clearTimeout(cacheTtlTimerRef.current);
      cacheTtlTimerRef.current = null;
    }

    const currentStatus = String(searchData?.status || 'idle');
    const previousStatus = previousStatusRef.current;
    previousStatusRef.current = currentStatus;

    if (currentStatus !== 'ready') {
      requestSequence.current += 1;
      cacheRef.current = {
        fetchedAt: 0,
        workbookListSignature: '',
        data: null
      };
      setPersonalState(INITIAL_PERSONAL_MACROS_STATE);
      return;
    }

    if (shouldDeferPersonalInitialFetch({
      status: currentStatus,
      previousStatus,
      hasDeferredInitialFetch: hasDeferredInitialFetchRef.current
    })) {
      hasDeferredInitialFetchRef.current = true;
      setPersonalState((previous) => ({
        ...previous,
        status: 'loading',
        error: null
      }));
      cacheTtlTimerRef.current = window.setTimeout(() => {
        setRefreshTick((previous) => previous + 1);
      }, PERSONAL_INITIAL_DEFER_MS);
      return () => {
        if (cacheTtlTimerRef.current) {
          clearTimeout(cacheTtlTimerRef.current);
          cacheTtlTimerRef.current = null;
        }
      };
    }

    const proceduresByWorkbookApi = window.excel?.vba?.proceduresByWorkbook;
    const personalStatusApi = window.excel?.personal?.status;
    if (!proceduresByWorkbookApi || !personalStatusApi) {
      setPersonalState({
        status: 'error',
        macros: [],
        workbookFound: false,
        workbook: null,
        fileExists: false,
        workbookPath: '',
        error: { message: 'PERSONAL.XLSB APIs are unavailable.' }
      });
      return;
    }

    const normalizedWorkbookListSignature = String(workbookListSignature || '').trim();
    const now = Date.now();
    const cached = cacheRef.current.data;
    const cacheAgeMs = now - cacheRef.current.fetchedAt;
    const cacheIsFresh = shouldUsePersonalCache({
      cachedData: cached,
      cachedAt: cacheRef.current.fetchedAt,
      cachedSignature: cacheRef.current.workbookListSignature,
      nextSignature: normalizedWorkbookListSignature,
      now
    });

    if (cacheIsFresh) {
      setPersonalState(cached);
      const remainingMs = Math.max(50, PERSONAL_REFRESH_TTL_MS - cacheAgeMs);
      cacheTtlTimerRef.current = window.setTimeout(() => {
        setRefreshTick((previous) => previous + 1);
      }, remainingMs);
      return () => {
        if (cacheTtlTimerRef.current) {
          clearTimeout(cacheTtlTimerRef.current);
          cacheTtlTimerRef.current = null;
        }
      };
    }

    let cancelled = false;
    const requestId = ++requestSequence.current;

    setPersonalState((previous) => ({
      ...previous,
      status: 'loading',
      error: null
    }));

    (async () => {
      try {
        const statusResult = await personalStatusApi();
        if (cancelled || requestId !== requestSequence.current) {
          return;
        }

        if (!statusResult?.success) {
          const message = String(statusResult?.message || statusResult?.error || 'Unable to load PERSONAL.XLSB status.');
          setPersonalState({
            status: 'error',
            macros: [],
            workbookFound: false,
            workbook: null,
            fileExists: false,
            workbookPath: '',
            error: { message }
          });
          return;
        }

        let workbookFound = statusResult?.workbookFound !== false;
        const workbook = statusResult?.workbook || null;
        const workbookPath = String(statusResult?.workbookPath || workbook?.path || '').trim();
        const fileExists = statusResult?.fileExists === true || workbookFound;
        let macros = [];

        if (workbookFound) {
          const proceduresResult = await proceduresByWorkbookApi({
            workbookName: PERSONAL_WORKBOOK_NAME,
            workbookPath
          });
          if (cancelled || requestId !== requestSequence.current) {
            return;
          }

          if (!proceduresResult?.success) {
            const message = String(proceduresResult?.message || proceduresResult?.error || 'Unable to load PERSONAL.XLSB macros.');
            setPersonalState({
              status: 'error',
              macros: [],
              workbookFound: false,
              workbook: null,
              fileExists,
              workbookPath,
              error: { message }
            });
            return;
          }

          if (proceduresResult?.workbookFound === false) {
            workbookFound = false;
            macros = [];
          } else {
            macros = normalizeMacros(proceduresResult?.procedures);
          }
        }

        const nextState = {
          status: 'ready',
          macros,
          workbookFound,
          workbook,
          fileExists,
          workbookPath,
          error: null
        };
        cacheRef.current = {
          fetchedAt: Date.now(),
          workbookListSignature: normalizedWorkbookListSignature,
          data: nextState
        };
        setPersonalState(nextState);
      } catch (error) {
        if (cancelled || requestId !== requestSequence.current) {
          return;
        }
        const message = error?.message ? String(error.message) : 'Unable to load PERSONAL.XLSB macros.';
        setPersonalState((previous) => ({
          ...previous,
          status: 'error',
          workbookFound: false,
          workbook: null,
          fileExists: false,
          workbookPath: '',
          error: { message }
        }));
      }
    })();

    return () => {
      cancelled = true;
      if (cacheTtlTimerRef.current) {
        clearTimeout(cacheTtlTimerRef.current);
        cacheTtlTimerRef.current = null;
      }
    };
  }, [
    refreshTick,
    searchData?.status,
    workbookListSignature
  ]);

  return {
    ...personalState,
    refresh
  };
}
