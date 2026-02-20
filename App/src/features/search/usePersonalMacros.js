import { useEffect, useRef, useState } from 'react';
import { normalizeMacros } from '../../lib/search-data.js';
import { PERSONAL_INITIAL_DEFER_MS, PERSONAL_REFRESH_TTL_MS } from './search-constants.js';

export const PERSONAL_WORKBOOK_NAME = 'PERSONAL.XLSB';

const INITIAL_PERSONAL_MACROS_STATE = {
  status: 'idle',
  macros: [],
  workbookFound: false,
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
    if (!proceduresByWorkbookApi) {
      setPersonalState({
        status: 'error',
        macros: [],
        workbookFound: false,
        error: { message: 'PERSONAL.XLSB macro API is unavailable.' }
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
        const result = await proceduresByWorkbookApi({ workbookName: PERSONAL_WORKBOOK_NAME });
        if (cancelled || requestId !== requestSequence.current) {
          return;
        }

        if (!result?.success) {
          const message = String(result?.message || result?.error || 'Unable to load PERSONAL.XLSB macros.');
          setPersonalState({
            status: 'error',
            macros: [],
            workbookFound: false,
            error: { message }
          });
          return;
        }

        const workbookFound = result?.workbookFound !== false;
        const macros = workbookFound ? normalizeMacros(result?.procedures) : [];
        const nextState = {
          status: 'ready',
          macros,
          workbookFound,
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

  return personalState;
}
