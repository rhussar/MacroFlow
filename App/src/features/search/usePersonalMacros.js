import { useCallback, useEffect, useRef, useState } from 'react';
import { normalizeMacros } from '../../lib/search-data.js';
import { PERSONAL_INITIAL_DEFER_MS, PERSONAL_REFRESH_TTL_MS } from './search-constants.js';

export const PERSONAL_WORKBOOK_NAME = 'PERSONAL.XLSB';

const INITIAL_PERSONAL_MACROS_STATE = {
  status: 'idle',
  macros: [],
  shortcutAudit: null,
  workbookFound: false,
  workbook: null,
  fileExists: false,
  workbookPath: '',
  windowVisible: null,
  windowHidden: false,
  error: null
};

const sharedPersonalCache = {
  fetchedAt: 0,
  workbookListSignature: '',
  data: null
};

export function resolvePersonalCacheSignature(nextSignature, cachedSignature = '') {
  const normalizedNextSignature = String(nextSignature || '').trim();
  if (normalizedNextSignature) {
    return normalizedNextSignature;
  }
  return String(cachedSignature || '').trim();
}

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

  const normalizedCachedSignature = String(cachedSignature || '').trim();
  const normalizedNextSignature = resolvePersonalCacheSignature(nextSignature, normalizedCachedSignature);
  if (normalizedCachedSignature !== normalizedNextSignature) {
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
  const [personalState, setPersonalState] = useState(() => {
    const currentStatus = String(searchData?.status || 'idle');
    if (currentStatus !== 'ready') {
      return INITIAL_PERSONAL_MACROS_STATE;
    }

    const now = Date.now();
    const normalizedWorkbookListSignature = resolvePersonalCacheSignature(
      workbookListSignature,
      sharedPersonalCache.workbookListSignature
    );
    const cacheIsFresh = shouldUsePersonalCache({
      cachedData: sharedPersonalCache.data,
      cachedAt: sharedPersonalCache.fetchedAt,
      cachedSignature: sharedPersonalCache.workbookListSignature,
      nextSignature: normalizedWorkbookListSignature,
      now
    });

    return cacheIsFresh ? sharedPersonalCache.data : INITIAL_PERSONAL_MACROS_STATE;
  });
  const [refreshTick, setRefreshTick] = useState(0);
  const requestSequence = useRef(0);
  const cacheTtlTimerRef = useRef(null);
  const previousStatusRef = useRef('idle');
  const hasDeferredInitialFetchRef = useRef(false);
  const forceRefreshRef = useRef(false);
  const refresh = useCallback(() => {
    forceRefreshRef.current = true;
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
      sharedPersonalCache.fetchedAt = 0;
      sharedPersonalCache.workbookListSignature = '';
      sharedPersonalCache.data = null;
      forceRefreshRef.current = false;
      setPersonalState(INITIAL_PERSONAL_MACROS_STATE);
      return;
    }

    const normalizedWorkbookListSignature = resolvePersonalCacheSignature(
      workbookListSignature,
      sharedPersonalCache.workbookListSignature
    );
    const now = Date.now();
    const forceRefresh = forceRefreshRef.current;
    forceRefreshRef.current = false;
    const cached = sharedPersonalCache.data;
    const cacheAgeMs = now - sharedPersonalCache.fetchedAt;
    const cacheIsFresh = !forceRefresh && shouldUsePersonalCache({
      cachedData: cached,
      cachedAt: sharedPersonalCache.fetchedAt,
      cachedSignature: sharedPersonalCache.workbookListSignature,
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

    if (
      PERSONAL_INITIAL_DEFER_MS > 0 &&
      shouldDeferPersonalInitialFetch({
      status: currentStatus,
      previousStatus,
      hasDeferredInitialFetch: hasDeferredInitialFetchRef.current
      })
    ) {
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

    const personalContextApi = window.excel?.personal?.context;
    const proceduresByWorkbookApi = window.excel?.vba?.proceduresByWorkbook;
    const personalStatusApi = window.excel?.personal?.status;
    if (typeof personalContextApi !== 'function' && (!proceduresByWorkbookApi || !personalStatusApi)) {
      setPersonalState({
        status: 'error',
        macros: [],
        shortcutAudit: null,
        workbookFound: false,
        workbook: null,
        fileExists: false,
        workbookPath: '',
        error: { message: 'PERSONAL.XLSB APIs are unavailable.' }
      });
      return;
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
        let workbookFound = false;
        let workbook = null;
        let workbookPath = '';
        let fileExists = false;
        let windowVisible = null;
        let windowHidden = false;
        let macros = [];
        let shortcutAudit = null;

        if (typeof personalContextApi === 'function') {
          const contextResult = await personalContextApi();
          if (cancelled || requestId !== requestSequence.current) {
            return;
          }

          if (!contextResult?.success) {
            const message = String(contextResult?.message || contextResult?.error || 'Unable to load PERSONAL.XLSB context.');
            setPersonalState({
              status: 'error',
              macros: [],
              shortcutAudit: null,
              workbookFound: false,
              workbook: null,
              fileExists: false,
              workbookPath: '',
              windowVisible: null,
              windowHidden: false,
              error: { message }
            });
            return;
          }

          workbookFound = contextResult?.workbookFound !== false;
          workbook = contextResult?.workbook || null;
          workbookPath = String(contextResult?.workbookPath || workbook?.path || '').trim();
          fileExists = contextResult?.fileExists === true || workbookFound;
          windowVisible = typeof contextResult?.windowVisible === 'boolean'
            ? contextResult.windowVisible
            : null;
          windowHidden = contextResult?.windowHidden === true;
          shortcutAudit = contextResult?.shortcutAudit && typeof contextResult.shortcutAudit === 'object'
            ? contextResult.shortcutAudit
            : null;

          if (workbookFound) {
            macros = normalizeMacros(contextResult?.procedures).map((m) => ({
              ...m,
              workbookName: PERSONAL_WORKBOOK_NAME,
              workbookPath,
              runTarget: `${PERSONAL_WORKBOOK_NAME}!${m.runTarget}`,
              fullName: `${PERSONAL_WORKBOOK_NAME}!${m.fullName}`
            }));
          }
        } else {
          const statusResult = await personalStatusApi();
          if (cancelled || requestId !== requestSequence.current) {
            return;
          }

          if (!statusResult?.success) {
            const message = String(statusResult?.message || statusResult?.error || 'Unable to load PERSONAL.XLSB status.');
            setPersonalState({
              status: 'error',
              macros: [],
              shortcutAudit: null,
              workbookFound: false,
              workbook: null,
              fileExists: false,
              workbookPath: '',
              windowVisible: null,
              windowHidden: false,
              error: { message }
            });
            return;
          }

          workbookFound = statusResult?.workbookFound !== false;
          workbook = statusResult?.workbook || null;
          workbookPath = String(statusResult?.workbookPath || workbook?.path || '').trim();
          fileExists = statusResult?.fileExists === true || workbookFound;
          windowVisible = typeof statusResult?.windowVisible === 'boolean'
            ? statusResult.windowVisible
            : null;
          windowHidden = statusResult?.windowHidden === true;

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
                shortcutAudit: null,
                workbookFound: false,
                workbook: null,
                fileExists,
                workbookPath,
                windowVisible,
                windowHidden,
                error: { message }
              });
              return;
            }

            if (proceduresResult?.workbookFound === false) {
              workbookFound = false;
              macros = [];
            } else {
              macros = normalizeMacros(proceduresResult?.procedures).map((m) => ({
                ...m,
                workbookName: PERSONAL_WORKBOOK_NAME,
                workbookPath,
                runTarget: `${PERSONAL_WORKBOOK_NAME}!${m.runTarget}`,
                fullName: `${PERSONAL_WORKBOOK_NAME}!${m.fullName}`
              }));
            }
          }
        }

        const nextState = {
          status: 'ready',
          macros,
          shortcutAudit,
          workbookFound,
          workbook,
          fileExists,
          workbookPath,
          windowVisible,
          windowHidden,
          error: null
        };
        sharedPersonalCache.fetchedAt = Date.now();
        sharedPersonalCache.workbookListSignature = normalizedWorkbookListSignature;
        sharedPersonalCache.data = nextState;
        setPersonalState(nextState);
      } catch (error) {
        if (cancelled || requestId !== requestSequence.current) {
          return;
        }
        const message = error?.message ? String(error.message) : 'Unable to load PERSONAL.XLSB macros.';
        sharedPersonalCache.fetchedAt = 0;
        sharedPersonalCache.workbookListSignature = '';
        sharedPersonalCache.data = null;
        setPersonalState((previous) => ({
          ...previous,
          status: 'error',
          shortcutAudit: null,
          workbookFound: false,
          workbook: null,
          fileExists: false,
          workbookPath: '',
          windowVisible: null,
          windowHidden: false,
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
