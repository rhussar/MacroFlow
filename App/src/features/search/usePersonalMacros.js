import { useCallback, useEffect, useRef, useState } from 'react';
import { normalizeMacros } from '../../lib/search-data.js';
import {
  PERSONAL_FOREGROUND_REFRESH_COOLDOWN_MS,
  PERSONAL_INITIAL_DEFER_MS,
  PERSONAL_REFRESH_TTL_MS
} from './search-constants.js';

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
  includeShortcutAudit: false,
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
  cachedIncludeShortcutAudit = false,
  nextSignature,
  nextIncludeShortcutAudit = false,
  invalidateOnSignatureChange = true,
  now,
  ttlMs = PERSONAL_REFRESH_TTL_MS
}) {
  if (!cachedData) {
    return false;
  }

  const normalizedCachedSignature = String(cachedSignature || '').trim();
  const normalizedNextSignature = resolvePersonalCacheSignature(nextSignature, normalizedCachedSignature);
  if (invalidateOnSignatureChange && normalizedCachedSignature !== normalizedNextSignature) {
    return false;
  }

  if (cachedIncludeShortcutAudit !== nextIncludeShortcutAudit) {
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

export function normalizePersonalForegroundRefreshPolicy(policy, fallback = 'stale') {
  const normalizedFallback = String(fallback || '').trim().toLowerCase() || 'stale';
  const normalizedPolicy = String(policy || '').trim().toLowerCase() || normalizedFallback;
  if (normalizedPolicy === 'always' || normalizedPolicy === 'stale' || normalizedPolicy === 'never') {
    return normalizedPolicy;
  }
  return normalizedFallback === 'always' || normalizedFallback === 'never' ? normalizedFallback : 'stale';
}

export function shouldRefreshPersonalOnForeground({
  policy = 'stale',
  cachedData,
  cachedAt,
  cachedSignature,
  cachedIncludeShortcutAudit = false,
  nextSignature,
  nextIncludeShortcutAudit = false,
  now,
  ttlMs = PERSONAL_REFRESH_TTL_MS
}) {
  const normalizedPolicy = normalizePersonalForegroundRefreshPolicy(policy);
  if (normalizedPolicy === 'never') {
    return false;
  }
  if (normalizedPolicy === 'always') {
    return true;
  }
  return !shouldUsePersonalCache({
    cachedData,
    cachedAt,
    cachedSignature,
    cachedIncludeShortcutAudit,
    nextSignature,
    nextIncludeShortcutAudit,
    now,
    ttlMs
  });
}

export function usePersonalMacros(searchData, workbookListSignature = '', options = {}) {
  const includeShortcutAudit = options?.includeShortcutAudit === true;
  const focusRefreshPolicy = normalizePersonalForegroundRefreshPolicy(
    options?.focusRefreshPolicy,
    'stale'
  );
  const visibilityRefreshPolicy = normalizePersonalForegroundRefreshPolicy(
    options?.visibilityRefreshPolicy,
    focusRefreshPolicy
  );
  const foregroundRefreshCooldownMs = Number.isFinite(Number(options?.foregroundRefreshCooldownMs))
    ? Math.max(0, Number(options.foregroundRefreshCooldownMs))
    : PERSONAL_FOREGROUND_REFRESH_COOLDOWN_MS;
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
      cachedIncludeShortcutAudit: sharedPersonalCache.includeShortcutAudit,
      nextSignature: normalizedWorkbookListSignature,
      nextIncludeShortcutAudit: includeShortcutAudit,
      invalidateOnSignatureChange: false,
      now
    });

    return cacheIsFresh ? sharedPersonalCache.data : INITIAL_PERSONAL_MACROS_STATE;
  });
  const [refreshTick, setRefreshTick] = useState(0);
  const requestSequence = useRef(0);
  const deferredFetchTimerRef = useRef(null);
  const previousStatusRef = useRef('idle');
  const hasDeferredInitialFetchRef = useRef(false);
  const forceRefreshRef = useRef(false);
  const lastForegroundRefreshAtRef = useRef(0);
  const refresh = useCallback(() => {
    forceRefreshRef.current = true;
    setRefreshTick((previous) => previous + 1);
  }, []);

  useEffect(() => {
    if (deferredFetchTimerRef.current) {
      clearTimeout(deferredFetchTimerRef.current);
      deferredFetchTimerRef.current = null;
    }

    const currentStatus = String(searchData?.status || 'idle');
    const previousStatus = previousStatusRef.current;
    previousStatusRef.current = currentStatus;

    if (currentStatus !== 'ready') {
      requestSequence.current += 1;
      hasDeferredInitialFetchRef.current = false;
      sharedPersonalCache.fetchedAt = 0;
      sharedPersonalCache.workbookListSignature = '';
      sharedPersonalCache.includeShortcutAudit = false;
      sharedPersonalCache.data = null;
      forceRefreshRef.current = false;
      lastForegroundRefreshAtRef.current = 0;
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
    const cacheIsFresh = !forceRefresh && shouldUsePersonalCache({
      cachedData: cached,
      cachedAt: sharedPersonalCache.fetchedAt,
      cachedSignature: sharedPersonalCache.workbookListSignature,
      cachedIncludeShortcutAudit: sharedPersonalCache.includeShortcutAudit,
      nextSignature: normalizedWorkbookListSignature,
      nextIncludeShortcutAudit: includeShortcutAudit,
      invalidateOnSignatureChange: false,
      now
    });

    if (cacheIsFresh) {
      setPersonalState(cached);
      return undefined;
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
      deferredFetchTimerRef.current = window.setTimeout(() => {
        setRefreshTick((previous) => previous + 1);
      }, PERSONAL_INITIAL_DEFER_MS);
      return () => {
        if (deferredFetchTimerRef.current) {
          clearTimeout(deferredFetchTimerRef.current);
          deferredFetchTimerRef.current = null;
        }
      };
    }

    const proceduresByWorkbookApi = window.excel?.vba?.proceduresByWorkbook;
    const personalStatusApi = window.excel?.personal?.status;
    const auditShortcutsByWorkbookApi = window.excel?.vba?.auditShortcutsByWorkbook;
    const canLoadShortcutAudit = includeShortcutAudit && typeof auditShortcutsByWorkbookApi === 'function';
    if (typeof proceduresByWorkbookApi !== 'function' || typeof personalStatusApi !== 'function') {
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
        error: { message: 'PERSONAL.XLSB APIs are unavailable.' }
      });
      return;
    }

    let cancelled = false;
    const requestId = ++requestSequence.current;
    lastForegroundRefreshAtRef.current = Date.now();

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

        setPersonalState({
          status: 'loading',
          macros: [],
          shortcutAudit: null,
          workbookFound,
          workbook,
          fileExists,
          workbookPath,
          windowVisible,
          windowHidden,
          error: null
        });

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
              workbookFound,
              workbook,
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
            workbook = null;
            macros = [];
          } else {
            workbook = proceduresResult?.workbook || workbook;
            workbookPath = String(proceduresResult?.workbook?.path || workbookPath || '').trim();
            macros = normalizeMacros(proceduresResult?.procedures).map((m) => ({
              ...m,
              workbookName: PERSONAL_WORKBOOK_NAME,
              workbookPath,
              runTarget: `${PERSONAL_WORKBOOK_NAME}!${m.runTarget}`,
              fullName: `${PERSONAL_WORKBOOK_NAME}!${m.fullName}`
            }));
          }
        }

        if (workbookFound && canLoadShortcutAudit) {
          try {
            const auditResult = await auditShortcutsByWorkbookApi({
              workbookName: PERSONAL_WORKBOOK_NAME,
              workbookPath
            });
            if (cancelled || requestId !== requestSequence.current) {
              return;
            }

            shortcutAudit = auditResult && typeof auditResult === 'object'
              ? auditResult
              : null;
          } catch (error) {
            shortcutAudit = {
              success: false,
              shortcuts: [],
              unmapped: [],
              message: String(error?.message || 'Shortcut audit failed.')
            };
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
        sharedPersonalCache.includeShortcutAudit = includeShortcutAudit;
        sharedPersonalCache.data = nextState;
        setPersonalState(nextState);
      } catch (error) {
        if (cancelled || requestId !== requestSequence.current) {
          return;
        }
        const message = error?.message ? String(error.message) : 'Unable to load PERSONAL.XLSB macros.';
        sharedPersonalCache.fetchedAt = 0;
        sharedPersonalCache.workbookListSignature = '';
        sharedPersonalCache.includeShortcutAudit = false;
        sharedPersonalCache.data = null;
        setPersonalState((previous) => ({
          ...previous,
          status: 'error',
          macros: [],
          shortcutAudit: null,
          error: { message }
        }));
      }
    })();

    return () => {
      cancelled = true;
      if (deferredFetchTimerRef.current) {
        clearTimeout(deferredFetchTimerRef.current);
        deferredFetchTimerRef.current = null;
      }
    };
  }, [
    refreshTick,
    searchData?.status,
    workbookListSignature,
    includeShortcutAudit
  ]);

  useEffect(() => {
    if (String(searchData?.status || 'idle') !== 'ready') {
      return undefined;
    }
    if (focusRefreshPolicy === 'never' && visibilityRefreshPolicy === 'never') {
      return undefined;
    }

    const maybeRefresh = (policy) => {
      const now = Date.now();
      if (now - lastForegroundRefreshAtRef.current < foregroundRefreshCooldownMs) {
        return;
      }

      const normalizedWorkbookListSignature = resolvePersonalCacheSignature(
        workbookListSignature,
        sharedPersonalCache.workbookListSignature
      );
      const shouldRefresh = shouldRefreshPersonalOnForeground({
        policy,
        cachedData: sharedPersonalCache.data,
        cachedAt: sharedPersonalCache.fetchedAt,
        cachedSignature: sharedPersonalCache.workbookListSignature,
        cachedIncludeShortcutAudit: sharedPersonalCache.includeShortcutAudit,
        nextSignature: normalizedWorkbookListSignature,
        nextIncludeShortcutAudit: includeShortcutAudit,
        now
      });

      if (!shouldRefresh) {
        return;
      }

      lastForegroundRefreshAtRef.current = now;
      refresh();
    };

    const handleFocus = () => {
      maybeRefresh(focusRefreshPolicy);
    };
    const handleVisibilityChange = () => {
      if (document.visibilityState === 'visible') {
        maybeRefresh(visibilityRefreshPolicy);
      }
    };

    window.addEventListener('focus', handleFocus);
    document.addEventListener('visibilitychange', handleVisibilityChange);
    return () => {
      window.removeEventListener('focus', handleFocus);
      document.removeEventListener('visibilitychange', handleVisibilityChange);
    };
  }, [
    focusRefreshPolicy,
    foregroundRefreshCooldownMs,
    refresh,
    searchData?.status,
    includeShortcutAudit,
    visibilityRefreshPolicy,
    workbookListSignature
  ]);

  return {
    ...personalState,
    refresh
  };
}
