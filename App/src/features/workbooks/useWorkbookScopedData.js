import { useEffect, useRef, useState } from 'react';
import { normalizeMacros, normalizeModules, mapSearchError } from '../../lib/search-data.js';
import { WORKBOOK_SCOPED_DATA_REFRESH_TTL_MS } from '../search/search-constants.js';
import {
  SEARCH_INVALIDATION_BUCKETS,
  getWorkbookScopedDataInvalidationScope,
  useSearchInvalidationRevision
} from '../search/search-invalidation.js';
import {
  namespaceMacrosForWorkbook,
  toWorkbookModel,
  toWorkbookRequest
} from './workbook-model.js';

export const INITIAL_WORKBOOK_SCOPED_DATA = {
  status: 'idle',
  workbook: null,
  modules: [],
  macros: [],
  error: null
};

const workbookScopedDataCache = new Map();

function buildWorkbookScopedApiUnavailableResult(workbook) {
  return {
    success: false,
    workbookFound: false,
    workbook,
    modules: [],
    macros: [],
    message: 'Workbook-scoped VBA APIs are unavailable.'
  };
}

function buildWorkbookScopedFailureMessage(modulesResult, proceduresResult) {
  return [modulesResult?.message, proceduresResult?.message]
    .filter(Boolean)
    .join(' | ') || 'Unable to load workbook data.';
}

function buildWorkbookScopedCacheKey(workbook, namespaceMacros) {
  const workbookKey = String(workbook?.key || workbook?.path || workbook?.name || '').trim();
  if (!workbookKey) {
    return '';
  }
  return `${workbookKey}::${namespaceMacros ? 'namespaced' : 'plain'}`;
}

function getFreshWorkbookScopedCacheEntry(workbook, namespaceMacros) {
  const cacheKey = buildWorkbookScopedCacheKey(workbook, namespaceMacros);
  if (!cacheKey) {
    return null;
  }

  const cachedEntry = workbookScopedDataCache.get(cacheKey);
  const cacheAgeMs = Date.now() - Number(cachedEntry?.fetchedAt || 0);
  const cacheIsFresh =
    Boolean(cachedEntry?.data) &&
    Number.isFinite(cacheAgeMs) &&
    cacheAgeMs >= 0 &&
    cacheAgeMs < WORKBOOK_SCOPED_DATA_REFRESH_TTL_MS;

  return cacheIsFresh ? cachedEntry : null;
}

export async function fetchWorkbookScopedData(workbook, options = {}) {
  const namespaceMacros = options?.namespaceMacros === true;
  const defaultWorkbookName = String(options?.defaultWorkbookName || '').trim() || 'Workbook';
  const requestedWorkbook = toWorkbookModel(workbook, { defaultName: defaultWorkbookName });

  if (!requestedWorkbook) {
    return {
      success: false,
      workbookFound: false,
      workbook: null,
      modules: [],
      macros: [],
      message: 'Workbook is required.'
    };
  }

  const modulesByWorkbookApi = window.excel?.vba?.modulesByWorkbook;
  const proceduresByWorkbookApi = window.excel?.vba?.proceduresByWorkbook;
  if (typeof modulesByWorkbookApi !== 'function' || typeof proceduresByWorkbookApi !== 'function') {
    return buildWorkbookScopedApiUnavailableResult(requestedWorkbook);
  }

  const workbookRequest = toWorkbookRequest(requestedWorkbook);
  const [modulesResult, proceduresResult] = await Promise.all([
    modulesByWorkbookApi(workbookRequest),
    proceduresByWorkbookApi(workbookRequest)
  ]);

  if (!modulesResult?.success || !proceduresResult?.success) {
    return {
      success: false,
      workbookFound: false,
      workbook: requestedWorkbook,
      modules: [],
      macros: [],
      message: buildWorkbookScopedFailureMessage(modulesResult, proceduresResult)
    };
  }

  const workbookFound = modulesResult?.workbookFound !== false && proceduresResult?.workbookFound !== false;
  if (!workbookFound) {
    return {
      success: true,
      workbookFound: false,
      workbook: requestedWorkbook,
      modules: [],
      macros: [],
      message: ''
    };
  }

  const resolvedWorkbook = toWorkbookModel(
    modulesResult?.workbook || proceduresResult?.workbook,
    {
      fallback: requestedWorkbook,
      defaultName: defaultWorkbookName
    }
  ) || requestedWorkbook;
  const normalizedWorkbook = {
    name: resolvedWorkbook.name,
    path: resolvedWorkbook.path
  };
  const modules = normalizeModules(modulesResult?.modules, normalizedWorkbook);
  const procedures = normalizeMacros(proceduresResult?.procedures);
  const macros = namespaceMacros
    ? namespaceMacrosForWorkbook(procedures, normalizedWorkbook)
    : procedures;

  return {
    success: true,
    workbookFound: true,
    workbook: resolvedWorkbook,
    modules,
    macros,
    message: ''
  };
}

export function useWorkbookScopedData(options = {}) {
  const {
    enabled = false,
    workbook = null,
    namespaceMacros = false,
    defaultWorkbookName = 'Workbook',
    missingWorkbookStatus = 'ready',
    missingWorkbookMessage = '',
    onWorkbookMissing = null
  } = options;
  const normalizedWorkbook = toWorkbookModel(workbook, { defaultName: defaultWorkbookName });
  const invalidationScope = getWorkbookScopedDataInvalidationScope(normalizedWorkbook, namespaceMacros);
  const workbookScopedInvalidationRevision = useSearchInvalidationRevision(
    SEARCH_INVALIDATION_BUCKETS.WORKBOOK_SCOPED_DATA,
    invalidationScope
  );
  const freshCacheEntry = enabled
    ? getFreshWorkbookScopedCacheEntry(normalizedWorkbook, namespaceMacros)
    : null;
  const [data, setData] = useState(() => {
    if (!enabled || !normalizedWorkbook || !freshCacheEntry?.data) {
      return INITIAL_WORKBOOK_SCOPED_DATA;
    }

    return freshCacheEntry.data;
  });
  const requestSequence = useRef(0);
  const lastHandledInvalidationRef = useRef(0);
  const onWorkbookMissingRef = useRef(onWorkbookMissing);
  onWorkbookMissingRef.current = onWorkbookMissing;

  useEffect(() => {
    if (!enabled) {
      requestSequence.current += 1;
      setData(INITIAL_WORKBOOK_SCOPED_DATA);
      return;
    }
    if (!normalizedWorkbook) {
      requestSequence.current += 1;
      setData(INITIAL_WORKBOOK_SCOPED_DATA);
      return;
    }

    const cacheKey = buildWorkbookScopedCacheKey(normalizedWorkbook, namespaceMacros);
    const invalidationChanged =
      workbookScopedInvalidationRevision > 0
      && workbookScopedInvalidationRevision !== lastHandledInvalidationRef.current;
    if (invalidationChanged) {
      lastHandledInvalidationRef.current = workbookScopedInvalidationRevision;
      if (cacheKey) {
        workbookScopedDataCache.delete(cacheKey);
      }
    }

    const cachedEntry = invalidationChanged
      ? null
      : getFreshWorkbookScopedCacheEntry(normalizedWorkbook, namespaceMacros);
    if (cachedEntry?.data) {
      setData(cachedEntry.data);
      return;
    }

    let cancelled = false;
    const requestId = ++requestSequence.current;
    setData((previous) => ({
      ...previous,
      status: 'loading',
      workbook: normalizedWorkbook,
      error: null
    }));

    (async () => {
      try {
        const result = await fetchWorkbookScopedData(normalizedWorkbook, {
          namespaceMacros,
          defaultWorkbookName
        });
        if (cancelled || requestId !== requestSequence.current) {
          return;
        }

        if (!result?.success) {
          const mapped = mapSearchError(result?.message);
          const nextData = {
            status: 'error',
            workbook: normalizedWorkbook,
            modules: [],
            macros: [],
            error: { message: mapped.message }
          };
          workbookScopedDataCache.set(cacheKey, {
            fetchedAt: Date.now(),
            data: nextData
          });
          setData(nextData);
          return;
        }

        if (result?.workbookFound === false) {
          const nextError = missingWorkbookStatus === 'error'
            ? {
                message: String(
                  missingWorkbookMessage || `The workbook "${normalizedWorkbook.name}" is no longer open.`
                )
              }
            : null;
          const nextData = {
            status: missingWorkbookStatus,
            workbook: normalizedWorkbook,
            modules: [],
            macros: [],
            error: nextError
          };
          workbookScopedDataCache.set(cacheKey, {
            fetchedAt: Date.now(),
            data: nextData
          });
          setData(nextData);

          if (typeof onWorkbookMissingRef.current === 'function') {
            void Promise.resolve(onWorkbookMissingRef.current());
          }
          return;
        }

        const nextData = {
          status: 'ready',
          workbook: result.workbook,
          modules: result.modules,
          macros: result.macros,
          error: null
        };
        workbookScopedDataCache.set(cacheKey, {
          fetchedAt: Date.now(),
          data: nextData
        });
        setData(nextData);
      } catch (error) {
        if (cancelled || requestId !== requestSequence.current) {
          return;
        }

        const mapped = mapSearchError(error?.message);
        const nextData = {
          status: 'error',
          workbook: normalizedWorkbook,
          modules: [],
          macros: [],
          error: { message: mapped.message }
        };
        workbookScopedDataCache.set(cacheKey, {
          fetchedAt: Date.now(),
          data: nextData
        });
        setData(nextData);
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [
    defaultWorkbookName,
    enabled,
    missingWorkbookMessage,
    missingWorkbookStatus,
    namespaceMacros,
    workbookScopedInvalidationRevision,
    workbook?.key,
    workbook?.name,
    workbook?.path
  ]);

  return data;
}
