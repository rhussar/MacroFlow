import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  normalizeListContextModules,
  sortAllFilesModules,
  sortWorkbooksForPicker,
  toWorkbookModel
} from '../workbooks/workbook-model.js';
import { createAsyncResourceStore } from './async-resource-store.js';
import {
  SEARCH_INVALIDATION_BUCKETS,
  useSearchInvalidationRevision
} from './search-invalidation.js';
import { WORKBOOK_PICKER_REFRESH_TTL_MS } from './search-constants.js';

const INITIAL_EXPLORER_ALL_FILES_STATE = {
  status: 'idle',
  workbooks: [],
  modules: [],
  error: null
};

const explorerAllFilesStore = createAsyncResourceStore({
  defaultKey: '__active-workbook__'
});

function buildExplorerAllFilesCacheKey(activeWorkbookKey) {
  return String(activeWorkbookKey || '').trim() || '__active-workbook__';
}

function getFreshExplorerAllFilesCacheEntry(activeWorkbookKey) {
  return explorerAllFilesStore.getFreshEntry({
    key: buildExplorerAllFilesCacheKey(activeWorkbookKey),
    ttlMs: WORKBOOK_PICKER_REFRESH_TTL_MS
  });
}

function buildExplorerAllFilesFallback(activeWorkbook, searchData) {
  const fallbackModules = sortAllFilesModules(
    Array.isArray(searchData?.modules) ? searchData.modules : [],
    activeWorkbook?.key || ''
  );

  return {
    status: 'ready',
    workbooks: [],
    modules: fallbackModules,
    error: null
  };
}

function getExplorerAllFilesRequestKey(activeWorkbookKey) {
  return buildExplorerAllFilesCacheKey(activeWorkbookKey);
}

function getOrCreateExplorerAllFilesRequest(activeWorkbookKey, requestFactory) {
  return explorerAllFilesStore.run(
    getExplorerAllFilesRequestKey(activeWorkbookKey),
    requestFactory
  );
}

export function useExplorerAllFilesData(searchData) {
  const requestSequence = useRef(0);
  const lastHandledExplorerInvalidationRef = useRef(0);
  const activeModules = Array.isArray(searchData?.modules) ? searchData.modules : [];
  const searchStatus = searchData?.status;
  const explorerInvalidationRevision = useSearchInvalidationRevision(
    SEARCH_INVALIDATION_BUCKETS.EXPLORER_ALL_FILES
  );
  const activeWorkbook = useMemo(
    () => toWorkbookModel(searchData?.workbook, { defaultName: 'Active Workbook' }),
    [searchData?.workbook?.name, searchData?.workbook?.path]
  );
  const activeWorkbookKey = activeWorkbook?.key || '';
  const freshCacheEntry = getFreshExplorerAllFilesCacheEntry(activeWorkbookKey);
  const [state, setState] = useState(() => {
    if (searchStatus !== 'ready') {
      return INITIAL_EXPLORER_ALL_FILES_STATE;
    }

    return freshCacheEntry?.value || buildExplorerAllFilesFallback(activeWorkbook, { modules: activeModules });
  });

  const refreshExplorerAllFiles = useCallback(async ({ silent = false } = {}) => {
    if (searchStatus !== 'ready') {
      setState(INITIAL_EXPLORER_ALL_FILES_STATE);
      return INITIAL_EXPLORER_ALL_FILES_STATE;
    }

    const listContextApi = window.excel?.workbook?.listContext;
    if (typeof listContextApi !== 'function') {
      const fallback = buildExplorerAllFilesFallback(activeWorkbook, { modules: activeModules });
      setState(fallback);
      return fallback;
    }

    const requestId = ++requestSequence.current;
    if (!silent) {
      setState((previous) => ({
        ...previous,
        status: 'loading',
        error: null
      }));
    }

    try {
      const nextState = await getOrCreateExplorerAllFilesRequest(
        activeWorkbookKey,
        async () => {
          const result = await listContextApi();
          if (!result?.success) {
            const fallback = buildExplorerAllFilesFallback(activeWorkbook, { modules: activeModules });
            return {
              ...fallback,
              error: { message: String(result?.message || result?.error || 'Unable to load open workbooks.') }
            };
          }

          const workbookModels = (Array.isArray(result?.workbooks) ? result.workbooks : [])
            .map((workbook) => toWorkbookModel(workbook))
            .filter(Boolean);

          if (activeWorkbook && !workbookModels.some((workbook) => workbook.key === activeWorkbook.key)) {
            workbookModels.unshift(activeWorkbook);
          }

          const sortedWorkbooks = sortWorkbooksForPicker(workbookModels, activeWorkbookKey);
          const normalizedContextModules = normalizeListContextModules(result?.allFilesModules);
          const moduleMap = new Map();
          normalizedContextModules.forEach((moduleItem) => {
            if (moduleItem?.id) {
              moduleMap.set(moduleItem.id, moduleItem);
            }
          });
          activeModules.forEach((moduleItem) => {
            if (moduleItem?.id) {
              moduleMap.set(moduleItem.id, moduleItem);
            }
          });

          const resolvedState = {
            status: 'ready',
            workbooks: sortedWorkbooks,
            modules: sortAllFilesModules(Array.from(moduleMap.values()), activeWorkbookKey),
            error: null
          };
          explorerAllFilesStore.setValue(
            buildExplorerAllFilesCacheKey(activeWorkbookKey),
            resolvedState
          );
          return resolvedState;
        }
      );
      if (requestId !== requestSequence.current) {
        return INITIAL_EXPLORER_ALL_FILES_STATE;
      }
      setState(nextState);
      return nextState;
    } catch (error) {
      if (requestId !== requestSequence.current) {
        return INITIAL_EXPLORER_ALL_FILES_STATE;
      }

      const fallback = buildExplorerAllFilesFallback(activeWorkbook, { modules: activeModules });
      const nextState = {
        ...fallback,
        error: { message: error?.message ? String(error.message) : 'Unable to load open workbooks.' }
      };
      setState(nextState);
      return nextState;
    }
  }, [activeModules, activeWorkbook, activeWorkbookKey, searchStatus]);

  useEffect(() => {
    if (searchStatus !== 'ready') {
      requestSequence.current += 1;
      setState(INITIAL_EXPLORER_ALL_FILES_STATE);
      return;
    }

    const cachedEntry = getFreshExplorerAllFilesCacheEntry(activeWorkbookKey);
    if (cachedEntry?.value) {
      setState(cachedEntry.value);
      return;
    }

    void refreshExplorerAllFiles({ silent: true });
  }, [activeWorkbookKey, refreshExplorerAllFiles, searchStatus]);

  useEffect(() => {
    if (explorerInvalidationRevision === 0) {
      return;
    }
    if (explorerInvalidationRevision === lastHandledExplorerInvalidationRef.current) {
      return;
    }

    lastHandledExplorerInvalidationRef.current = explorerInvalidationRevision;
    explorerAllFilesStore.clear(buildExplorerAllFilesCacheKey(activeWorkbookKey));

    if (searchStatus === 'ready') {
      void refreshExplorerAllFiles({ silent: true });
    }
  }, [activeWorkbookKey, explorerInvalidationRevision, refreshExplorerAllFiles, searchStatus]);

  const workbookListSignature = useMemo(() => {
    const workbooks = Array.isArray(state.workbooks) ? state.workbooks : [];
    return workbooks
      .map((workbook) => String(workbook?.key || ''))
      .filter(Boolean)
      .sort()
      .join('|');
  }, [state.workbooks]);

  return {
    ...state,
    workbookListSignature,
    refreshExplorerAllFiles
  };
}
