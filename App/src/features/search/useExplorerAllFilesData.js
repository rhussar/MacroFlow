import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  normalizeListContextModules,
  sortAllFilesModules,
  sortWorkbooksForPicker,
  toWorkbookModel
} from '../workbooks/workbook-model.js';
import { WORKBOOK_PICKER_REFRESH_TTL_MS } from './search-constants.js';

const INITIAL_EXPLORER_ALL_FILES_STATE = {
  status: 'idle',
  workbooks: [],
  modules: [],
  error: null
};

const explorerAllFilesCache = new Map();

function buildExplorerAllFilesCacheKey(activeWorkbookKey) {
  return String(activeWorkbookKey || '').trim() || '__active-workbook__';
}

function getFreshExplorerAllFilesCacheEntry(activeWorkbookKey) {
  const cacheKey = buildExplorerAllFilesCacheKey(activeWorkbookKey);
  const cachedEntry = explorerAllFilesCache.get(cacheKey);
  const cacheAgeMs = Date.now() - Number(cachedEntry?.fetchedAt || 0);
  const cacheIsFresh =
    Boolean(cachedEntry?.data) &&
    Number.isFinite(cacheAgeMs) &&
    cacheAgeMs >= 0 &&
    cacheAgeMs < WORKBOOK_PICKER_REFRESH_TTL_MS;

  return cacheIsFresh ? cachedEntry : null;
}

function buildExplorerAllFilesFallback(activeWorkbook, searchData) {
  const fallbackWorkbooks = activeWorkbook ? [activeWorkbook] : [];
  const fallbackModules = sortAllFilesModules(
    Array.isArray(searchData?.modules) ? searchData.modules : [],
    activeWorkbook?.key || ''
  );

  return {
    status: 'ready',
    workbooks: fallbackWorkbooks,
    modules: fallbackModules,
    error: null
  };
}

export function useExplorerAllFilesData(searchData) {
  const requestSequence = useRef(0);
  const activeWorkbook = useMemo(
    () => toWorkbookModel(searchData?.workbook, { defaultName: 'Active Workbook' }),
    [searchData?.workbook?.name, searchData?.workbook?.path]
  );
  const activeWorkbookKey = activeWorkbook?.key || '';
  const freshCacheEntry = getFreshExplorerAllFilesCacheEntry(activeWorkbookKey);
  const [state, setState] = useState(() => {
    if (searchData?.status !== 'ready') {
      return INITIAL_EXPLORER_ALL_FILES_STATE;
    }

    return freshCacheEntry?.data || buildExplorerAllFilesFallback(activeWorkbook, searchData);
  });

  const refreshExplorerAllFiles = useCallback(async ({ silent = false } = {}) => {
    if (searchData?.status !== 'ready') {
      setState(INITIAL_EXPLORER_ALL_FILES_STATE);
      return INITIAL_EXPLORER_ALL_FILES_STATE;
    }

    const listContextApi = window.excel?.workbook?.listContext;
    if (typeof listContextApi !== 'function') {
      const fallback = buildExplorerAllFilesFallback(activeWorkbook, searchData);
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
      const result = await listContextApi();
      if (requestId !== requestSequence.current) {
        return INITIAL_EXPLORER_ALL_FILES_STATE;
      }

      if (!result?.success) {
        const fallback = buildExplorerAllFilesFallback(activeWorkbook, searchData);
        const nextState = {
          ...fallback,
          error: { message: String(result?.message || result?.error || 'Unable to load open workbooks.') }
        };
        setState(nextState);
        return nextState;
      }

      const workbookModels = (Array.isArray(result?.workbooks) ? result.workbooks : [])
        .map((workbook) => toWorkbookModel(workbook))
        .filter(Boolean);

      if (activeWorkbook && !workbookModels.some((workbook) => workbook.key === activeWorkbook.key)) {
        workbookModels.unshift(activeWorkbook);
      }

      const sortedWorkbooks = sortWorkbooksForPicker(workbookModels, activeWorkbookKey);
      const normalizedContextModules = normalizeListContextModules(result?.allFilesModules);
      const activeModules = Array.isArray(searchData?.modules) ? searchData.modules : [];
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

      const nextState = {
        status: 'ready',
        workbooks: sortedWorkbooks,
        modules: sortAllFilesModules(Array.from(moduleMap.values()), activeWorkbookKey),
        error: null
      };
      explorerAllFilesCache.set(buildExplorerAllFilesCacheKey(activeWorkbookKey), {
        fetchedAt: Date.now(),
        data: nextState
      });
      setState(nextState);
      return nextState;
    } catch (error) {
      if (requestId !== requestSequence.current) {
        return INITIAL_EXPLORER_ALL_FILES_STATE;
      }

      const fallback = buildExplorerAllFilesFallback(activeWorkbook, searchData);
      const nextState = {
        ...fallback,
        error: { message: error?.message ? String(error.message) : 'Unable to load open workbooks.' }
      };
      setState(nextState);
      return nextState;
    }
  }, [activeWorkbook, activeWorkbookKey, searchData]);

  useEffect(() => {
    if (searchData?.status !== 'ready') {
      requestSequence.current += 1;
      setState(INITIAL_EXPLORER_ALL_FILES_STATE);
      return;
    }

    const cachedEntry = getFreshExplorerAllFilesCacheEntry(activeWorkbookKey);
    if (cachedEntry?.data) {
      setState(cachedEntry.data);
      return;
    }

    void refreshExplorerAllFiles({ silent: true });
  }, [activeWorkbookKey, refreshExplorerAllFiles, searchData?.status, searchData?.workbook?.name, searchData?.workbook?.path]);

  return {
    ...state,
    refreshExplorerAllFiles
  };
}
