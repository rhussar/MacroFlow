import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  resolveSelectedWorkbookKey,
  sortWorkbooksForPicker,
  toWorkbookModel
} from '../workbooks/workbook-model.js';
import {
  INITIAL_WORKBOOK_SCOPED_DATA,
  useWorkbookScopedData
} from '../workbooks/useWorkbookScopedData.js';
import { WORKBOOK_PICKER_REFRESH_TTL_MS } from './search-constants.js';

const INITIAL_PICKER_STATE = {
  status: 'idle',
  workbooks: [],
  error: null
};

const INITIAL_SELECTED_WORKBOOK_DATA = {
  status: 'idle',
  workbook: null,
  modules: [],
  macros: [],
  error: null
};

const workbookPickerCache = new Map();

function buildWorkbookPickerCacheKey(activeWorkbookKey) {
  return String(activeWorkbookKey || '').trim() || '__active-workbook__';
}

function getFreshWorkbookPickerCacheEntry(activeWorkbookKey) {
  const cacheKey = buildWorkbookPickerCacheKey(activeWorkbookKey);
  const cachedEntry = workbookPickerCache.get(cacheKey);
  const cacheAgeMs = Date.now() - Number(cachedEntry?.fetchedAt || 0);
  const cacheIsFresh =
    Array.isArray(cachedEntry?.workbooks) &&
    Number.isFinite(cacheAgeMs) &&
    cacheAgeMs >= 0 &&
    cacheAgeMs < WORKBOOK_PICKER_REFRESH_TTL_MS;

  return cacheIsFresh ? cachedEntry : null;
}

export function useWorkbookPickerData(searchData, options = {}) {
  const preferredWorkbookKey = String(options?.preferredWorkbookKey || '').trim();
  const activeWorkbook = useMemo(
    () => toWorkbookModel(searchData?.workbook, { defaultName: 'Active Workbook' }),
    [searchData?.workbook?.name, searchData?.workbook?.path]
  );
  const activeWorkbookKey = activeWorkbook?.key || '';
  const freshPickerCacheEntry = getFreshWorkbookPickerCacheEntry(activeWorkbookKey);
  const [pickerState, setPickerState] = useState(() => {
    if (searchData?.status !== 'ready' || !freshPickerCacheEntry) {
      return INITIAL_PICKER_STATE;
    }

    return {
      status: 'ready',
      workbooks: freshPickerCacheEntry.workbooks,
      error: null
    };
  });
  const [selectedWorkbookKey, setSelectedWorkbookKey] = useState(() => {
    if (searchData?.status !== 'ready' || !freshPickerCacheEntry) {
      return '';
    }

    return resolveSelectedWorkbookKey({
      requestedKey: preferredWorkbookKey,
      workbooks: freshPickerCacheEntry.workbooks,
      activeWorkbookKey
    });
  });
  const listRequestSequence = useRef(0);

  const refreshWorkbooks = useCallback(async ({ silent = false } = {}) => {
    if (searchData?.status !== 'ready') {
      setPickerState(INITIAL_PICKER_STATE);
      return [];
    }

    const workbookListApi = window.excel?.workbook?.list;
    const workbookListContextApi = window.excel?.workbook?.listContext;
    if (!workbookListApi && typeof workbookListContextApi !== 'function') {
      setPickerState({
        status: 'error',
        workbooks: [],
        error: { message: 'Workbook list API is unavailable.' }
      });
      return [];
    }

    const requestId = ++listRequestSequence.current;
    if (!silent) {
      setPickerState((previous) => ({
        ...previous,
        status: 'loading',
        error: null
      }));
    }

    try {
      const result = typeof workbookListContextApi === 'function'
        ? await workbookListContextApi()
        : await workbookListApi();
      if (requestId !== listRequestSequence.current) {
        return [];
      }

      if (!result?.success) {
        const message = String(result?.message || result?.error || 'Unable to load open workbooks.');
        setPickerState({
          status: 'error',
          workbooks: [],
          error: { message }
        });
        return [];
      }

      const workbookModels = (Array.isArray(result?.workbooks) ? result.workbooks : [])
        .map((workbook) => toWorkbookModel(workbook))
        .filter(Boolean);

      if (activeWorkbook && !workbookModels.some((workbook) => workbook.key === activeWorkbook.key)) {
        workbookModels.unshift(activeWorkbook);
      }

      const sortedWorkbooks = sortWorkbooksForPicker(workbookModels, activeWorkbookKey);
      setPickerState({
        status: 'ready',
        workbooks: sortedWorkbooks,
        error: null
      });
      workbookPickerCache.set(buildWorkbookPickerCacheKey(activeWorkbookKey), {
        fetchedAt: Date.now(),
        workbooks: sortedWorkbooks
      });

      setSelectedWorkbookKey((previousKey) =>
        resolveSelectedWorkbookKey({
          requestedKey: previousKey || preferredWorkbookKey,
          workbooks: sortedWorkbooks,
          activeWorkbookKey
        })
      );
      return sortedWorkbooks;
    } catch (error) {
      if (requestId !== listRequestSequence.current) {
        return [];
      }
      const message = error?.message ? String(error.message) : 'Unable to load open workbooks.';
      setPickerState({
        status: 'error',
        workbooks: [],
        error: { message }
      });
      return [];
    }
  }, [activeWorkbook, activeWorkbookKey, preferredWorkbookKey, searchData?.status]);

  useEffect(() => {
    if (!preferredWorkbookKey) {
      return;
    }

    setSelectedWorkbookKey((previousKey) => (previousKey ? previousKey : preferredWorkbookKey));
  }, [preferredWorkbookKey]);

  useEffect(() => {
    if (searchData?.status !== 'ready') {
      listRequestSequence.current += 1;
      setPickerState(INITIAL_PICKER_STATE);
      setSelectedWorkbookKey('');
      return;
    }

    const cachedEntry = getFreshWorkbookPickerCacheEntry(activeWorkbookKey);
    if (cachedEntry) {
      const cachedWorkbooks = cachedEntry.workbooks;
      setPickerState({
        status: 'ready',
        workbooks: cachedWorkbooks,
        error: null
      });
      setSelectedWorkbookKey((previousKey) =>
        resolveSelectedWorkbookKey({
          requestedKey: previousKey || preferredWorkbookKey,
          workbooks: cachedWorkbooks,
          activeWorkbookKey
        })
      );
      return;
    }

    refreshWorkbooks({ silent: true });
  }, [activeWorkbookKey, preferredWorkbookKey, refreshWorkbooks, searchData?.status, searchData?.workbook?.name, searchData?.workbook?.path]);

  const selectedWorkbook = useMemo(() => {
    const source = Array.isArray(pickerState.workbooks) ? pickerState.workbooks : [];
    if (source.length === 0) {
      return activeWorkbook;
    }

    const resolvedKey = resolveSelectedWorkbookKey({
      requestedKey: selectedWorkbookKey,
      workbooks: source,
      activeWorkbookKey
    });
    return source.find((workbook) => workbook.key === resolvedKey) || activeWorkbook || source[0] || null;
  }, [activeWorkbook, activeWorkbookKey, pickerState.workbooks, selectedWorkbookKey]);

  const isSelectedActiveWorkbook =
    Boolean(selectedWorkbook?.key) && Boolean(activeWorkbookKey) && selectedWorkbook.key === activeWorkbookKey;
  const selectedWorkbookScopedData = useWorkbookScopedData({
    enabled: searchData?.status === 'ready' && Boolean(selectedWorkbook) && !isSelectedActiveWorkbook,
    workbook: selectedWorkbook,
    namespaceMacros: true,
    defaultWorkbookName: 'Active Workbook',
    missingWorkbookStatus: 'ready',
    onWorkbookMissing: useCallback(() => refreshWorkbooks({ silent: true }), [refreshWorkbooks])
  });
  const selectedWorkbookData = useMemo(() => {
    if (searchData?.status !== 'ready') {
      return INITIAL_SELECTED_WORKBOOK_DATA;
    }

    if (!selectedWorkbook) {
      return {
        status: 'ready',
        workbook: null,
        modules: [],
        macros: [],
        error: null
      };
    }

    if (isSelectedActiveWorkbook) {
      return {
        status: 'ready',
        workbook: activeWorkbook,
        modules: Array.isArray(searchData?.modules) ? searchData.modules : [],
        macros: Array.isArray(searchData?.macros) ? searchData.macros : [],
        error: null
      };
    }

    return selectedWorkbookScopedData.status === 'idle'
      ? {
          ...INITIAL_WORKBOOK_SCOPED_DATA,
          status: 'loading',
          workbook: selectedWorkbook
        }
      : selectedWorkbookScopedData;
  }, [
    activeWorkbook,
    isSelectedActiveWorkbook,
    searchData?.macros,
    searchData?.modules,
    searchData?.status,
    selectedWorkbook,
    selectedWorkbookScopedData
  ]);

  const workbookListSignature = useMemo(() => {
    const rows = Array.isArray(pickerState.workbooks) ? pickerState.workbooks : [];
    return rows
      .map((workbook) => String(workbook?.key || ''))
      .filter(Boolean)
      .sort()
      .join('|');
  }, [pickerState.workbooks]);

  const handleSelectedWorkbookChange = useCallback((nextWorkbookKey) => {
    setSelectedWorkbookKey(String(nextWorkbookKey || '').trim());
  }, []);

  return {
    pickerStatus: pickerState.status,
    pickerError: pickerState.error,
    workbooks: pickerState.workbooks,
    selectedWorkbook,
    selectedWorkbookKey: selectedWorkbook?.key || '',
    setSelectedWorkbookKey: handleSelectedWorkbookChange,
    refreshWorkbooks,
    selectedWorkbookData,
    isSelectedActiveWorkbook,
    workbookListSignature
  };
}
