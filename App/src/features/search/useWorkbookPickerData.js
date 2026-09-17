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
import { createAsyncResourceStore } from './async-resource-store.js';
import {
  SEARCH_INVALIDATION_BUCKETS,
  useSearchInvalidationRevision
} from './search-invalidation.js';
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

const workbookPickerStore = createAsyncResourceStore({
  defaultKey: '__active-workbook__'
});

function buildWorkbookPickerCacheKey(activeWorkbookKey) {
  return String(activeWorkbookKey || '').trim() || '__active-workbook__';
}

function getFreshWorkbookPickerCacheEntry(activeWorkbookKey) {
  return workbookPickerStore.getFreshEntry({
    key: buildWorkbookPickerCacheKey(activeWorkbookKey),
    ttlMs: WORKBOOK_PICKER_REFRESH_TTL_MS,
    isValid: (value) => Array.isArray(value)
  });
}

function buildSortedWorkbookPickerRows(result, activeWorkbook, activeWorkbookKey) {
  const workbookModels = (Array.isArray(result?.workbooks) ? result.workbooks : [])
    .map((workbook) => toWorkbookModel(workbook))
    .filter(Boolean);

  if (activeWorkbook && !workbookModels.some((workbook) => workbook.key === activeWorkbook.key)) {
    workbookModels.unshift(activeWorkbook);
  }

  return sortWorkbooksForPicker(workbookModels, activeWorkbookKey);
}

function getWorkbookPickerRequestKey(activeWorkbookKey) {
  return buildWorkbookPickerCacheKey(activeWorkbookKey);
}

function getOrCreateWorkbookPickerRequest(activeWorkbookKey, requestFactory) {
  return workbookPickerStore.run(
    getWorkbookPickerRequestKey(activeWorkbookKey),
    requestFactory
  );
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
      workbooks: freshPickerCacheEntry.value,
      error: null
    };
  });
  const [selectedWorkbookKey, setSelectedWorkbookKey] = useState(() => {
    if (searchData?.status !== 'ready' || !freshPickerCacheEntry) {
      return '';
    }

    return resolveSelectedWorkbookKey({
      requestedKey: preferredWorkbookKey,
      workbooks: freshPickerCacheEntry.value,
      activeWorkbookKey
    });
  });
  const listRequestSequence = useRef(0);
  const lastHandledWorkbookListInvalidationRef = useRef(0);
  const workbookListInvalidationRevision = useSearchInvalidationRevision(
    SEARCH_INVALIDATION_BUCKETS.WORKBOOK_LIST
  );

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
      const sortedWorkbooks = await getOrCreateWorkbookPickerRequest(
        activeWorkbookKey,
        async () => {
          const result = typeof workbookListApi === 'function'
            ? await workbookListApi()
            : await workbookListContextApi();

          if (!result?.success) {
            throw new Error(String(result?.message || result?.error || 'Unable to load open workbooks.'));
          }

          return buildSortedWorkbookPickerRows(result, activeWorkbook, activeWorkbookKey);
        }
      );
      if (requestId !== listRequestSequence.current) {
        return [];
      }
      setPickerState({
        status: 'ready',
        workbooks: sortedWorkbooks,
        error: null
      });
      workbookPickerStore.setValue(
        buildWorkbookPickerCacheKey(activeWorkbookKey),
        sortedWorkbooks
      );

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

  const ensureFreshWorkbooks = useCallback(async ({ silent = true } = {}) => {
    if (searchData?.status !== 'ready') {
      setPickerState(INITIAL_PICKER_STATE);
      return [];
    }

    const cachedEntry = getFreshWorkbookPickerCacheEntry(activeWorkbookKey);
    if (cachedEntry?.value) {
      const cachedWorkbooks = cachedEntry.value;
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
      return cachedWorkbooks;
    }

    return refreshWorkbooks({ silent });
  }, [activeWorkbookKey, preferredWorkbookKey, refreshWorkbooks, searchData?.status]);

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
      const cachedWorkbooks = cachedEntry.value;
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
  }, [activeWorkbookKey, preferredWorkbookKey, refreshWorkbooks, searchData?.status]);

  useEffect(() => {
    if (workbookListInvalidationRevision === 0) {
      return;
    }
    if (workbookListInvalidationRevision === lastHandledWorkbookListInvalidationRef.current) {
      return;
    }

    lastHandledWorkbookListInvalidationRef.current = workbookListInvalidationRevision;
    workbookPickerStore.clear(buildWorkbookPickerCacheKey(activeWorkbookKey));

    if (searchData?.status === 'ready') {
      void refreshWorkbooks({ silent: true });
    }
  }, [activeWorkbookKey, refreshWorkbooks, searchData?.status, workbookListInvalidationRevision]);

  const selectedWorkbook = useMemo(() => {
    const source = Array.isArray(pickerState.workbooks) ? pickerState.workbooks : [];
    const isPersonal = (wb) => wb && String(wb.name || '').trim().toUpperCase() === 'PERSONAL.XLSB';

    if (source.length === 0) {
      return isPersonal(activeWorkbook) ? null : activeWorkbook;
    }

    const resolvedKey = resolveSelectedWorkbookKey({
      requestedKey: selectedWorkbookKey,
      workbooks: source,
      activeWorkbookKey
    });
    const matched = source.find((workbook) => workbook.key === resolvedKey);
    if (matched) {
      return matched;
    }
    const nonPersonalFallback = source.find((wb) => !isPersonal(wb));
    return nonPersonalFallback || (!isPersonal(activeWorkbook) ? activeWorkbook : null) || source[0] || null;
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
    ensureFreshWorkbooks,
    selectedWorkbookData,
    isSelectedActiveWorkbook,
    workbookListSignature
  };
}
