import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { normalizeMacros, normalizeModules } from '../../lib/search-data.js';

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

const INITIAL_ALL_FILES_DATA = {
  status: 'idle',
  modules: [],
  error: null
};

export function getWorkbookKey(workbook) {
  const name = String(workbook?.name || '').trim();
  const path = String(workbook?.path || '').trim();
  return path || name || '';
}

function toWorkbookRequest(workbook) {
  return {
    workbookName: String(workbook?.name || '').trim(),
    workbookPath: String(workbook?.path || '').trim()
  };
}

function toWorkbookModel(workbook) {
  const name = String(workbook?.name || '').trim();
  const path = String(workbook?.path || '').trim();
  if (!name && !path) {
    return null;
  }

  return {
    name: name || 'Active Workbook',
    path,
    key: path || name
  };
}

export function normalizeListContextModules(rawModules = []) {
  const source = Array.isArray(rawModules) ? rawModules : [];
  const normalized = [];

  source.forEach((moduleItem) => {
    const workbook = {
      name: String(moduleItem?.workbookName || '').trim(),
      path: String(moduleItem?.workbookPath || '').trim()
    };
    const rows = normalizeModules([moduleItem], workbook);
    rows.forEach((row) => normalized.push(row));
  });

  return normalized;
}

export function sortWorkbooksForPicker(workbooks, activeWorkbookKey) {
  const source = Array.isArray(workbooks) ? workbooks.filter(Boolean) : [];
  const uniqueByKey = new Map();
  source.forEach((workbook) => {
    const key = String(workbook?.key || getWorkbookKey(workbook));
    if (!key || uniqueByKey.has(key)) {
      return;
    }
    uniqueByKey.set(key, {
      name: String(workbook?.name || '').trim() || 'Workbook',
      path: String(workbook?.path || '').trim(),
      key
    });
  });

  const rows = Array.from(uniqueByKey.values());
  rows.sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));

  if (!activeWorkbookKey) {
    return rows;
  }

  const activeIndex = rows.findIndex((row) => row.key === activeWorkbookKey);
  if (activeIndex <= 0) {
    return rows;
  }

  const [activeWorkbook] = rows.splice(activeIndex, 1);
  rows.unshift(activeWorkbook);
  return rows;
}

export function resolveSelectedWorkbookKey({ requestedKey, workbooks, activeWorkbookKey }) {
  const source = Array.isArray(workbooks) ? workbooks : [];
  const normalizedRequested = String(requestedKey || '').trim();
  const normalizedActive = String(activeWorkbookKey || '').trim();
  if (normalizedRequested && source.some((workbook) => workbook.key === normalizedRequested)) {
    return normalizedRequested;
  }
  if (normalizedActive && source.some((workbook) => workbook.key === normalizedActive)) {
    return normalizedActive;
  }
  return source[0]?.key || '';
}

export function qualifyWorkbookNameForRun(workbookName) {
  const safe = String(workbookName || '').trim();
  if (!safe) {
    return '';
  }
  if (/\s/.test(safe) || safe.includes("'")) {
    return `'${safe.replace(/'/g, "''")}'`;
  }
  return safe;
}

export function qualifyMacroFullName(workbookName, runTarget) {
  const safeRunTarget = String(runTarget || '').trim();
  if (!safeRunTarget) {
    return '';
  }
  if (safeRunTarget.includes('!')) {
    return safeRunTarget;
  }
  const qualifiedWorkbook = qualifyWorkbookNameForRun(workbookName);
  if (!qualifiedWorkbook) {
    return safeRunTarget;
  }
  return `${qualifiedWorkbook}!${safeRunTarget}`;
}

export function namespaceMacrosForWorkbook(macros, workbook) {
  const source = Array.isArray(macros) ? macros : [];
  const workbookKey = getWorkbookKey(workbook) || 'workbook';
  return source.map((macro) => {
    const rawMacroId = String(macro?.id || `${macro?.module || 'module'}.${macro?.name || 'macro'}`);
    const runTarget = String(macro?.runTarget || `${macro?.module || ''}.${macro?.name || ''}`).trim();
    return {
      ...macro,
      id: `${workbookKey}::macro::${rawMacroId}`,
      fullName: qualifyMacroFullName(workbook?.name, runTarget)
    };
  });
}

export function sortAllFilesModules(modules, activeWorkbookKey = '') {
  const source = Array.isArray(modules) ? modules.filter(Boolean) : [];
  const activeKey = String(activeWorkbookKey || '').trim();
  const activeRows = [];
  const nonActiveRows = [];

  source.forEach((moduleItem) => {
    const workbookKey = getWorkbookKey({
      name: moduleItem?.workbookName,
      path: moduleItem?.workbookPath
    });
    if (activeKey && workbookKey === activeKey) {
      activeRows.push(moduleItem);
      return;
    }
    nonActiveRows.push(moduleItem);
  });

  const compareByName = (a, b) => {
    const aName = String(a?.name || '');
    const bName = String(b?.name || '');
    return aName.localeCompare(bName, undefined, { sensitivity: 'base' });
  };

  activeRows.sort(compareByName);
  nonActiveRows.sort((a, b) => {
    const workbookCompare = String(a?.workbookName || '').localeCompare(
      String(b?.workbookName || ''),
      undefined,
      { sensitivity: 'base' }
    );
    if (workbookCompare !== 0) {
      return workbookCompare;
    }
    return compareByName(a, b);
  });

  return [...activeRows, ...nonActiveRows];
}

export function useWorkbookPickerData(searchData, options = {}) {
  const preferredWorkbookKey = String(options?.preferredWorkbookKey || '').trim();
  const [pickerState, setPickerState] = useState(INITIAL_PICKER_STATE);
  const [selectedWorkbookKey, setSelectedWorkbookKey] = useState('');
  const [selectedWorkbookData, setSelectedWorkbookData] = useState(INITIAL_SELECTED_WORKBOOK_DATA);
  const [allFilesData, setAllFilesData] = useState(INITIAL_ALL_FILES_DATA);
  const listRequestSequence = useRef(0);
  const workbookDataRequestSequence = useRef(0);
  const allFilesRequestSequence = useRef(0);

  const activeWorkbook = useMemo(
    () => toWorkbookModel(searchData?.workbook),
    [searchData?.workbook?.name, searchData?.workbook?.path]
  );
  const activeWorkbookKey = activeWorkbook?.key || '';

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

      if (typeof workbookListContextApi === 'function') {
        const contextModules = normalizeListContextModules(result?.allFilesModules);
        const fallbackModules = Array.isArray(searchData?.modules) ? searchData.modules : [];
        const mergedModules = contextModules.length > 0 ? contextModules : fallbackModules;
        setAllFilesData({
          status: 'ready',
          modules: sortAllFilesModules(mergedModules, activeWorkbookKey),
          error: null
        });
      }

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
  }, [activeWorkbook, activeWorkbookKey, preferredWorkbookKey, searchData?.modules, searchData?.status]);

  useEffect(() => {
    if (!preferredWorkbookKey) {
      return;
    }

    setSelectedWorkbookKey((previousKey) => (previousKey ? previousKey : preferredWorkbookKey));
  }, [preferredWorkbookKey]);

  useEffect(() => {
    if (searchData?.status !== 'ready') {
      listRequestSequence.current += 1;
      workbookDataRequestSequence.current += 1;
      allFilesRequestSequence.current += 1;
      setPickerState(INITIAL_PICKER_STATE);
      setSelectedWorkbookKey('');
      setSelectedWorkbookData(INITIAL_SELECTED_WORKBOOK_DATA);
      setAllFilesData(INITIAL_ALL_FILES_DATA);
      return;
    }

    refreshWorkbooks({ silent: true });
  }, [refreshWorkbooks, searchData?.status, searchData?.workbook?.name, searchData?.workbook?.path]);

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

  useEffect(() => {
    if (searchData?.status !== 'ready') {
      return;
    }

    if (!selectedWorkbook) {
      setSelectedWorkbookData({
        status: 'ready',
        workbook: null,
        modules: [],
        macros: [],
        error: null
      });
      return;
    }

    if (!isSelectedActiveWorkbook) {
      return;
    }

    setSelectedWorkbookData({
      status: 'ready',
      workbook: activeWorkbook,
      modules: Array.isArray(searchData?.modules) ? searchData.modules : [],
      macros: Array.isArray(searchData?.macros) ? searchData.macros : [],
      error: null
    });
  }, [
    activeWorkbook,
    isSelectedActiveWorkbook,
    searchData?.macros,
    searchData?.modules,
    searchData?.status,
    selectedWorkbook
  ]);

  useEffect(() => {
    if (searchData?.status !== 'ready') {
      return;
    }
    if (!selectedWorkbook || isSelectedActiveWorkbook) {
      return;
    }

    const modulesByWorkbookApi = window.excel?.vba?.modulesByWorkbook;
    const proceduresByWorkbookApi = window.excel?.vba?.proceduresByWorkbook;
    if (!modulesByWorkbookApi || !proceduresByWorkbookApi) {
      setSelectedWorkbookData({
        status: 'error',
        workbook: selectedWorkbook,
        modules: [],
        macros: [],
        error: { message: 'Workbook-scoped VBA APIs are unavailable.' }
      });
      return;
    }

    let cancelled = false;
    const requestId = ++workbookDataRequestSequence.current;
    setSelectedWorkbookData((previous) => ({
      ...previous,
      status: 'loading',
      workbook: selectedWorkbook,
      error: null
    }));

    (async () => {
      try {
        const workbookRequest = toWorkbookRequest(selectedWorkbook);
        const [modulesResult, proceduresResult] = await Promise.all([
          modulesByWorkbookApi(workbookRequest),
          proceduresByWorkbookApi(workbookRequest)
        ]);
        if (cancelled || requestId !== workbookDataRequestSequence.current) {
          return;
        }

        if (!modulesResult?.success || !proceduresResult?.success) {
          const message = [modulesResult?.message, proceduresResult?.message]
            .filter(Boolean)
            .join(' | ') || 'Unable to load workbook data.';
          setSelectedWorkbookData({
            status: 'error',
            workbook: selectedWorkbook,
            modules: [],
            macros: [],
            error: { message }
          });
          return;
        }

        const workbookFound = modulesResult?.workbookFound !== false && proceduresResult?.workbookFound !== false;
        if (!workbookFound) {
          setSelectedWorkbookData({
            status: 'ready',
            workbook: selectedWorkbook,
            modules: [],
            macros: [],
            error: null
          });
          await refreshWorkbooks({ silent: true });
          return;
        }

        const selectedWorkbookInfo = toWorkbookModel(
          modulesResult?.workbook || proceduresResult?.workbook || selectedWorkbook
        ) || selectedWorkbook;
        const normalizedWorkbook = {
          name: selectedWorkbookInfo.name,
          path: selectedWorkbookInfo.path
        };
        const modules = normalizeModules(modulesResult?.modules, normalizedWorkbook);
        const macros = namespaceMacrosForWorkbook(
          normalizeMacros(proceduresResult?.procedures),
          normalizedWorkbook
        );
        setSelectedWorkbookData({
          status: 'ready',
          workbook: selectedWorkbookInfo,
          modules,
          macros,
          error: null
        });
      } catch (error) {
        if (cancelled || requestId !== workbookDataRequestSequence.current) {
          return;
        }
        const message = error?.message ? String(error.message) : 'Unable to load workbook data.';
        setSelectedWorkbookData({
          status: 'error',
          workbook: selectedWorkbook,
          modules: [],
          macros: [],
          error: { message }
        });
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [
    isSelectedActiveWorkbook,
    refreshWorkbooks,
    searchData?.status,
    selectedWorkbook?.key,
    selectedWorkbook?.name,
    selectedWorkbook?.path
  ]);

  useEffect(() => {
    if (searchData?.status !== 'ready') {
      return;
    }

    const workbookListContextApi = window.excel?.workbook?.listContext;
    if (typeof workbookListContextApi === 'function' && pickerState.status !== 'error') {
      return;
    }

    const modulesByWorkbookApi = window.excel?.vba?.modulesByWorkbook;
    if (!modulesByWorkbookApi) {
      const activeModules = Array.isArray(searchData?.modules) ? searchData.modules : [];
      setAllFilesData({
        status: activeModules.length > 0 ? 'ready' : 'error',
        modules: sortAllFilesModules(activeModules, activeWorkbookKey),
        error: activeModules.length > 0 ? null : { message: 'Workbook-scoped module API is unavailable.' }
      });
      return;
    }

    const activeModules = Array.isArray(searchData?.modules) ? searchData.modules : [];
    const workbookRows = Array.isArray(pickerState.workbooks) ? pickerState.workbooks : [];
    const nonActiveWorkbooks = workbookRows.filter((workbook) => workbook?.key && workbook.key !== activeWorkbookKey);

    let cancelled = false;
    const requestId = ++allFilesRequestSequence.current;
    setAllFilesData((previous) => ({
      status: 'loading',
      modules: previous.modules,
      error: null
    }));

    (async () => {
      try {
        const moduleMap = new Map();
        activeModules.forEach((moduleItem) => {
          if (moduleItem?.id) {
            moduleMap.set(moduleItem.id, moduleItem);
          }
        });

        let shouldRefreshWorkbooks = false;
        const failures = [];
        const results = await Promise.all(
          nonActiveWorkbooks.map(async (workbook) => {
            const workbookRequest = toWorkbookRequest(workbook);
            try {
              const result = await modulesByWorkbookApi(workbookRequest);
              return { workbook, result, error: null };
            } catch (error) {
              return { workbook, result: null, error };
            }
          })
        );

        if (cancelled || requestId !== allFilesRequestSequence.current) {
          return;
        }

        results.forEach(({ workbook, result, error }) => {
          if (error) {
            failures.push(error?.message ? String(error.message) : `Unable to load modules for ${workbook?.name || 'workbook'}.`);
            return;
          }
          if (!result?.success) {
            failures.push(String(result?.message || `Unable to load modules for ${workbook?.name || 'workbook'}.`));
            return;
          }
          if (result?.workbookFound === false) {
            shouldRefreshWorkbooks = true;
            return;
          }

          const workbookInfo = toWorkbookModel(result?.workbook || workbook) || workbook;
          const normalizedWorkbook = {
            name: workbookInfo?.name || '',
            path: workbookInfo?.path || ''
          };
          const normalizedModules = normalizeModules(result?.modules, normalizedWorkbook);
          normalizedModules.forEach((moduleItem) => {
            if (moduleItem?.id) {
              moduleMap.set(moduleItem.id, moduleItem);
            }
          });
        });

        const mergedModules = sortAllFilesModules(Array.from(moduleMap.values()), activeWorkbookKey);
        const isError = mergedModules.length === 0 && failures.length > 0;
        setAllFilesData({
          status: isError ? 'error' : 'ready',
          modules: mergedModules,
          error: isError
            ? { message: failures.join(' | ') || 'Unable to load modules from open workbooks.' }
            : null
        });

        if (shouldRefreshWorkbooks) {
          await refreshWorkbooks({ silent: true });
        }
      } catch (error) {
        if (cancelled || requestId !== allFilesRequestSequence.current) {
          return;
        }
        const fallbackModules = sortAllFilesModules(activeModules, activeWorkbookKey);
        const message = error?.message ? String(error.message) : 'Unable to load modules from open workbooks.';
        setAllFilesData({
          status: fallbackModules.length > 0 ? 'ready' : 'error',
          modules: fallbackModules,
          error: fallbackModules.length > 0 ? null : { message }
        });
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [
    activeWorkbookKey,
    pickerState.status,
    pickerState.workbooks,
    refreshWorkbooks,
    searchData?.modules,
    searchData?.status
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
    allFilesData,
    workbookListSignature
  };
}
