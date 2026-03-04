import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { FolderIcon, ReturnIcon, CloseIcon, ChevronDownIcon, WorkbookIcon, SearchIcon, FilePageIcon, AllFilesFolderIcon, InfoIcon, GlobeIcon } from './icons';
import {
  getSearchStatusView,
  selectAllFilesModules,
  selectActiveWorkbookMacros,
  selectPersonalGlobalMacros,
  selectPersonalGlobalSectionModel
} from '../features/search/search-selectors';
import { normalizeMacros } from '../lib/search-data.js';
import { formatShortcutPrefix } from '../lib/shortcut-keybind';
import { usePersonalMacros, PERSONAL_WORKBOOK_NAME } from '../features/search/usePersonalMacros';
import { useWorkbookPickerData } from '../features/search/useWorkbookPickerData';
import { useWorkbookShortcutState } from '../features/shortcuts/useWorkbookShortcutState';
import {
  buildWorkbookModuleRequest,
  canShowModuleContextActions,
  isValidVbaModuleName,
  shouldCommitModuleRename
} from '../features/search/module-actions';

const defaultSearchData = {
  status: 'idle',
  workbook: null,
  modules: [],
  macros: [],
  shortcutAudit: null,
  error: null
};

function qualifyWorkbookNameForRun(workbookName) {
  const safe = String(workbookName || '').trim();
  if (!safe) {
    return '';
  }
  if (/\s/.test(safe) || safe.includes("'")) {
    return `'${safe.replace(/'/g, "''")}'`;
  }
  return safe;
}

function qualifyMacroTargetForRun(workbookName, runTarget) {
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

const ALL_FILES_MACRO_LOAD_TIMEOUT_MS = 8000;
const ALL_FILES_LOADING_STALE_MS = ALL_FILES_MACRO_LOAD_TIMEOUT_MS + 1500;

function getAllFilesWorkbookCacheKey(workbookName, workbookPath) {
  const safeWorkbookPath = String(workbookPath || '').trim();
  const safeWorkbookName = String(workbookName || '').trim();
  return safeWorkbookPath || safeWorkbookName || '';
}

function getAllFilesMacroLoadErrorMessage(result) {
  const reason = String(result?.reason || '').trim().toLowerCase();
  const reasonCode = String(result?.reasonCode || '').trim().toUpperCase();
  const pollingPaused =
    result?.paused === true
    || reason === 'polling_paused'
    || reasonCode === 'NO_EXCEL'
    || reasonCode === 'NO_WORKBOOK'
    || reasonCode === 'NO_VISIBLE_WINDOWS';

  if (pollingPaused) {
    return 'Excel is reconnecting in the background. Reopen or focus the workbook, then try again.';
  }

  return String(result?.message || 'Unable to load macros for this workbook.');
}

function getAllFilesMacroLoadTimeoutResult() {
  return {
    success: false,
    workbookFound: false,
    macros: [],
    message: `Timed out loading workbook macros after ${ALL_FILES_MACRO_LOAD_TIMEOUT_MS}ms.`,
    reason: 'timeout'
  };
}

function workbookIdentityMatches({
  workbookNameA,
  workbookPathA,
  workbookNameB,
  workbookPathB
}) {
  const safePathA = String(workbookPathA || '').trim().toLowerCase();
  const safePathB = String(workbookPathB || '').trim().toLowerCase();
  if (safePathA && safePathB) {
    return safePathA === safePathB;
  }

  const safeNameA = String(workbookNameA || '').trim().toLowerCase();
  const safeNameB = String(workbookNameB || '').trim().toLowerCase();
  return Boolean(safeNameA) && Boolean(safeNameB) && safeNameA === safeNameB;
}

const SearchMode = ({
  searchQuery,
  onSearchChange,
  onBuildModeClick,
  onFileClick,
  onRunMacro,
  searchData = defaultSearchData,
  selectedMacroId = null,
  shortcutByMacroId = {},
  shortcutDraftByMacroId = {},
  shortcutInputErrorByMacroId = {},
  shortcutSavingMacroId = null,
  onShortcutDraftChange,
  onShortcutCommit,
  onActionStatus,
  shortcutSaveInFlightRef,
  selectedWorkbookForBuild = null,
  onSelectedWorkbookForBuildChange,
  onRefreshSearchData,
  onClose
}) => {
  const status = searchData.status || 'idle';
  const workbookPickerState = useWorkbookPickerData(searchData, {
    preferredWorkbookKey: selectedWorkbookForBuild?.key || ''
  });
  const [isWorkbookMenuOpen, setWorkbookMenuOpen] = useState(false);
  const [personalActionInFlight, setPersonalActionInFlight] = useState(false);
  const [moduleContextMenu, setModuleContextMenu] = useState(null);
  const [moduleRenameState, setModuleRenameState] = useState(null);
  const [moduleDeleteTarget, setModuleDeleteTarget] = useState(null);
  const [moduleActionInFlight, setModuleActionInFlight] = useState(false);
  const [expandedAllFilesModuleIds, setExpandedAllFilesModuleIds] = useState(() => new Set());
  const [allFilesMacrosByModuleId, setAllFilesMacrosByModuleId] = useState({});
  const [allFilesInfoHover, setAllFilesInfoHover] = useState(false);
  const [allFilesInfoPinned, setAllFilesInfoPinned] = useState(false);
  const [personalInfoHover, setPersonalInfoHover] = useState(false);
  const [personalInfoPinned, setPersonalInfoPinned] = useState(false);
  const workbookMenuRef = useRef(null);
  const moduleContextMenuRef = useRef(null);
  const allFilesInfoRef = useRef(null);
  const personalInfoRef = useRef(null);
  const renameCommitInFlightRef = useRef(false);
  const searchInputRef = useRef(null);
  const allFilesModuleRequestSequenceRef = useRef(0);
  const allFilesWorkbookMacroCacheRef = useRef(new Map());
  const allFilesWorkbookMacroInFlightRef = useRef(new Map());
  const showAllFilesInfo = allFilesInfoHover || allFilesInfoPinned;
  const showPersonalInfo = personalInfoHover || personalInfoPinned;

  // JS-driven drag handler for the search bar area.
  // Distinguishes click (focus input) from drag (move window).
  const handleSearchBarMouseDown = useCallback((e) => {
    // Only handle left button, and skip if input is already focused
    if (e.button !== 0) return;
    if (document.activeElement === searchInputRef.current) return;

    e.preventDefault();
    const startScreenX = e.screenX;
    const startScreenY = e.screenY;
    let lastScreenX = startScreenX;
    let lastScreenY = startScreenY;
    let dragging = false;
    const THRESHOLD = 3;

    const onMouseMove = (ev) => {
      const dx = ev.screenX - startScreenX;
      const dy = ev.screenY - startScreenY;
      if (!dragging && (Math.abs(dx) > THRESHOLD || Math.abs(dy) > THRESHOLD)) {
        dragging = true;
      }
      if (dragging) {
        const moveDx = ev.screenX - lastScreenX;
        const moveDy = ev.screenY - lastScreenY;
        lastScreenX = ev.screenX;
        lastScreenY = ev.screenY;
        window.excel?.window?.moveBy(moveDx, moveDy);
      }
    };

    const onMouseUp = () => {
      document.removeEventListener('mousemove', onMouseMove);
      document.removeEventListener('mouseup', onMouseUp);
      if (!dragging) {
        searchInputRef.current?.focus();
      }
    };

    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
  }, []);

  const selectedWorkbook = workbookPickerState.selectedWorkbook;
  const selectedWorkbookLabel = selectedWorkbook?.name || 'Active Workbook';

  const selectedWorkbookData = workbookPickerState.selectedWorkbookData;
  const selectedWorkbookErrorMessage = selectedWorkbookData.error?.message
    ? String(selectedWorkbookData.error.message)
    : 'Unable to load workbook data.';

  const usingActiveWorkbookShortcuts = workbookPickerState.isSelectedActiveWorkbook;

  const displayedWorkbookData = useMemo(() => {
    if (status !== 'ready') {
      return {
        status: 'idle',
        workbook: null,
        modules: [],
        macros: [],
        error: null
      };
    }

    if (usingActiveWorkbookShortcuts) {
      return {
        status: 'ready',
        workbook: searchData?.workbook || selectedWorkbook,
        modules: Array.isArray(searchData?.modules) ? searchData.modules : [],
        macros: Array.isArray(searchData?.macros) ? searchData.macros : [],
        error: null
      };
    }

    return selectedWorkbookData;
  }, [searchData?.macros, searchData?.modules, searchData?.workbook, selectedWorkbook, selectedWorkbookData, status, usingActiveWorkbookShortcuts]);

  const workbookShortcutState = useWorkbookShortcutState({
    enabled: status === 'ready' && !usingActiveWorkbookShortcuts && Boolean(selectedWorkbook?.name),
    workbook: selectedWorkbook,
    macros: displayedWorkbookData.macros,
    setActionStatus: onActionStatus,
    shortcutSaveInFlightRef
  });

  const effectiveShortcutByMacroId = usingActiveWorkbookShortcuts
    ? shortcutByMacroId
    : workbookShortcutState.shortcutByMacroId;
  const effectiveShortcutDraftByMacroId = usingActiveWorkbookShortcuts
    ? shortcutDraftByMacroId
    : workbookShortcutState.shortcutDraftByMacroId;
  const effectiveShortcutInputErrorByMacroId = usingActiveWorkbookShortcuts
    ? shortcutInputErrorByMacroId
    : workbookShortcutState.shortcutInputErrorByMacroId;
  const effectiveShortcutSavingMacroId = usingActiveWorkbookShortcuts
    ? shortcutSavingMacroId
    : workbookShortcutState.shortcutSavingMacroId;
  const handleEffectiveShortcutDraftChange = usingActiveWorkbookShortcuts
    ? onShortcutDraftChange
    : workbookShortcutState.handleShortcutDraftChange;
  const handleEffectiveShortcutCommit = usingActiveWorkbookShortcuts
    ? onShortcutCommit
    : workbookShortcutState.handleShortcutCommit;

  const filteredFiles = useMemo(
    () => selectAllFilesModules(workbookPickerState.allFilesData.modules, searchQuery),
    [workbookPickerState.allFilesData.modules, searchQuery]
  );

  const setAllFilesModuleMacroLoadState = useCallback((moduleId, requestToken, nextState) => {
    setAllFilesMacrosByModuleId((previous) => {
      const existing = previous[moduleId];
      if (!existing || existing.requestToken !== requestToken) {
        return previous;
      }
      return {
        ...previous,
        [moduleId]: nextState
      };
    });
  }, []);

  const raceWorkbookMacroPromiseWithTimeout = useCallback(async (promise) => {
    let timeoutId = null;
    try {
      const timeoutPromise = new Promise((resolve) => {
        timeoutId = window.setTimeout(() => {
          resolve(getAllFilesMacroLoadTimeoutResult());
        }, ALL_FILES_MACRO_LOAD_TIMEOUT_MS);
      });
      return await Promise.race([promise, timeoutPromise]);
    } finally {
      if (timeoutId) {
        window.clearTimeout(timeoutId);
      }
    }
  }, []);

  const loadWorkbookMacrosForAllFiles = useCallback(async ({ workbookName, workbookPath }) => {
    const safeWorkbookName = String(workbookName || '').trim();
    const safeWorkbookPath = String(workbookPath || '').trim();
    const workbookKey = getAllFilesWorkbookCacheKey(safeWorkbookName, safeWorkbookPath);
    if (!workbookKey) {
      return {
        success: false,
        workbookFound: false,
        macros: [],
        message: 'Workbook name or path is required.'
      };
    }

    const cachedMacros = allFilesWorkbookMacroCacheRef.current.get(workbookKey);
    if (Array.isArray(cachedMacros)) {
      return {
        success: true,
        workbookFound: true,
        macros: cachedMacros
      };
    }

    const inFlight = allFilesWorkbookMacroInFlightRef.current.get(workbookKey);
    if (inFlight) {
      return raceWorkbookMacroPromiseWithTimeout(inFlight);
    }

    const proceduresByWorkbookApi = window.excel?.vba?.proceduresByWorkbook;
    if (typeof proceduresByWorkbookApi !== 'function') {
      return {
        success: false,
        workbookFound: false,
        macros: [],
        message: 'Workbook procedure API is unavailable.'
      };
    }

    const workbookLoadPromise = (async () => {
      try {
        const result = await proceduresByWorkbookApi({
          workbookName: safeWorkbookName,
          workbookPath: safeWorkbookPath
        });

        if (!result?.success) {
          return {
            success: false,
            workbookFound: result?.workbookFound !== false,
            macros: [],
            message: getAllFilesMacroLoadErrorMessage(result)
          };
        }

        if (result?.workbookFound === false) {
          return {
            success: true,
            workbookFound: false,
            macros: []
          };
        }

        const macros = normalizeMacros(result?.procedures);
        allFilesWorkbookMacroCacheRef.current.set(workbookKey, macros);
        return {
          success: true,
          workbookFound: true,
          macros
        };
      } catch (error) {
        return {
          success: false,
          workbookFound: false,
          macros: [],
          message: error?.message ? String(error.message) : 'Unable to load macros for this workbook.'
        };
      }
    })();

    allFilesWorkbookMacroInFlightRef.current.set(workbookKey, workbookLoadPromise);
    workbookLoadPromise.finally(() => {
      if (allFilesWorkbookMacroInFlightRef.current.get(workbookKey) === workbookLoadPromise) {
        allFilesWorkbookMacroInFlightRef.current.delete(workbookKey);
      }
    });

    return raceWorkbookMacroPromiseWithTimeout(workbookLoadPromise);
  }, [raceWorkbookMacroPromiseWithTimeout]);

  const loadAllFilesModuleMacros = useCallback(async (moduleItem, { force = false } = {}) => {
    const moduleId = String(moduleItem?.id || '').trim();
    const moduleName = String(moduleItem?.name || '').trim();
    const workbookName = String(moduleItem?.workbookName || '').trim();
    const workbookPath = String(moduleItem?.workbookPath || '').trim();
    if (!moduleId || !moduleName || (!workbookName && !workbookPath)) {
      return;
    }

    const existing = allFilesMacrosByModuleId[moduleId];
    const hasCachedResult = existing?.status === 'ready';
    const loadingAgeMs = existing?.status === 'loading'
      ? (Date.now() - Number(existing?.requestedAt || 0))
      : Number.POSITIVE_INFINITY;
    const loadingIsFresh = existing?.status === 'loading'
      && Number.isFinite(loadingAgeMs)
      && loadingAgeMs >= 0
      && loadingAgeMs < ALL_FILES_LOADING_STALE_MS;
    if (!force && (hasCachedResult || loadingIsFresh)) {
      return;
    }

    const requestToken = ++allFilesModuleRequestSequenceRef.current;
    setAllFilesMacrosByModuleId((previous) => ({
      ...previous,
      [moduleId]: {
        status: 'loading',
        macros: [],
        error: null,
        requestToken,
        requestedAt: Date.now()
      }
    }));

    const displayedWorkbookName = String(displayedWorkbookData?.workbook?.name || '').trim();
    const displayedWorkbookPath = String(displayedWorkbookData?.workbook?.path || '').trim();
    const canUseDisplayedWorkbookMacros =
      displayedWorkbookData?.status === 'ready'
      && workbookIdentityMatches({
        workbookNameA: workbookName,
        workbookPathA: workbookPath,
        workbookNameB: displayedWorkbookName,
        workbookPathB: displayedWorkbookPath
      })
      && Array.isArray(displayedWorkbookData?.macros);

    if (canUseDisplayedWorkbookMacros) {
      const displayedMacroRows = displayedWorkbookData.macros
        .filter((macro) => String(macro?.module || '').trim().toLowerCase() === moduleName.toLowerCase())
        .map((macro) => {
          const qualifiedFullName = qualifyMacroTargetForRun(workbookName, macro.runTarget);
          return {
            ...macro,
            id: `${moduleId}::macro::${macro.id}`,
            fullName: qualifiedFullName || macro.fullName || macro.runTarget,
            workbookName,
            workbookPath
          };
        })
        .sort((a, b) => String(a?.name || '').localeCompare(String(b?.name || ''), undefined, { sensitivity: 'base' }));

      setAllFilesModuleMacroLoadState(moduleId, requestToken, {
        status: 'ready',
        macros: displayedMacroRows,
        error: null
      });
      return;
    }

    try {
      const workbookResult = await loadWorkbookMacrosForAllFiles({
        workbookName,
        workbookPath
      });
      if (!workbookResult?.success) {
        setAllFilesModuleMacroLoadState(moduleId, requestToken, {
          status: 'error',
          macros: [],
          error: String(workbookResult?.message || 'Unable to load macros for this module.')
        });
        return;
      }

      const normalizedModuleName = moduleName.toLowerCase();
      const macroRows = (Array.isArray(workbookResult?.macros) ? workbookResult.macros : [])
        .filter((macro) => String(macro?.module || '').trim().toLowerCase() === normalizedModuleName)
        .map((macro) => {
          const qualifiedFullName = qualifyMacroTargetForRun(workbookName, macro.runTarget);
          return {
            ...macro,
            id: `${moduleId}::macro::${macro.id}`,
            fullName: qualifiedFullName || macro.fullName || macro.runTarget,
            workbookName,
            workbookPath
          };
        })
        .sort((a, b) => String(a?.name || '').localeCompare(String(b?.name || ''), undefined, { sensitivity: 'base' }));

      setAllFilesModuleMacroLoadState(moduleId, requestToken, {
        status: 'ready',
        macros: macroRows,
        error: null
      });
    } catch (error) {
      const message = error?.message ? String(error.message) : 'Unable to load macros for this module.';
      setAllFilesModuleMacroLoadState(moduleId, requestToken, {
        status: 'error',
        macros: [],
        error: message
      });
    }
  }, [
    allFilesMacrosByModuleId,
    displayedWorkbookData?.macros,
    displayedWorkbookData?.status,
    displayedWorkbookData?.workbook?.name,
    displayedWorkbookData?.workbook?.path,
    loadWorkbookMacrosForAllFiles,
    setAllFilesModuleMacroLoadState
  ]);

  const toggleAllFilesModuleExpanded = useCallback((moduleItem) => {
    const moduleId = String(moduleItem?.id || '').trim();
    if (!moduleId) {
      return;
    }

    const isExpanded = expandedAllFilesModuleIds.has(moduleId);
    setExpandedAllFilesModuleIds((previous) => {
      const next = new Set(previous);
      if (next.has(moduleId)) {
        next.delete(moduleId);
      } else {
        next.add(moduleId);
      }
      return next;
    });

    if (!isExpanded) {
      void loadAllFilesModuleMacros(moduleItem);
    }
  }, [expandedAllFilesModuleIds, loadAllFilesModuleMacros]);

  useEffect(() => {
    if (status !== 'ready' || expandedAllFilesModuleIds.size < 1) {
      return;
    }

    const modulesById = new Map(
      (Array.isArray(workbookPickerState.allFilesData.modules) ? workbookPickerState.allFilesData.modules : [])
        .map((moduleItem) => [String(moduleItem?.id || '').trim(), moduleItem])
    );

    expandedAllFilesModuleIds.forEach((moduleId) => {
      const normalizedModuleId = String(moduleId || '').trim();
      if (!normalizedModuleId) {
        return;
      }
      const moduleItem = modulesById.get(normalizedModuleId);
      if (!moduleItem) {
        return;
      }

      const currentState = allFilesMacrosByModuleId[normalizedModuleId];
      const isIdle = !currentState || currentState.status === 'idle';

      if (isIdle) {
        void loadAllFilesModuleMacros(moduleItem);
      }
    });
  }, [
    allFilesMacrosByModuleId,
    expandedAllFilesModuleIds,
    loadAllFilesModuleMacros,
    status,
    workbookPickerState.allFilesData.modules
  ]);

  useEffect(() => {
    if (status !== 'ready') {
      return;
    }

    const timerId = window.setInterval(() => {
      const now = Date.now();
      setAllFilesMacrosByModuleId((previous) => {
        let changed = false;
        const next = { ...previous };

        Object.entries(previous).forEach(([moduleId, moduleState]) => {
          if (moduleState?.status !== 'loading') {
            return;
          }

          const loadingAgeMs = now - Number(moduleState?.requestedAt || 0);
          const isFresh = Number.isFinite(loadingAgeMs)
            && loadingAgeMs >= 0
            && loadingAgeMs < ALL_FILES_LOADING_STALE_MS;
          if (isFresh) {
            return;
          }

          changed = true;
          next[moduleId] = {
            status: 'error',
            macros: [],
            error: 'Timed out loading macros. Collapse and expand to retry.'
          };
        });

        return changed ? next : previous;
      });
    }, 1000);

    return () => {
      window.clearInterval(timerId);
    };
  }, [status]);

  const activeMacroRows = useMemo(
    () => selectActiveWorkbookMacros(displayedWorkbookData.macros, searchQuery, effectiveShortcutByMacroId),
    [displayedWorkbookData.macros, effectiveShortcutByMacroId, searchQuery]
  );

  const personalMacrosState = usePersonalMacros(searchData, workbookPickerState.workbookListSignature);

  // Editable shortcut state for PERSONAL.XLSB – same pattern as the selected
  // workbook's useWorkbookShortcutState.  Enabled when the personal section is
  // visible (i.e. selected workbook is NOT PERSONAL.XLSB) and macros are loaded.
  const personalSectionVisible = String(selectedWorkbook?.name || '').trim().toUpperCase() !== PERSONAL_WORKBOOK_NAME;
  const personalShortcutState = useWorkbookShortcutState({
    enabled: status === 'ready' && personalSectionVisible
      && personalMacrosState.workbookFound && personalMacrosState.macros.length > 0,
    workbook: personalMacrosState.workbook || { name: PERSONAL_WORKBOOK_NAME, path: personalMacrosState.workbookPath },
    macros: personalMacrosState.macros,
    setActionStatus: onActionStatus,
    shortcutSaveInFlightRef
  });

  const personalMacroRows = useMemo(
    () => selectPersonalGlobalMacros(personalMacrosState.macros, searchQuery, personalShortcutState.shortcutByMacroId),
    [personalMacrosState.macros, personalShortcutState.shortcutByMacroId, searchQuery]
  );
  const personalSectionModel = useMemo(
    () => selectPersonalGlobalSectionModel({
      workbookName: selectedWorkbook?.name,
      rows: personalMacroRows,
      totalMacros: personalMacrosState.macros.length,
      status: personalMacrosState.status,
      workbookFound: personalMacrosState.workbookFound,
      fileExists: personalMacrosState.fileExists,
      error: personalMacrosState.error
    }),
    [
      personalMacroRows,
      personalMacrosState.error,
      personalMacrosState.fileExists,
      personalMacrosState.macros.length,
      personalMacrosState.status,
      personalMacrosState.workbookFound,
      selectedWorkbook?.name
    ]
  );

  const getPersonalActionButtonLabel = useCallback((action) => {
    if (action === 'create_global_macro') {
      return 'Create a macro +';
    }
    if (action === 'create_file') {
      return 'Create file +';
    }
    if (action === 'open_file') {
      return 'Open file +';
    }
    return '';
  }, []);

  const handlePersonalAction = useCallback(async (action) => {
    if (!action || personalActionInFlight) {
      return;
    }

    const workbookPath = String(personalMacrosState.workbook?.path || personalMacrosState.workbookPath || '').trim();
    const workbookTarget = {
      name: 'PERSONAL.XLSB',
      path: workbookPath,
      key: workbookPath || 'PERSONAL.XLSB'
    };

    if (action === 'create_global_macro') {
      onBuildModeClick?.(workbookTarget);
      return;
    }

    const personalApi = window.excel?.personal;
    if (!personalApi) {
      onActionStatus?.('error', 'PERSONAL.XLSB actions are unavailable.');
      return;
    }

    const runAction = action === 'create_file'
      ? personalApi.create
      : personalApi.open;
    if (typeof runAction !== 'function') {
      onActionStatus?.('error', 'PERSONAL.XLSB actions are unavailable.');
      return;
    }

    setPersonalActionInFlight(true);
    onActionStatus?.(
      'running',
      action === 'create_file' ? 'Creating PERSONAL.xlsb...' : 'Opening PERSONAL.xlsb...'
    );

    try {
      const result = await runAction();
      if (!result?.success) {
        const message = String(
          result?.message ||
          (action === 'create_file'
            ? 'Unable to create PERSONAL.xlsb.'
            : 'Unable to open PERSONAL.xlsb.')
        );
        onActionStatus?.('error', message);
        return;
      }

      await Promise.allSettled([
        Promise.resolve(personalMacrosState.refresh?.()),
        Promise.resolve(workbookPickerState.refreshWorkbooks?.({ silent: true }))
      ]);

      onActionStatus?.(
        'success',
        action === 'create_file'
          ? 'Created and opened PERSONAL.xlsb.'
          : (result.alreadyOpen ? 'PERSONAL.xlsb is already open.' : 'Opened PERSONAL.xlsb.')
      );
    } catch (error) {
      const message = error?.message
        ? String(error.message)
        : (action === 'create_file'
          ? 'Unable to create PERSONAL.xlsb.'
          : 'Unable to open PERSONAL.xlsb.');
      onActionStatus?.('error', message);
    } finally {
      setPersonalActionInFlight(false);
    }
  }, [
    onActionStatus,
    onBuildModeClick,
    personalActionInFlight,
    personalMacrosState.refresh,
    personalMacrosState.workbook?.path,
    personalMacrosState.workbookPath,
    workbookPickerState.refreshWorkbooks
  ]);

  const refreshModuleData = useCallback(async () => {
    await Promise.allSettled([
      typeof onRefreshSearchData === 'function'
        ? Promise.resolve(onRefreshSearchData({ silent: true }))
        : Promise.resolve(),
      Promise.resolve(workbookPickerState.refreshWorkbooks?.({ silent: true }))
    ]);
  }, [onRefreshSearchData, workbookPickerState.refreshWorkbooks]);

  const closeModuleContextMenu = useCallback(() => {
    setModuleContextMenu(null);
  }, []);

  const handleOpenModuleContextMenu = useCallback((event, moduleItem) => {
    event.preventDefault();
    event.stopPropagation();

    if (!canShowModuleContextActions(moduleItem)) {
      setModuleContextMenu(null);
      return;
    }

    setModuleContextMenu({
      x: Number(event.clientX) || 0,
      y: Number(event.clientY) || 0,
      module: moduleItem
    });
  }, []);

  const handleStartRenameModule = useCallback((moduleItem) => {
    setModuleRenameState({
      moduleId: String(moduleItem?.id || ''),
      module: moduleItem,
      draft: String(moduleItem?.name || '')
    });
    setModuleContextMenu(null);
  }, []);

  const handleRenameDraftChange = useCallback((nextDraft) => {
    setModuleRenameState((previous) => {
      if (!previous) {
        return previous;
      }
      return {
        ...previous,
        draft: String(nextDraft || '')
      };
    });
  }, []);

  const handleCancelRenameModule = useCallback(() => {
    setModuleRenameState(null);
  }, []);

  const handleCommitRenameModule = useCallback(async () => {
    if (renameCommitInFlightRef.current) {
      return;
    }

    const currentRename = moduleRenameState;
    if (!currentRename?.module) {
      return;
    }

    const currentName = String(currentRename.module.name || '').trim();
    const nextName = String(currentRename.draft || '').trim();
    if (!shouldCommitModuleRename({ currentName, nextName })) {
      setModuleRenameState(null);
      return;
    }

    if (!isValidVbaModuleName(nextName)) {
      onActionStatus?.('error', 'Invalid module name. Use letters, numbers, and underscores, and start with a letter.');
      return;
    }

    const renameApi = window.excel?.vba?.renameModuleByWorkbook;
    if (typeof renameApi !== 'function') {
      onActionStatus?.('error', 'Module rename API is unavailable.');
      return;
    }

    const request = buildWorkbookModuleRequest(currentRename.module, selectedWorkbook || searchData?.workbook);
    if (!request.workbookName && !request.workbookPath) {
      onActionStatus?.('error', 'Unable to resolve workbook for this module.');
      return;
    }

    renameCommitInFlightRef.current = true;
    setModuleActionInFlight(true);
    onActionStatus?.('running', `Renaming module "${currentName}"...`);
    try {
      const result = await renameApi({
        ...request,
        nextModuleName: nextName
      });
      if (!result?.success) {
        const message = String(result?.message || 'Unable to rename module.');
        onActionStatus?.('error', message);
        return;
      }

      setModuleRenameState(null);
      await refreshModuleData();
      onActionStatus?.('success', String(result?.message || `Renamed module "${currentName}" to "${nextName}".`));
    } catch (error) {
      const message = error?.message ? String(error.message) : 'Unable to rename module.';
      onActionStatus?.('error', message);
    } finally {
      setModuleActionInFlight(false);
      renameCommitInFlightRef.current = false;
    }
  }, [
    moduleRenameState,
    onActionStatus,
    refreshModuleData,
    searchData?.workbook,
    selectedWorkbook
  ]);

  const handleRequestDeleteModule = useCallback((moduleItem) => {
    setModuleDeleteTarget(moduleItem);
    setModuleContextMenu(null);
  }, []);

  const handleCancelDeleteModule = useCallback(() => {
    setModuleDeleteTarget(null);
  }, []);

  const handleConfirmDeleteModule = useCallback(async () => {
    if (!moduleDeleteTarget || moduleActionInFlight) {
      return;
    }

    const deleteApi = window.excel?.vba?.deleteModuleByWorkbook;
    if (typeof deleteApi !== 'function') {
      onActionStatus?.('error', 'Module delete API is unavailable.');
      return;
    }

    const request = buildWorkbookModuleRequest(moduleDeleteTarget, selectedWorkbook || searchData?.workbook);
    if (!request.workbookName && !request.workbookPath) {
      onActionStatus?.('error', 'Unable to resolve workbook for this module.');
      return;
    }

    const moduleName = String(moduleDeleteTarget?.name || '').trim();
    setModuleActionInFlight(true);
    onActionStatus?.('running', `Deleting module "${moduleName}"...`);
    try {
      const result = await deleteApi(request);
      if (!result?.success) {
        const message = String(result?.message || 'Unable to delete module.');
        onActionStatus?.('error', message);
        return;
      }

      setModuleDeleteTarget(null);
      setModuleRenameState((previous) => (
        previous?.moduleId === String(moduleDeleteTarget?.id || '') ? null : previous
      ));
      await refreshModuleData();
      onActionStatus?.('success', String(result?.message || `Deleted module "${moduleName}".`));
    } catch (error) {
      const message = error?.message ? String(error.message) : 'Unable to delete module.';
      onActionStatus?.('error', message);
    } finally {
      setModuleActionInFlight(false);
    }
  }, [
    moduleActionInFlight,
    moduleDeleteTarget,
    onActionStatus,
    refreshModuleData,
    searchData?.workbook,
    selectedWorkbook
  ]);

  useEffect(() => {
    if (!isWorkbookMenuOpen) {
      return;
    }

    const handlePointerDown = (event) => {
      if (workbookMenuRef.current && !workbookMenuRef.current.contains(event.target)) {
        setWorkbookMenuOpen(false);
      }
    };

    const handleEscape = (event) => {
      if (event.key === 'Escape') {
        setWorkbookMenuOpen(false);
      }
    };

    document.addEventListener('mousedown', handlePointerDown);
    window.addEventListener('keydown', handleEscape);
    return () => {
      document.removeEventListener('mousedown', handlePointerDown);
      window.removeEventListener('keydown', handleEscape);
    };
  }, [isWorkbookMenuOpen]);

  useEffect(() => {
    if (!moduleContextMenu) {
      return;
    }

    const handlePointerDown = (event) => {
      if (moduleContextMenuRef.current && !moduleContextMenuRef.current.contains(event.target)) {
        setModuleContextMenu(null);
      }
    };

    const handleEscape = (event) => {
      if (event.key === 'Escape') {
        setModuleContextMenu(null);
      }
    };

    document.addEventListener('mousedown', handlePointerDown);
    window.addEventListener('keydown', handleEscape);
    window.addEventListener('scroll', closeModuleContextMenu, true);
    return () => {
      document.removeEventListener('mousedown', handlePointerDown);
      window.removeEventListener('keydown', handleEscape);
      window.removeEventListener('scroll', closeModuleContextMenu, true);
    };
  }, [closeModuleContextMenu, moduleContextMenu]);

  useEffect(() => {
    if (!allFilesInfoPinned) {
      return;
    }

    const handlePointerDown = (event) => {
      if (allFilesInfoRef.current && !allFilesInfoRef.current.contains(event.target)) {
        setAllFilesInfoPinned(false);
      }
    };

    const handleEscape = (event) => {
      if (event.key === 'Escape') {
        setAllFilesInfoPinned(false);
      }
    };

    document.addEventListener('mousedown', handlePointerDown);
    window.addEventListener('keydown', handleEscape);
    return () => {
      document.removeEventListener('mousedown', handlePointerDown);
      window.removeEventListener('keydown', handleEscape);
    };
  }, [allFilesInfoPinned]);

  useEffect(() => {
    if (!personalInfoPinned) return;
    const handlePointerDown = (event) => {
      if (personalInfoRef.current && !personalInfoRef.current.contains(event.target)) {
        setPersonalInfoPinned(false);
      }
    };
    const handleEscape = (event) => {
      if (event.key === 'Escape') setPersonalInfoPinned(false);
    };
    document.addEventListener('mousedown', handlePointerDown);
    window.addEventListener('keydown', handleEscape);
    return () => {
      document.removeEventListener('mousedown', handlePointerDown);
      window.removeEventListener('keydown', handleEscape);
    };
  }, [personalInfoPinned]);

  useEffect(() => {
    if (status !== 'ready') {
      setWorkbookMenuOpen(false);
      setModuleContextMenu(null);
      setModuleRenameState(null);
      setModuleDeleteTarget(null);
      setExpandedAllFilesModuleIds(new Set());
      setAllFilesMacrosByModuleId({});
      allFilesWorkbookMacroCacheRef.current.clear();
      allFilesWorkbookMacroInFlightRef.current.clear();
      setAllFilesInfoHover(false);
      setAllFilesInfoPinned(false);
    }
  }, [status]);

  useEffect(() => {
    const allFilesModules = Array.isArray(workbookPickerState.allFilesData.modules)
      ? workbookPickerState.allFilesData.modules
      : [];
    const validModuleIds = new Set(
      allFilesModules
        .map((moduleItem) => String(moduleItem?.id || '').trim())
        .filter(Boolean)
    );
    const validWorkbookKeys = new Set(
      allFilesModules
        .map((moduleItem) => getAllFilesWorkbookCacheKey(moduleItem?.workbookName, moduleItem?.workbookPath))
        .filter(Boolean)
    );

    setExpandedAllFilesModuleIds((previous) => {
      let changed = false;
      const next = new Set();
      previous.forEach((id) => {
        if (validModuleIds.has(id)) {
          next.add(id);
        } else {
          changed = true;
        }
      });
      return changed ? next : previous;
    });

    setAllFilesMacrosByModuleId((previous) => {
      const nextEntries = Object.entries(previous).filter(([moduleId]) => validModuleIds.has(moduleId));
      if (nextEntries.length === Object.keys(previous).length) {
        return previous;
      }
      return Object.fromEntries(nextEntries);
    });

    Array.from(allFilesWorkbookMacroCacheRef.current.keys()).forEach((workbookKey) => {
      if (!validWorkbookKeys.has(workbookKey)) {
        allFilesWorkbookMacroCacheRef.current.delete(workbookKey);
      }
    });

    Array.from(allFilesWorkbookMacroInFlightRef.current.keys()).forEach((workbookKey) => {
      if (!validWorkbookKeys.has(workbookKey)) {
        allFilesWorkbookMacroInFlightRef.current.delete(workbookKey);
      }
    });
  }, [workbookPickerState.allFilesData.modules]);

  const toggleWorkbookMenu = () => {
    setWorkbookMenuOpen((previous) => {
      const next = !previous;
      if (next) {
        workbookPickerState.refreshWorkbooks({ silent: true });
      }
      return next;
    });
  };

  const handleWorkbookSelection = (workbookKey) => {
    workbookPickerState.setSelectedWorkbookKey(workbookKey);
    setWorkbookMenuOpen(false);
  };

  useEffect(() => {
    if (!onSelectedWorkbookForBuildChange) {
      return;
    }
    onSelectedWorkbookForBuildChange(selectedWorkbook || null);
  }, [
    onSelectedWorkbookForBuildChange,
    selectedWorkbook?.key,
    selectedWorkbook?.name,
    selectedWorkbook?.path
  ]);

  const renderNonReadyState = () => {
    const statusView = getSearchStatusView(searchData);

    return (
      <div className={`search-status-panel search-status-${statusView.status}`}>
        <div className="search-status-header">
          {statusView.isLoading && <span className="status-spinner" />}
          <span className="search-status-title">{statusView.title}</span>
        </div>
        <p className="search-status-message">{statusView.message}</p>
      </div>
    );
  };

  const workbookDataIsLoading = displayedWorkbookData.status === 'loading';
  const workbookDataHasError = displayedWorkbookData.status === 'error';
  const workbookDataIsReady = displayedWorkbookData.status === 'ready';
  const canRenderMacroRows = workbookDataIsReady || (workbookDataIsLoading && activeMacroRows.length > 0);
  const workbookMacroCount = Array.isArray(displayedWorkbookData.macros)
    ? displayedWorkbookData.macros.length
    : 0;

  const macrosEmptyMessage = workbookDataIsLoading
    ? `Loading macros from ${selectedWorkbookLabel}...`
    : workbookDataHasError
      ? selectedWorkbookErrorMessage
      : workbookMacroCount < 1
        ? 'This workbook has no macros.'
        : 'No macros match this search.';

  const allFilesDataIsLoading = workbookPickerState.allFilesData.status === 'loading';
  const allFilesDataHasError = workbookPickerState.allFilesData.status === 'error';
  const hasAnyAllFilesModules = workbookPickerState.allFilesData.modules.length > 0;
  const modulesEmptyMessage = allFilesDataIsLoading && !hasAnyAllFilesModules
    ? 'Loading modules from open workbooks...'
    : allFilesDataHasError && !hasAnyAllFilesModules
      ? (workbookPickerState.allFilesData.error?.message || 'Unable to load modules from open workbooks.')
      : 'No modules match this search.';

  const selectedMacroForRowHighlight = usingActiveWorkbookShortcuts ? selectedMacroId : null;

  const renderReadyState = () => (
    <>
      <section className="search-ready-section search-ready-section-first">
        <div className="macro-workbook-picker-wrap" ref={workbookMenuRef}>
          <div className="macro-workbook-picker-group">
            <button
              type="button"
              className={`macro-workbook-picker ${isWorkbookMenuOpen ? 'open' : ''}`}
              aria-label="Select workbook"
              aria-expanded={isWorkbookMenuOpen}
              onClick={toggleWorkbookMenu}
            >
              <FilePageIcon size={16} className="macro-workbook-picker-wb-icon" />
              <span className="macro-workbook-picker-label">{selectedWorkbookLabel}</span>
            </button>
            <button
              type="button"
              className={`macro-workbook-picker-chevron ${isWorkbookMenuOpen ? 'open' : ''}`}
              aria-label="Toggle workbook menu"
              onClick={toggleWorkbookMenu}
            >
              <ChevronDownIcon size={14} />
            </button>
          </div>

          {isWorkbookMenuOpen && (
            <div className="macro-workbook-menu" role="listbox" aria-label="Open workbooks">
              <div className="macro-workbook-menu-title">Select open workbook</div>
              <div className="macro-workbook-menu-divider" />

              {workbookPickerState.pickerStatus === 'loading' && workbookPickerState.workbooks.length === 0 && (
                <div className="macro-workbook-menu-state">Loading open workbooks...</div>
              )}

              {workbookPickerState.pickerStatus === 'error' && (
                <div className="macro-workbook-menu-state error">
                  {workbookPickerState.pickerError?.message || 'Unable to load open workbooks.'}
                </div>
              )}

              {workbookPickerState.workbooks.map((workbook) => (
                <button
                  key={workbook.key}
                  type="button"
                  className={`macro-workbook-option ${workbook.key === workbookPickerState.selectedWorkbookKey ? 'selected' : ''}`}
                  onClick={() => handleWorkbookSelection(workbook.key)}
                >
                  <span className="macro-workbook-option-icon">
                    <WorkbookIcon size={14} />
                  </span>
                  <span className="macro-workbook-option-label">{workbook.name}</span>
                </button>
              ))}

              {workbookPickerState.pickerStatus === 'ready' && workbookPickerState.workbooks.length === 0 && (
                <div className="macro-workbook-menu-state">No open workbooks found.</div>
              )}
            </div>
          )}
        </div>

        <div className="shortcuts-grid">
          {(workbookDataHasError || activeMacroRows.length === 0) && (
            <div className="search-empty-state search-empty-state-grid">{macrosEmptyMessage}</div>
          )}

          {canRenderMacroRows && activeMacroRows.map((row) => {
            const macro = row.macro;
            const currentShortcutLetter =
              effectiveShortcutDraftByMacroId[macro.id] ?? effectiveShortcutByMacroId[macro.id] ?? '';
            const hasInputError = Boolean(effectiveShortcutInputErrorByMacroId[macro.id]);
            const isSaving = effectiveShortcutSavingMacroId === macro.id;
            const shortcutPrefix = formatShortcutPrefix(currentShortcutLetter);

            return (
              <div
                key={row.uiId}
                className={`shortcut-item ${selectedMacroForRowHighlight === macro.id ? 'selected' : ''} ${isSaving ? 'saving' : ''}`}
              >
                <div className="shortcut-run-target">
                  <button
                    type="button"
                    className="shortcut-icon-btn"
                    onClick={() => onRunMacro?.(macro)}
                  >
                    <ReturnIcon size={20} />
                  </button>
                  <span className="shortcut-name">{macro.name}</span>
                </div>
                <div className="shortcut-binding" onClick={(event) => event.stopPropagation()}>
                  <span className="shortcut-prefix">Ctrl +</span>
                  {shortcutPrefix.includes('Shift') && (
                    <span className="shortcut-shift">Shift +</span>
                  )}
                  <input
                    type="text"
                    className={`shortcut-keycap-input ${currentShortcutLetter ? '' : 'is-empty'} ${hasInputError ? 'has-error' : ''}`}
                    value={currentShortcutLetter}
                    placeholder=""
                    maxLength={1}
                    autoCapitalize="off"
                    autoComplete="off"
                    spellCheck={false}
                    aria-label={`Shortcut letter for ${macro.name}`}
                    onChange={(event) => handleEffectiveShortcutDraftChange?.(macro.id, event.target.value)}
                    onBlur={() => handleEffectiveShortcutCommit?.(macro, 'blur')}
                    onKeyDown={(event) => {
                      if (event.key === 'Enter') {
                        event.preventDefault();
                        event.currentTarget.blur();
                      } else if (event.key === 'Escape') {
                        event.preventDefault();
                        handleEffectiveShortcutDraftChange?.(macro.id, effectiveShortcutByMacroId[macro.id] || '');
                        event.currentTarget.blur();
                      }
                    }}
                    onClick={(event) => event.stopPropagation()}
                    disabled={isSaving}
                  />
                </div>
              </div>
            );
          })}
        </div>
      </section>

      {!personalSectionModel.hidden && (
        <section className="search-ready-section">
          <div className="personal-picker-wrap">
            <div ref={personalInfoRef} className="personal-picker-group">
              <div className="personal-picker">
                <GlobeIcon size={16} className="personal-picker-icon" />
                <span className="personal-picker-label">PERSONAL.XLSB</span>
              </div>
              <button
                type="button"
                className="personal-info-btn"
                aria-label="About PERSONAL.XLSB"
                aria-expanded={showPersonalInfo}
                onMouseEnter={() => setPersonalInfoHover(true)}
                onMouseLeave={() => setPersonalInfoHover(false)}
                onFocus={() => setPersonalInfoHover(true)}
                onBlur={() => setPersonalInfoHover(false)}
                onClick={(event) => {
                  event.stopPropagation();
                  setPersonalInfoPinned((prev) => !prev);
                }}
              >
                <InfoIcon size={16} />
              </button>
              {showPersonalInfo && (
                <div className="personal-info-tooltip" role="tooltip">
                  PERSONAL.XLSB is a hidden workbook that opens automatically with Excel. Macros stored here are available globally across all workbooks.
                </div>
              )}
            </div>
          </div>

          {personalSectionModel.isEmpty ? (
            <div className="search-empty-state global-macros-empty">
              {personalSectionModel.action && (
                <button
                  type="button"
                  className="global-macros-empty-action"
                  onClick={() => handlePersonalAction(personalSectionModel.action)}
                  disabled={personalActionInFlight}
                >
                  {getPersonalActionButtonLabel(personalSectionModel.action)}
                </button>
              )}
              {!personalSectionModel.action && (
                <span className="global-macros-empty-text">
                  {personalSectionModel.emptyMessage}
                </span>
              )}
            </div>
          ) : (
            <div className="shortcuts-grid">
              {personalMacroRows.map((row) => {
                const macro = row.macro;
                const currentShortcutLetter =
                  personalShortcutState.shortcutDraftByMacroId[macro.id]
                    ?? personalShortcutState.shortcutByMacroId[macro.id] ?? '';
                const hasInputError = Boolean(personalShortcutState.shortcutInputErrorByMacroId[macro.id]);
                const isSaving = personalShortcutState.shortcutSavingMacroId === macro.id;
                const shortcutPrefix = formatShortcutPrefix(currentShortcutLetter);

                return (
                  <div
                    key={row.uiId}
                    className={`shortcut-item ${selectedMacroId === macro.id ? 'selected' : ''} ${isSaving ? 'saving' : ''}`}
                  >
                    <div className="shortcut-run-target">
                      <button
                        type="button"
                        className="shortcut-icon-btn"
                        onClick={() => onRunMacro?.(macro)}
                      >
                        <ReturnIcon size={20} />
                      </button>
                      <span className="shortcut-name">{macro.name}</span>
                    </div>
                    <div className="shortcut-binding" onClick={(event) => event.stopPropagation()}>
                      <span className="shortcut-prefix">Ctrl +</span>
                      {shortcutPrefix.includes('Shift') && (
                        <span className="shortcut-shift">Shift +</span>
                      )}
                      <input
                        type="text"
                        className={`shortcut-keycap-input ${currentShortcutLetter ? '' : 'is-empty'} ${hasInputError ? 'has-error' : ''}`}
                        value={currentShortcutLetter}
                        placeholder=""
                        maxLength={1}
                        autoCapitalize="off"
                        autoComplete="off"
                        spellCheck={false}
                        aria-label={`Shortcut letter for ${macro.name}`}
                        onChange={(event) => personalShortcutState.handleShortcutDraftChange(macro.id, event.target.value)}
                        onBlur={() => personalShortcutState.handleShortcutCommit(macro, 'blur')}
                        onKeyDown={(event) => {
                          if (event.key === 'Enter') {
                            event.preventDefault();
                            event.currentTarget.blur();
                          } else if (event.key === 'Escape') {
                            event.preventDefault();
                            personalShortcutState.handleShortcutDraftChange(macro.id, personalShortcutState.shortcutByMacroId[macro.id] || '');
                            event.currentTarget.blur();
                          }
                        }}
                        onClick={(event) => event.stopPropagation()}
                        disabled={isSaving}
                      />
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </section>
      )}

      <section className="search-ready-section">
        <div className="allfiles-picker-wrap">
          <div
            ref={allFilesInfoRef}
            className="allfiles-picker-group"
          >
            <div className="allfiles-picker">
              <AllFilesFolderIcon size={16} className="allfiles-picker-icon" />
              <span className="allfiles-picker-label">All Files</span>
            </div>
            <button
              type="button"
              className="allfiles-info-btn"
              aria-label="About All Files"
              aria-expanded={showAllFilesInfo}
              onMouseEnter={() => setAllFilesInfoHover(true)}
              onMouseLeave={() => setAllFilesInfoHover(false)}
              onFocus={() => setAllFilesInfoHover(true)}
              onBlur={() => setAllFilesInfoHover(false)}
              onClick={(event) => {
                event.stopPropagation();
                setAllFilesInfoPinned((previous) => !previous);
              }}
            >
              <InfoIcon size={16} />
            </button>
            {showAllFilesInfo && (
              <div className="allfiles-info-tooltip" role="tooltip">
                Browse all VBA modules and macros in the selected workbook.
              </div>
            )}
          </div>
        </div>

        <div className="file-list">
          {filteredFiles.length === 0 && (
            <div className="search-empty-state">{modulesEmptyMessage}</div>
          )}

          {status === 'ready' && filteredFiles.map((file) => {
            const moduleId = String(file?.id || '').trim();
            const isRenaming = moduleRenameState?.moduleId === file.id;
            const isExpanded = expandedAllFilesModuleIds.has(moduleId);
            const moduleWorkbookLabel = String(
              file?.workbookName || selectedWorkbook?.name || 'Workbook'
            ).trim();
            const moduleMacrosState = allFilesMacrosByModuleId[moduleId] || {
              status: 'idle',
              macros: [],
              error: null
            };

            return (
              <React.Fragment key={file.id}>
                <div
                  className="file-item file-item-module"
                  onClick={() => {
                    if (isRenaming) {
                      return;
                    }
                    onFileClick?.({
                      module: file,
                      workbook: {
                        name: String(file?.workbookName || selectedWorkbook?.name || '').trim(),
                        path: String(file?.workbookPath || selectedWorkbook?.path || '').trim(),
                        key: String(file?.workbookPath || file?.workbookName || selectedWorkbook?.key || '').trim()
                      },
                      moduleId: String(file?.id || '').trim(),
                      moduleName: String(file?.name || '').trim(),
                      source: 'all-open-workbooks'
                    });
                  }}
                  onContextMenu={(event) => handleOpenModuleContextMenu(event, file)}
                >
                  <button
                    type="button"
                    className={`allfiles-module-chevron ${isExpanded ? 'expanded' : ''}`}
                    aria-label={`Toggle macros in ${String(file?.name || 'module')}`}
                    aria-expanded={isExpanded}
                    onClick={(event) => {
                      event.stopPropagation();
                      toggleAllFilesModuleExpanded(file);
                    }}
                  >
                    <ChevronDownIcon size={16} className={`tree-chevron-icon ${isExpanded ? '' : 'tree-chevron-icon--collapsed'}`} />
                  </button>
                  <div className="file-icon">
                    <FolderIcon size={20} />
                  </div>
                  <div className="file-info">
                    <span className="file-name">
                      {isRenaming ? (
                        <input
                          type="text"
                          className="module-inline-rename-input"
                          value={moduleRenameState?.draft || ''}
                          autoFocus
                          spellCheck={false}
                          maxLength={80}
                          onChange={(event) => handleRenameDraftChange(event.target.value)}
                          onClick={(event) => event.stopPropagation()}
                          onKeyDown={(event) => {
                            if (event.key === 'Enter') {
                              event.preventDefault();
                              void handleCommitRenameModule();
                            } else if (event.key === 'Escape') {
                              event.preventDefault();
                              handleCancelRenameModule();
                            }
                          }}
                          onBlur={() => {
                            void handleCommitRenameModule();
                          }}
                        />
                      ) : (
                        file.name
                      )}
                    </span>
                  </div>
                  <span className="file-type">{moduleWorkbookLabel}</span>
                </div>

                {isExpanded && (
                  <div className="allfiles-macro-list">
                    {moduleMacrosState.status === 'loading' && (
                      <div className="allfiles-macro-state">Loading macros...</div>
                    )}
                    {moduleMacrosState.status === 'error' && (
                      <div className="allfiles-macro-state error">
                        {String(moduleMacrosState.error || 'Unable to load macros.')}
                      </div>
                    )}
                    {moduleMacrosState.status === 'ready' && moduleMacrosState.macros.length < 1 && (
                      <div className="allfiles-macro-state">No macros found in this module.</div>
                    )}
                    {moduleMacrosState.status === 'ready' && moduleMacrosState.macros.map((macro) => (
                      <button
                        key={macro.id}
                        type="button"
                        className="allfiles-macro-item"
                        onClick={(event) => {
                          event.stopPropagation();
                          onRunMacro?.(macro);
                        }}
                      >
                        <span className="allfiles-macro-icon">
                          <ReturnIcon size={14} />
                        </span>
                        <span className="allfiles-macro-name">{macro.name}</span>
                      </button>
                    ))}
                  </div>
                )}
              </React.Fragment>
            );
          })}
        </div>
      </section>

      {moduleContextMenu && (
        <div
          ref={moduleContextMenuRef}
          className="module-context-menu"
          style={{ left: `${moduleContextMenu.x}px`, top: `${moduleContextMenu.y}px` }}
        >
          <button
            type="button"
            className="module-context-menu-item"
            onClick={() => handleStartRenameModule(moduleContextMenu.module)}
            disabled={moduleActionInFlight}
          >
            Rename
          </button>
          <button
            type="button"
            className="module-context-menu-item danger"
            onClick={() => handleRequestDeleteModule(moduleContextMenu.module)}
            disabled={moduleActionInFlight}
          >
            Delete
          </button>
        </div>
      )}

      {moduleDeleteTarget && (
        <div className="module-action-overlay" role="dialog" aria-modal="true" aria-label="Delete module confirmation">
          <div className="module-action-dialog">
            <h3 className="module-action-title">Delete module?</h3>
            <p className="module-action-message">
              {`Delete "${moduleDeleteTarget.name}" from "${moduleDeleteTarget.workbookName || selectedWorkbook?.name || 'workbook'}"?`}
            </p>
            <div className="module-action-buttons">
              <button
                type="button"
                className="module-action-btn"
                onClick={handleCancelDeleteModule}
                disabled={moduleActionInFlight}
              >
                Cancel
              </button>
              <button
                type="button"
                className="module-action-btn danger"
                onClick={() => { void handleConfirmDeleteModule(); }}
                disabled={moduleActionInFlight}
              >
                Delete
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );

  return (
    <>
      {/* Header / Search Bar */}
      <header className="header">
        <div className="drag-region" />
        <div
          className="search-input-wrapper"
          onMouseDown={handleSearchBarMouseDown}
        >
          <SearchIcon size={16} className="search-input-icon" />
          <input
            ref={searchInputRef}
            type="text"
            className="search-input"
            placeholder="Search files and manage macros"
            value={searchQuery}
            onChange={(event) => onSearchChange(event.target.value)}
            autoFocus
          />
        </div>

        <div className="header-actions">
          <button
            className="build-mode-btn"
            onClick={() => onBuildModeClick?.(selectedWorkbook || selectedWorkbookForBuild || null)}
          >
            New macro
          </button>
          <button className="close-btn" onClick={onClose}>
            <CloseIcon />
          </button>
        </div>
      </header>

      {/* Main Content */}
      <main className="main-content">
        {status === 'ready' ? renderReadyState() : renderNonReadyState()}
      </main>
    </>
  );
};

export default SearchMode;
