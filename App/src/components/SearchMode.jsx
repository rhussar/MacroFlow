import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { FolderIcon, ReturnIcon, CloseIcon, ChevronDownIcon, WorkbookIcon } from './icons';
import {
  getSearchStatusView,
  selectAllFilesModules,
  selectActiveWorkbookMacros,
  selectPersonalGlobalMacros,
  selectPersonalGlobalSectionModel
} from '../features/search/search-selectors';
import { formatShortcutPrefix } from '../lib/shortcut-keybind';
import { usePersonalMacros } from '../features/search/usePersonalMacros';
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
  const workbookMenuRef = useRef(null);
  const moduleContextMenuRef = useRef(null);
  const renameCommitInFlightRef = useRef(false);

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
  const activeMacroRows = useMemo(
    () => selectActiveWorkbookMacros(displayedWorkbookData.macros, searchQuery, effectiveShortcutByMacroId),
    [displayedWorkbookData.macros, effectiveShortcutByMacroId, searchQuery]
  );

  const personalMacrosState = usePersonalMacros(searchData, workbookPickerState.workbookListSignature);
  const personalMacroRows = useMemo(
    () => selectPersonalGlobalMacros(personalMacrosState.macros, searchQuery),
    [personalMacrosState.macros, searchQuery]
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
    if (status !== 'ready') {
      setWorkbookMenuOpen(false);
      setModuleContextMenu(null);
      setModuleRenameState(null);
      setModuleDeleteTarget(null);
    }
  }, [status]);

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

  const macrosEmptyMessage = workbookDataIsLoading
    ? `Loading macros from ${selectedWorkbookLabel}...`
    : workbookDataHasError
      ? selectedWorkbookErrorMessage
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
          <button
            type="button"
            className={`macro-workbook-picker ${isWorkbookMenuOpen ? 'open' : ''}`}
            aria-label="Select workbook"
            aria-expanded={isWorkbookMenuOpen}
            onClick={toggleWorkbookMenu}
          >
            <span className="macro-workbook-picker-label">{selectedWorkbookLabel}</span>
            <ChevronDownIcon size={14} className="macro-workbook-picker-icon" />
          </button>

          {isWorkbookMenuOpen && (
            <div className="macro-workbook-menu" role="listbox" aria-label="Open workbooks">
              <div className="macro-workbook-menu-title">Select open workbook</div>

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
                <button
                  type="button"
                  className="shortcut-run-target"
                  onClick={() => onRunMacro?.(macro)}
                >
                  <span className="shortcut-icon">
                    <ReturnIcon size={16} />
                  </span>
                  <span className="shortcut-name">{macro.name}</span>
                </button>
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
          <div className="section-header">
            <span className="section-title">Global Macros (PERSONAL.XLSB)</span>
            <span className="section-count">{personalSectionModel.count} items</span>
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
                return (
                  <div
                    key={row.uiId}
                    className="shortcut-item readonly"
                  >
                    <div className="shortcut-run-target readonly-target">
                      <span className="shortcut-icon">
                        <ReturnIcon size={16} />
                      </span>
                      <span className="shortcut-name">{macro.name}</span>
                    </div>
                    <div className="shortcut-binding">
                      <span className="global-shortcut-placeholder">Unavailable</span>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </section>
      )}

      <section className="search-ready-section">
        <div className="section-header">
          <span className="section-title">All Files</span>
          <span className="section-count">{filteredFiles.length} items</span>
        </div>

        <div className="file-list">
          {filteredFiles.length === 0 && (
            <div className="search-empty-state">{modulesEmptyMessage}</div>
          )}

          {status === 'ready' && filteredFiles.map((file) => {
            const isRenaming = moduleRenameState?.moduleId === file.id;

            return (
              <div
                key={file.id}
                className="file-item"
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
                <div className="file-icon">
                  <FolderIcon size={24} />
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
                    {file.workbookName && <span className="file-tag">{file.workbookName}</span>}
                  </span>
                </div>
                <span className="file-type">Macro Folder</span>
              </div>
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
        <div className="search-input-wrapper">
          <input
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
