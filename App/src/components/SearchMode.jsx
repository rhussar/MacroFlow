import React, { useEffect, useMemo, useRef, useState } from 'react';
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
  onClose
}) => {
  const status = searchData.status || 'idle';
  const workbookPickerState = useWorkbookPickerData(searchData, {
    preferredWorkbookKey: selectedWorkbookForBuild?.key || ''
  });
  const [isWorkbookMenuOpen, setWorkbookMenuOpen] = useState(false);
  const workbookMenuRef = useRef(null);

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
      status: personalMacrosState.status,
      workbookFound: personalMacrosState.workbookFound,
      error: personalMacrosState.error
    }),
    [personalMacroRows, personalMacrosState.error, personalMacrosState.status, personalMacrosState.workbookFound, selectedWorkbook?.name]
  );

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
    if (status !== 'ready') {
      setWorkbookMenuOpen(false);
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
              {personalSectionModel.emptyMessage}
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

          {status === 'ready' && filteredFiles.map((file) => (
            <div
              key={file.id}
              className="file-item"
              onClick={() => onFileClick?.({
                module: file,
                workbook: {
                  name: String(file?.workbookName || selectedWorkbook?.name || '').trim(),
                  path: String(file?.workbookPath || selectedWorkbook?.path || '').trim(),
                  key: String(file?.workbookPath || file?.workbookName || selectedWorkbook?.key || '').trim()
                },
                moduleId: String(file?.id || '').trim(),
                moduleName: String(file?.name || '').trim(),
                source: 'all-open-workbooks'
              })}
            >
              <div className="file-icon">
                <FolderIcon size={24} />
              </div>
              <div className="file-info">
                <span className="file-name">
                  {file.name}
                  {file.workbookName && <span className="file-tag">{file.workbookName}</span>}
                </span>
              </div>
              <span className="file-type">Macro Folder</span>
            </div>
          ))}
        </div>
      </section>
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
            AI Build Mode
            <span className="kbd">Tab</span>
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
