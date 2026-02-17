import React from 'react';
import { FolderIcon, ReturnIcon, CloseIcon } from './icons';
import {
  filterMacrosByQuery,
  filterModulesByQuery,
  getWorkbookContext
} from '../features/search/search-selectors';
import { formatShortcutPrefix } from '../lib/shortcut-keybind';

const defaultSearchData = {
  status: 'idle',
  workbook: null,
  modules: [],
  macros: [],
  error: null
};

const statusConfig = {
  idle: {
    title: 'Loading workbook data',
    message: 'Connecting to the active Excel workbook.'
  },
  loading: {
    title: 'Loading workbook data',
    message: 'Refreshing modules and macros from Excel.'
  },
  no_excel: {
    title: 'Excel is not running',
    message: 'Open Excel. MacroFlow will retry automatically.'
  },
  no_workbook: {
    title: 'No active workbook',
    message: 'Open or create a workbook. MacroFlow will retry automatically.'
  },
  error: {
    title: 'Could not load workbook data',
    message: 'Something went wrong while reading workbook data. Retrying automatically.'
  }
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
  shortcutSavingMacroId = null,
  onShortcutDraftChange,
  onShortcutCommit,
  actionState = 'idle',
  actionMessage = '',
  onClose
}) => {
  const status = searchData.status || 'idle';
  const workbookContext = getWorkbookContext(searchData.workbook);
  const allMacros = Array.isArray(searchData.macros) ? searchData.macros : [];

  const filteredFiles = filterModulesByQuery(searchData.modules, searchQuery);
  const filteredMacros = filterMacrosByQuery(allMacros, searchQuery, shortcutByMacroId);

  const renderNonReadyState = () => {
    const config = statusConfig[status] || statusConfig.error;
    const message = status === 'error' && searchData.error?.message
      ? searchData.error.message
      : config.message;
    const isLoadingState = status === 'idle' || status === 'loading';

    return (
      <div className={`search-status-panel search-status-${status}`}>
        <div className="search-status-header">
          {isLoadingState && <span className="status-spinner" />}
          <span className="search-status-title">{config.title}</span>
        </div>
        <p className="search-status-message">{message}</p>
      </div>
    );
  };

  const renderActionStatus = () => {
    if (actionState === 'idle' || !actionMessage) {
      return null;
    }

    const statusTitle = actionState === 'running'
      ? 'Processing'
      : actionState === 'success'
        ? 'Success'
        : 'Action failed';

    return (
      <div className={`search-run-status search-run-status-${actionState}`}>
        <div className="search-run-status-header">
          {actionState === 'running' && <span className="status-spinner" />}
          <span className="search-run-status-title">{statusTitle}</span>
        </div>
        <p className="search-run-status-message">{actionMessage}</p>
      </div>
    );
  };

  const renderReadyState = () => (
    <>
      <div className="search-context-bar">
        <span className="search-context-text" title={workbookContext.path}>
          {workbookContext.label}
        </span>
      </div>

      {renderActionStatus()}

      <div className="section-header">
        <span className="section-title">All Files</span>
        <span className="section-count">{filteredFiles.length} items</span>
      </div>

      <div className="file-list">
        {filteredFiles.length === 0 && (
          <div className="search-empty-state">No modules match this search.</div>
        )}

        {filteredFiles.map((file) => (
          <div
            key={file.id}
            className="file-item"
            onClick={() => onFileClick(file)}
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
            <span className="file-type">{`${file.type} - ${file.lineCount} lines`}</span>
          </div>
        ))}
      </div>

      <div className="section-header">
        <span className="section-title">Macros</span>
        <span className="section-count">{filteredMacros.length} items</span>
      </div>

      <div className="shortcuts-grid">
        {filteredMacros.length === 0 && (
          <div className="search-empty-state search-empty-state-grid">No macros match this search.</div>
        )}

        {filteredMacros.map((macro) => {
          const currentShortcutLetter = shortcutDraftByMacroId[macro.id] ?? shortcutByMacroId[macro.id] ?? '';
          const isSaving = shortcutSavingMacroId === macro.id;
          const shortcutPrefix = formatShortcutPrefix(currentShortcutLetter);

          return (
            <div
              key={macro.id}
              className={`shortcut-item ${selectedMacroId === macro.id ? 'selected' : ''} ${isSaving ? 'saving' : ''}`}
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
                  className={`shortcut-keycap-input ${currentShortcutLetter ? '' : 'is-empty'}`}
                  value={currentShortcutLetter}
                  placeholder=""
                  maxLength={1}
                  autoCapitalize="off"
                  autoComplete="off"
                  spellCheck={false}
                  aria-label={`Shortcut letter for ${macro.name}`}
                  onChange={(event) => onShortcutDraftChange?.(macro.id, event.target.value)}
                  onBlur={() => onShortcutCommit?.(macro, 'blur')}
                  onKeyDown={(event) => {
                    if (event.key === 'Enter') {
                      event.preventDefault();
                      event.currentTarget.blur();
                    } else if (event.key === 'Escape') {
                      event.preventDefault();
                      onShortcutDraftChange?.(macro.id, shortcutByMacroId[macro.id] || '');
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
            onChange={(e) => onSearchChange(e.target.value)}
            autoFocus
          />
        </div>

        <div className="header-actions">
          <button className="build-mode-btn" onClick={onBuildModeClick}>
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
