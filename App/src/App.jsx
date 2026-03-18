import React, { useState, useEffect, useCallback, useRef } from 'react';
import './App.css';
import lightningIcon from '../assets/lightning.png';

// Import pages
import ShortcutsPage from './pages/ShortcutsPage';
import CreatePage from './pages/CreatePage';
import FilesPage from './pages/FilesPage';
import SettingsMenu from './components/SettingsMenu';
import { SettingsIcon, SidebarIcon, MinimizeIcon, CloseIcon } from './components/icons';
import { useSearchData } from './features/search/useSearchData';
import { useMacroRun } from './features/run/useMacroRun';
import { shouldClearShortcutState, useShortcutState } from './features/shortcuts/useShortcutState';
import { normalizeBuildWorkbook, resolveBuildLaunchMode } from './features/build/build-target';

const TAB_ITEMS = [
  { key: 'shortcuts', label: 'Shortcuts' },
  { key: 'create', label: 'Create' },
  { key: 'files', label: 'Files' }
];

/**
 * Main App Component
 *
 * Modes:
 * - 'shortcuts': Macro shortcuts grid with workbook picker
 * - 'create': AI Build mode for generating macros (includes inline code editor)
 * - 'files': File explorer with details panel
 */
function App() {
  // Current view mode
  const [mode, setMode] = useState('shortcuts');

  // Settings menu open state
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [actionState, setActionState] = useState('idle');
  const [actionMessage, setActionMessage] = useState('');
  const [selectedWorkbookForBuild, setSelectedWorkbookForBuild] = useState(null);
  const [buildLaunchContext, setBuildLaunchContext] = useState(() => ({
    mode: 'new_module',
    moduleName: '',
    source: 'toolbar',
    originMode: 'shortcuts'
  }));
  const [buildChatOpen, setBuildChatOpen] = useState(true);
  const [filesSidebarOpen, setFilesSidebarOpen] = useState(true);
  const loadSearchDataRef = useRef(null);
  const shortcutSaveInFlightRef = useRef(false);

  useEffect(() => {
    const preloadKey = 'lightning-icon';
    let preloadLink = document.querySelector(`link[data-preload-key="${preloadKey}"]`);
    if (!preloadLink) {
      preloadLink = document.createElement('link');
      preloadLink.rel = 'preload';
      preloadLink.as = 'image';
      preloadLink.href = lightningIcon;
      preloadLink.setAttribute('data-preload-key', preloadKey);
      document.head.appendChild(preloadLink);
    }

    const image = new Image();
    image.decoding = 'async';
    image.src = lightningIcon;

    return () => {
      image.src = '';
    };
  }, []);

  const setActionStatus = useCallback((status, message) => {
    setActionState(status);
    setActionMessage(message);
  }, []);

  useEffect(() => {
    if ((actionState !== 'success' && actionState !== 'error') || !actionMessage) {
      return undefined;
    }

    const timeoutId = window.setTimeout(() => {
      setActionState('idle');
      setActionMessage('');
    }, 4000);

    return () => window.clearTimeout(timeoutId);
  }, [actionState, actionMessage]);

  const {
    selectedMacro,
    setSelectedMacro,
    runState,
    handleRunMacro,
    macroRunInFlightRef
  } = useMacroRun({
    loadSearchDataRef,
    setActionStatus
  });

  const { searchData, loadSearchData } = useSearchData({
    mode: mode === 'shortcuts' ? 'search' : mode === 'create' ? 'build' : 'explorer',
    runState,
    macroRunInFlightRef,
    shortcutSaveInFlightRef
  });

  useEffect(() => {
    loadSearchDataRef.current = loadSearchData;
  }, [loadSearchData]);

  useEffect(() => {
    if (searchData.status !== 'ready') {
      setSelectedMacro(null);
      return;
    }

    setSelectedMacro((previous) => {
      if (!previous) {
        return null;
      }
      const matched = searchData.macros.find((macro) => macro.id === previous.id);
      return matched || null;
    });
  }, [searchData.macros, searchData.status, setSelectedMacro]);

  const {
    shortcutByMacroId,
    shortcutDraftByMacroId,
    shortcutInputErrorByMacroId,
    shortcutSavingMacroId,
    handleShortcutDraftChange,
    handleShortcutCommit
  } = useShortcutState({
    scope: 'active',
    enabled: searchData.status === 'ready',
    clearOnDisabled: shouldClearShortcutState(searchData.status),
    workbook: searchData.workbook,
    macros: searchData.macros,
    seededAudit: searchData.shortcutAudit,
    setActionStatus,
    shortcutSaveInFlightRef
  });

  const modeRef = useRef(mode);
  const settingsOpenRef = useRef(settingsOpen);
  const selectedWorkbookForBuildRef = useRef(selectedWorkbookForBuild);
  const searchWorkbookRef = useRef(searchData?.workbook || null);
  const openBuildModeRef = useRef(() => {});
  modeRef.current = mode;
  settingsOpenRef.current = settingsOpen;
  selectedWorkbookForBuildRef.current = selectedWorkbookForBuild;
  searchWorkbookRef.current = searchData?.workbook || null;

  const handleSelectedWorkbookForBuildChange = useCallback((workbook) => {
    const normalized = normalizeBuildWorkbook(workbook);
    setSelectedWorkbookForBuild((previous) => {
      const previousKey = String(previous?.key || '').trim();
      const nextKey = String(normalized?.key || '').trim();
      if (previousKey === nextKey) {
        return previous;
      }
      return normalized;
    });
  }, []);

  useEffect(() => {
    if (selectedWorkbookForBuildRef.current?.key) {
      return;
    }
    const fallbackWorkbook = normalizeBuildWorkbook(searchData?.workbook);
    if (!fallbackWorkbook) {
      return;
    }
    setSelectedWorkbookForBuild(fallbackWorkbook);
  }, [searchData?.workbook]);

  useEffect(() => {
    const setSelectedWorkbookApi = window.excel?.security?.setSelectedWorkbook;
    if (typeof setSelectedWorkbookApi !== 'function') {
      return;
    }

    const preferredWorkbook = normalizeBuildWorkbook(selectedWorkbookForBuild || searchData?.workbook);
    const workbookName = String(preferredWorkbook?.name || '').trim();
    const workbookPath = String(preferredWorkbook?.path || '').trim();

    void setSelectedWorkbookApi({ workbookName, workbookPath });
  }, [searchData?.workbook?.name, searchData?.workbook?.path, selectedWorkbookForBuild?.key, selectedWorkbookForBuild?.name, selectedWorkbookForBuild?.path]);

  const openBuildMode = useCallback((workbook = null, launchOptions = {}) => {
    const normalizedWorkbook = normalizeBuildWorkbook(
      workbook || selectedWorkbookForBuildRef.current || searchWorkbookRef.current
    );
    const launchMode = resolveBuildLaunchMode(launchOptions?.mode);
    const launchModuleName = String(launchOptions?.moduleName || '').trim();
    const launchSource = String(launchOptions?.source || '').trim() || 'toolbar';
    const launchOriginMode = String(launchOptions?.originMode || '').trim() === 'files'
      ? 'files'
      : 'shortcuts';

    if (normalizedWorkbook) {
      setSelectedWorkbookForBuild(normalizedWorkbook);
    }
    setBuildLaunchContext({
      mode: launchMode,
      moduleName: launchMode === 'existing_module' ? launchModuleName : '',
      source: launchSource,
      originMode: launchOriginMode
    });
    setMode('create');
  }, []);

  openBuildModeRef.current = openBuildMode;

  // Handle keyboard shortcuts — uses refs so the listener is registered once
  useEffect(() => {
    const handleKeyDown = (e) => {
      // Tab opens Create mode from Shortcuts mode only.
      if (e.key === 'Tab' && !e.ctrlKey && !e.shiftKey && !e.altKey) {
        if (modeRef.current !== 'shortcuts') {
          return;
        }

        const target = e.target;
        const tagName = String(target?.tagName || '').toUpperCase();
        const isEditableTarget = Boolean(
          target?.isContentEditable
          || tagName === 'INPUT'
          || tagName === 'TEXTAREA'
          || tagName === 'SELECT'
        );

        if (!isEditableTarget) {
          e.preventDefault();
          openBuildModeRef.current(null, { mode: 'new_module', source: 'hotkey' });
        }
      }

      // Escape to close settings
      if (e.key === 'Escape') {
        if (settingsOpenRef.current) {
          setSettingsOpen(false);
        }
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, []);

  // Stable callback references for child components
  const handleClose = useCallback(() => {
    if (window.excel?.app?.close) {
      window.excel.app.close();
    }
  }, []);

  const handleMinimize = useCallback(() => {
    if (window.excel?.app?.minimize) {
      window.excel.app.minimize();
    }
  }, []);

  const handleQuit = useCallback(() => {
    if (window.excel?.app?.close) {
      window.excel.app.close();
    }
  }, []);

  const goToCreate = useCallback((workbook = null) => {
    openBuildMode(workbook, { mode: 'new_module', source: 'toolbar', originMode: 'shortcuts' });
  }, [openBuildMode]);

  const handleCreateBack = useCallback(() => {
    if (buildLaunchContext.originMode === 'files') {
      setMode('files');
      return;
    }
    setMode('shortcuts');
  }, [buildLaunchContext.originMode]);

  const toggleSettings = useCallback(() => setSettingsOpen((prev) => !prev), []);
  const closeSettings = useCallback(() => setSettingsOpen(false), []);

  const handleTabChange = useCallback((newMode) => {
    if (newMode === 'create') {
      openBuildMode(null, { mode: 'new_module', source: 'tab', originMode: 'shortcuts' });
    } else {
      setMode(newMode);
    }
  }, [openBuildMode]);

  const toggleBuildChat = useCallback(() => {
    setBuildChatOpen((prev) => !prev);
  }, []);

  const toggleFilesSidebar = useCallback(() => {
    setFilesSidebarOpen((prev) => !prev);
  }, []);

  // Render current mode content
  const renderContent = () => {
    switch (mode) {
      case 'shortcuts':
        return (
          <ShortcutsPage
            onBuildModeClick={goToCreate}
            onRunMacro={handleRunMacro}
            searchData={searchData}
            selectedMacroId={selectedMacro?.id || null}
            shortcutByMacroId={shortcutByMacroId}
            shortcutDraftByMacroId={shortcutDraftByMacroId}
            shortcutInputErrorByMacroId={shortcutInputErrorByMacroId}
            shortcutSavingMacroId={shortcutSavingMacroId}
            onShortcutDraftChange={handleShortcutDraftChange}
            onShortcutCommit={handleShortcutCommit}
            onActionStatus={setActionStatus}
            shortcutSaveInFlightRef={shortcutSaveInFlightRef}
            selectedWorkbookForBuild={selectedWorkbookForBuild}
            onSelectedWorkbookForBuildChange={handleSelectedWorkbookForBuildChange}
          />
        );

      case 'create':
        return (
          <CreatePage
            onBack={handleCreateBack}
            onClose={handleClose}
            targetWorkbook={selectedWorkbookForBuild}
            launchMode={buildLaunchContext.mode}
            launchModuleName={buildLaunchContext.moduleName}
            launchSource={buildLaunchContext.source}
            onRefreshSearchData={loadSearchData}
            chatOpen={buildChatOpen}
            onChatToggle={toggleBuildChat}
          />
        );

      case 'files':
        return (
          <FilesPage
            searchData={searchData}
            onActionStatus={setActionStatus}
            onRefreshSearchData={loadSearchData}
            sidebarOpen={filesSidebarOpen}
            onEditModule={openBuildMode}
          />
        );

      default:
        return null;
    }
  };

  const showBottomActionBanner = (
    (actionState === 'running' || actionState === 'success' || actionState === 'error') &&
    Boolean(actionMessage)
  );

  return (
    <div className="app-container">
      {/* Shared Header with Tab Navigator */}
      <header className="header">
        <div className="drag-region" />
        <div className="header-left">
          {(mode === 'create' || mode === 'files') && (
            <button
              className="sidebar-toggle-btn"
              onClick={mode === 'create' ? toggleBuildChat : toggleFilesSidebar}
              title="Toggle sidebar"
            >
              <SidebarIcon size={19} />
            </button>
          )}
        </div>
        <div className="header-center">
          <nav className="tab-navigator">
            {TAB_ITEMS.map((tab) => (
              <button
                key={tab.key}
                className={`tab-navigator-btn${mode === tab.key ? ' active' : ''}`}
                onClick={() => handleTabChange(tab.key)}
              >
                {tab.label}
              </button>
            ))}
          </nav>
        </div>
        <div className="header-right">
          <button
            className="window-control-btn"
            onClick={handleMinimize}
            title="Minimize"
          >
            <MinimizeIcon size={18} />
          </button>
          <button
            className="window-control-btn close"
            onClick={handleClose}
            title="Close"
          >
            <CloseIcon size={18} />
          </button>
        </div>
      </header>

      {/* Main Content */}
      {renderContent()}

      {showBottomActionBanner && (
        <div className={`app-bottom-status app-bottom-status-${actionState}`}>
          {actionState === 'running' && <span className="status-spinner" />}
          <span className="app-bottom-status-label">
            {actionState === 'running'
              ? 'Processing'
              : actionState === 'success'
                ? 'Success'
                : 'Action failed'}
          </span>
          <span className="app-bottom-status-message">{actionMessage}</span>
        </div>
      )}

      {/* Settings button (bottom-left) */}
      {mode !== 'create' && !(mode === 'files' && !filesSidebarOpen) && (
        <div className="settings-floating" onClick={toggleSettings}>
          <SettingsIcon size={16} />
        </div>
      )}

      {/* Settings Menu */}
      <SettingsMenu
        isOpen={settingsOpen}
        onClose={closeSettings}
        onQuit={handleQuit}
      />
    </div>
  );
}

export default App;
