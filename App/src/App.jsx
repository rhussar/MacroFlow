import React, { useState, useEffect, useCallback, useRef } from 'react';
import './App.css';

// Import components
import SearchMode from './components/SearchMode';
import BuildMode from './components/BuildMode';
import FileExplorer from './components/FileExplorer';
import SettingsMenu from './components/SettingsMenu';
import { MacroFlowLogo } from './components/icons';
import { useSearchData } from './features/search/useSearchData';
import { useMacroRun } from './features/run/useMacroRun';
import { useShortcutState } from './features/shortcuts/useShortcutState';
import { normalizeBuildWorkbook, resolveBuildLaunchMode } from './features/build/build-target';

/**
 * Main App Component
 *
 * Modes:
 * - 'search': Main search window with files and VBA shortcuts
 * - 'build': AI Build mode for generating macros (includes inline code editor)
 * - 'explorer': File explorer with details panel
 */
function App() {
  // Current view mode
  const [mode, setMode] = useState('search');

  // Search query (shared between search and explorer)
  const [searchQuery, setSearchQuery] = useState('');

  // Settings menu open state
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [actionState, setActionState] = useState('idle');
  const [actionMessage, setActionMessage] = useState('');
  const [explorerContext, setExplorerContext] = useState(null);
  const [selectedWorkbookForBuild, setSelectedWorkbookForBuild] = useState(null);
  const [buildLaunchContext, setBuildLaunchContext] = useState(() => ({
    mode: 'new_module',
    moduleName: '',
    source: 'toolbar'
  }));
  const loadSearchDataRef = useRef(null);
  const shortcutSaveInFlightRef = useRef(false);

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
    mode,
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
    searchData,
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

    if (normalizedWorkbook) {
      setSelectedWorkbookForBuild(normalizedWorkbook);
    }
    setBuildLaunchContext({
      mode: launchMode,
      moduleName: launchMode === 'existing_module' ? launchModuleName : '',
      source: launchSource
    });
    setMode('build');
  }, []);

  openBuildModeRef.current = openBuildMode;

  // Handle keyboard shortcuts — uses refs so the listener is registered once
  useEffect(() => {
    const handleKeyDown = (e) => {
      // Tab opens Build mode from Search mode only.
      // Build-mode exits are handled inside BuildMode so boundary-save cannot be bypassed.
      if (e.key === 'Tab' && !e.ctrlKey && !e.shiftKey && !e.altKey) {
        if (modeRef.current !== 'search') {
          return;
        }

        const target = e.target;
        const isSearchInput = target.classList?.contains('search-input');

        if (isSearchInput || target.tagName !== 'INPUT') {
          e.preventDefault();
          openBuildModeRef.current(null, { mode: 'new_module', source: 'hotkey' });
        }
      }

      // Escape to close or go back
      if (e.key === 'Escape') {
        if (settingsOpenRef.current) {
          setSettingsOpen(false);
        } else if (modeRef.current === 'explorer') {
          setExplorerContext(null);
          setMode('search');
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

  const handleQuit = useCallback(() => {
    if (window.excel?.app?.close) {
      window.excel.app.close();
    }
  }, []);

  const handleFileClick = useCallback((fileContext) => {
    const workbook = fileContext?.workbook || null;
    const workbookName = String(workbook?.name || '').trim();
    const workbookPath = String(workbook?.path || '').trim();
    const workbookKey = String(workbook?.key || workbookPath || workbookName).trim();
    const moduleId = String(fileContext?.moduleId || fileContext?.module?.id || '').trim();
    const moduleName = String(fileContext?.moduleName || fileContext?.module?.name || '').trim();
    const source = String(fileContext?.source || '').trim();

    if (source === 'all-open-workbooks') {
      openBuildMode(workbook, {
        mode: 'existing_module',
        moduleName,
        source: 'all-files'
      });
      return;
    }

    setExplorerContext(
      workbookKey || moduleId || moduleName
        ? {
            workbook: {
              name: workbookName,
              path: workbookPath,
              key: workbookKey
            },
            initialModuleId: moduleId,
            initialModuleName: moduleName
          }
        : null
    );
    setMode('explorer');
  }, [openBuildMode]);

  const goToBuild = useCallback((workbook = null) => {
    openBuildMode(workbook, { mode: 'new_module', source: 'toolbar' });
  }, [openBuildMode]);
  const goToSearch = useCallback(() => {
    setExplorerContext(null);
    setMode('search');
  }, []);
  const toggleSettings = useCallback(() => setSettingsOpen((prev) => !prev), []);
  const closeSettings = useCallback(() => setSettingsOpen(false), []);
  const handleExplorerContextConsumed = useCallback(() => {}, []);

  // Render current mode content
  const renderContent = () => {
    switch (mode) {
      case 'search':
        return (
          <SearchMode
            searchQuery={searchQuery}
            onSearchChange={setSearchQuery}
            onBuildModeClick={goToBuild}
            onFileClick={handleFileClick}
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
            onRefreshSearchData={loadSearchData}
            onClose={handleClose}
          />
        );

      case 'build':
        return (
          <BuildMode
            onBack={goToSearch}
            onClose={handleClose}
            targetWorkbook={selectedWorkbookForBuild}
            launchMode={buildLaunchContext.mode}
            launchModuleName={buildLaunchContext.moduleName}
            launchSource={buildLaunchContext.source}
            onRefreshSearchData={loadSearchData}
          />
        );

      case 'explorer':
        return (
          <FileExplorer
            onBack={goToSearch}
            onClose={handleClose}
            searchData={searchData}
            shortcutByMacroId={shortcutByMacroId}
            explorerContext={explorerContext}
            onExplorerContextConsumed={handleExplorerContextConsumed}
            onActionStatus={setActionStatus}
            onRefreshSearchData={loadSearchData}
          />
        );

      default:
        return null;
    }
  };

  // Check if we should show the default footer
  const showDefaultFooter = mode === 'search' || mode === 'explorer';
  const showBottomActionBanner = (
    (mode === 'search' || mode === 'explorer') &&
    showDefaultFooter &&
    (actionState === 'running' || actionState === 'success' || actionState === 'error') &&
    Boolean(actionMessage)
  );

  return (
    <div className="app-container">
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

      {/* Default Footer (for search and explorer modes) */}
      {showDefaultFooter && (
        <footer className="footer">
          <div className="footer-left">
            <div
              className="logo"
              onClick={toggleSettings}
            >
              <MacroFlowLogo size={20} />
            </div>
          </div>
        </footer>
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
