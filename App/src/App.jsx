import React, { useState, useEffect, useCallback, useRef } from 'react';
import './App.css';

// Import components
import SearchMode from './components/SearchMode';
import BuildMode from './components/BuildMode';
import FileExplorer from './components/FileExplorer';
import ManualEditMode from './components/ManualEditMode';
import SettingsMenu from './components/SettingsMenu';
import { MacroFlowLogo } from './components/icons';
import { useSearchData } from './features/search/useSearchData';
import { useMacroRun } from './features/run/useMacroRun';
import { useShortcutState } from './features/shortcuts/useShortcutState';

/**
 * Main App Component
 *
 * Modes:
 * - 'search': Main search window with files and VBA shortcuts
 * - 'build': AI Build mode for generating macros
 * - 'explorer': File explorer with details panel
 * - 'edit': Manual code editor mode
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
  const loadSearchDataRef = useRef(null);

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
    macroRunInFlightRef
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
    shortcutSavingMacroId,
    handleShortcutDraftChange,
    handleShortcutCommit
  } = useShortcutState({
    searchData,
    setActionStatus
  });

  const modeRef = useRef(mode);
  const settingsOpenRef = useRef(settingsOpen);
  modeRef.current = mode;
  settingsOpenRef.current = settingsOpen;

  // Handle keyboard shortcuts — uses refs so the listener is registered once
  useEffect(() => {
    const handleKeyDown = (e) => {
      // Tab to toggle between Search and Build modes
      if (e.key === 'Tab' && !e.ctrlKey && !e.shiftKey && !e.altKey) {
        const target = e.target;
        const isSearchInput = target.classList?.contains('search-input');

        if (isSearchInput || target.tagName !== 'INPUT') {
          e.preventDefault();
          setMode((prev) => {
            if (prev === 'search') return 'build';
            if (prev === 'build') return 'search';
            return prev;
          });
        }
      }

      // Alt+M for new macro (go to build mode)
      if (e.altKey && e.key === 'm') {
        e.preventDefault();
        setMode('build');
      }

      // Escape to close or go back
      if (e.key === 'Escape') {
        if (settingsOpenRef.current) {
          setSettingsOpen(false);
        } else if (modeRef.current === 'explorer' || modeRef.current === 'edit') {
          setMode('search');
        } else if (modeRef.current === 'build') {
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
    } else {
      console.log('Close app');
    }
  }, []);

  const handleQuit = useCallback(() => {
    if (window.excel?.app?.close) {
      window.excel.app.close();
    }
  }, []);

  const handleFileClick = useCallback((file) => {
    console.log('File clicked:', file);
    setMode('explorer');
  }, []);

  const goToBuild = useCallback(() => setMode('build'), []);
  const goToSearch = useCallback(() => setMode('search'), []);
  const goToEdit = useCallback(() => setMode('edit'), []);
  const toggleSettings = useCallback(() => setSettingsOpen((prev) => !prev), []);
  const closeSettings = useCallback(() => setSettingsOpen(false), []);

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
            shortcutSavingMacroId={shortcutSavingMacroId}
            onShortcutDraftChange={handleShortcutDraftChange}
            onShortcutCommit={handleShortcutCommit}
            onClose={handleClose}
          />
        );

      case 'build':
        return (
          <BuildMode
            onBack={goToSearch}
            onClose={handleClose}
            onEditMode={goToEdit}
          />
        );

      case 'explorer':
        // MF-103 scope: Explorer shortcut management is deferred until Explorer is live-data backed.
        return (
          <FileExplorer
            onBack={goToSearch}
            onClose={handleClose}
          />
        );

      case 'edit':
        return (
          <ManualEditMode
            onBack={goToBuild}
            onClose={handleClose}
          />
        );

      default:
        return null;
    }
  };

  // Check if we should show the default footer
  const showDefaultFooter = mode === 'search' || mode === 'explorer';
  const showBottomActionBanner = (
    mode === 'search' &&
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
          <div className="footer-right">
            <span className="footer-action">
              New macro
              <span className="kbd">Alt</span>
              <span className="kbd">M</span>
            </span>
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
