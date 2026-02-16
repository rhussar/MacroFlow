import React, { useState, useEffect, useCallback, useRef } from 'react';
import './App.css';

// Import components
import SearchMode from './components/SearchMode';
import BuildMode from './components/BuildMode';
import FileExplorer from './components/FileExplorer';
import ManualEditMode from './components/ManualEditMode';
import SettingsMenu from './components/SettingsMenu';
import { MacroFlowLogo } from './components/icons';
import {
  mapSearchError,
  normalizeMacros,
  normalizeModules,
  normalizeWorkbook
} from './lib/search-data';
import {
  mapAuditShortcutsToMacroIds,
  normalizeShortcutKey
} from './lib/shortcut-audit';

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
  const SHORTCUT_REFRESH_TTL_MS = 5000;
  const SEARCH_FOCUS_REFRESH_COOLDOWN_MS = 2500;
  const SEARCH_FULL_REFRESH_STALE_MS = 12000;
  const initialSearchData = {
    status: 'idle',
    workbook: null,
    modules: [],
    macros: [],
    error: null
  };

  // Current view mode
  const [mode, setMode] = useState('search');

  // Search query (shared between search and explorer)
  const [searchQuery, setSearchQuery] = useState('');

  // Settings menu open state
  const [settingsOpen, setSettingsOpen] = useState(false);

  // Live search data from Excel
  const [searchData, setSearchData] = useState(initialSearchData);
  const [selectedMacro, setSelectedMacro] = useState(null);
  const [runState, setRunState] = useState('idle');
  const [actionState, setActionState] = useState('idle');
  const [actionMessage, setActionMessage] = useState('');
  const [shortcutByMacroId, setShortcutByMacroId] = useState({});
  const [shortcutDraftByMacroId, setShortcutDraftByMacroId] = useState({});
  const [shortcutSavingMacroId, setShortcutSavingMacroId] = useState(null);
  const searchRequestSequence = useRef(0);
  const searchLoadInFlight = useRef(false);
  const macroRunInFlight = useRef(false);
  const workbookPingInFlight = useRef(false);
  const lastFocusRefreshAttemptAt = useRef(0);
  const lastFullSearchRefreshAt = useRef(0);
  const lastWorkbookSignature = useRef('');
  const shortcutAuditRequestSequence = useRef(0);
  const shortcutAuditInFlight = useRef(false);
  const shortcutSnapshotRef = useRef('');
  const shortcutSnapshotTimestampRef = useRef(0);
  const shortcutByMacroIdRef = useRef({});
  const shortcutDraftByMacroIdRef = useRef({});
  const shortcutLoadErrorRef = useRef('');

  const loadSearchData = useCallback(async ({ silent = false } = {}) => {
    if (searchLoadInFlight.current) {
      return;
    }

    searchLoadInFlight.current = true;
    const requestId = ++searchRequestSequence.current;

    if (!silent) {
      setSearchData((prev) => ({
        ...prev,
        status: 'loading',
        error: null
      }));
    }

    try {
      const workbookApi = window.excel?.workbook?.info;
      const modulesApi = window.excel?.vba?.modules;
      const proceduresApi = window.excel?.vba?.procedures;

      if (!workbookApi || !modulesApi || !proceduresApi) {
        const mappedError = mapSearchError('NO_EXCEL: Excel bridge API is unavailable.');
        if (requestId !== searchRequestSequence.current) {
          return;
        }
        setSelectedMacro(null);
        lastWorkbookSignature.current = '';
        setSearchData({
          status: mappedError.status,
          workbook: null,
          modules: [],
          macros: [],
          error: {
            code: mappedError.code,
            message: mappedError.message
          }
        });
        return;
      }

      const [workbookResult, modulesResult, proceduresResult] = await Promise.all([
        workbookApi(),
        modulesApi(),
        proceduresApi()
      ]);

      if (requestId !== searchRequestSequence.current) {
        return;
      }

      const failedResults = [workbookResult, modulesResult, proceduresResult].filter(
        (result) => !result?.success
      );

      if (failedResults.length > 0) {
        const failureMessage = failedResults
          .map((result) => result?.message)
          .filter(Boolean)
          .join(' | ');
        const mappedError = mapSearchError(failureMessage);

        setSelectedMacro(null);
        lastWorkbookSignature.current = '';
        setSearchData({
          status: mappedError.status,
          workbook: null,
          modules: [],
          macros: [],
          error: {
            code: mappedError.code,
            message: mappedError.message
          }
        });
        return;
      }

      const fallbackWorkbook = modulesResult?.workbook || proceduresResult?.workbook || null;
      const workbook = normalizeWorkbook(workbookResult, fallbackWorkbook);
      const modules = normalizeModules(modulesResult?.modules, workbook);
      const macros = normalizeMacros(proceduresResult?.procedures);
      lastWorkbookSignature.current = `${workbook?.path || ''}::${workbook?.name || ''}`;
      lastFullSearchRefreshAt.current = Date.now();

      setSelectedMacro((previous) => {
        if (!previous) {
          return null;
        }
        const matched = macros.find((macro) => macro.id === previous.id);
        return matched || null;
      });

      setSearchData({
        status: 'ready',
        workbook,
        modules,
        macros,
        error: null
      });
    } catch (error) {
      if (requestId !== searchRequestSequence.current) {
        return;
      }
      const mappedError = mapSearchError(error?.message);
      setSelectedMacro(null);
      lastWorkbookSignature.current = '';
      setSearchData({
        status: mappedError.status,
        workbook: null,
        modules: [],
        macros: [],
        error: {
          code: mappedError.code,
          message: mappedError.message
        }
      });
    }
    finally {
      searchLoadInFlight.current = false;
    }
  }, []);

  const refreshSearchOnForeground = useCallback(async () => {
    if (mode !== 'search' || runState === 'running' || macroRunInFlight.current) {
      return;
    }

    const now = Date.now();
    if (now - lastFocusRefreshAttemptAt.current < SEARCH_FOCUS_REFRESH_COOLDOWN_MS) {
      return;
    }
    lastFocusRefreshAttemptAt.current = now;

    const workbookApi = window.excel?.workbook?.info;
    if (!workbookApi || workbookPingInFlight.current) {
      await loadSearchData({ silent: true });
      return;
    }

    workbookPingInFlight.current = true;
    try {
      const ping = await workbookApi();
      if (!ping?.success) {
        await loadSearchData({ silent: true });
        return;
      }

      const workbookSignature = `${String(ping?.path || '').trim()}::${String(ping?.name || '').trim()}`;
      const workbookChanged = workbookSignature !== lastWorkbookSignature.current;
      const stale = Date.now() - lastFullSearchRefreshAt.current > SEARCH_FULL_REFRESH_STALE_MS;
      const shouldRefreshFull = searchData.status !== 'ready' || workbookChanged || stale;

      if (shouldRefreshFull) {
        await loadSearchData({ silent: true });
      }
    } catch (error) {
      await loadSearchData({ silent: true });
    } finally {
      workbookPingInFlight.current = false;
    }
  }, [
    SEARCH_FOCUS_REFRESH_COOLDOWN_MS,
    SEARCH_FULL_REFRESH_STALE_MS,
    loadSearchData,
    mode,
    runState,
    searchData.status
  ]);

  useEffect(() => {
    if (mode !== 'search' || runState === 'running') {
      return undefined;
    }

    // Entering Search should always perform one full refresh.
    loadSearchData({ silent: true });

    const handleFocus = () => {
      refreshSearchOnForeground();
    };
    const handleVisibilityChange = () => {
      if (document.visibilityState === 'visible') {
        refreshSearchOnForeground();
      }
    };

    window.addEventListener('focus', handleFocus);
    document.addEventListener('visibilitychange', handleVisibilityChange);

    return () => {
      window.removeEventListener('focus', handleFocus);
      document.removeEventListener('visibilitychange', handleVisibilityChange);
    };
  }, [mode, loadSearchData, refreshSearchOnForeground, runState]);

  const loadMacroShortcuts = useCallback(async ({ force = false } = {}) => {
    if (searchData.status !== 'ready') {
      return;
    }

    const macros = Array.isArray(searchData.macros) ? searchData.macros : [];
    const workbookKey = searchData.workbook?.path || searchData.workbook?.name || 'active-workbook';
    const snapshotKey = `${workbookKey}::${macros.map((macro) => macro.id).sort().join('|')}`;
    const snapshotUnchanged = snapshotKey === shortcutSnapshotRef.current;
    const snapshotAgeMs = Date.now() - shortcutSnapshotTimestampRef.current;
    const snapshotStillFresh = snapshotAgeMs < SHORTCUT_REFRESH_TTL_MS;
    if (!force && snapshotUnchanged && snapshotStillFresh) {
      return;
    }

    const auditApi = window.excel?.vba?.auditShortcuts;
    if (!auditApi) {
      if (force || Object.keys(shortcutByMacroIdRef.current).length === 0) {
        const message = 'Shortcut refresh failed: Excel VBA audit API is unavailable.';
        if (message !== shortcutLoadErrorRef.current) {
          setActionState('error');
          setActionMessage(message);
          shortcutLoadErrorRef.current = message;
        }
      }
      return;
    }
    if (shortcutAuditInFlight.current) {
      return;
    }

    shortcutAuditInFlight.current = true;
    const requestId = ++shortcutAuditRequestSequence.current;

    try {
      const result = await auditApi();
      if (requestId !== shortcutAuditRequestSequence.current) {
        return;
      }
      if (!result?.success) {
        if (force || Object.keys(shortcutByMacroIdRef.current).length === 0) {
          const backendMessage = result?.message || 'Unknown error.';
          const message = `Shortcut refresh failed: ${backendMessage}`;
          if (message !== shortcutLoadErrorRef.current) {
            setActionState('error');
            setActionMessage(message);
            shortcutLoadErrorRef.current = message;
          }
        }
        return;
      }

      const shortcutMap = mapAuditShortcutsToMacroIds(result, macros);
      const previousSavedMap = shortcutByMacroIdRef.current;
      const previousDraftMap = shortcutDraftByMacroIdRef.current;
      const nextDraftMap = {};
      macros.forEach((macro) => {
        const savedShortcut = previousSavedMap[macro.id] || '';
        const fetchedShortcut = shortcutMap[macro.id] || '';
        const hasDraft = Object.prototype.hasOwnProperty.call(previousDraftMap, macro.id);
        const currentDraft = hasDraft ? previousDraftMap[macro.id] : fetchedShortcut;
        const isDirtyDraft = hasDraft && currentDraft !== savedShortcut;
        nextDraftMap[macro.id] = isDirtyDraft ? currentDraft : fetchedShortcut;
      });

      shortcutSnapshotRef.current = snapshotKey;
      shortcutSnapshotTimestampRef.current = Date.now();
      shortcutLoadErrorRef.current = '';
      setShortcutByMacroId(shortcutMap);
      setShortcutDraftByMacroId(nextDraftMap);
      shortcutByMacroIdRef.current = shortcutMap;
      shortcutDraftByMacroIdRef.current = nextDraftMap;
    } catch (error) {
      if (force || Object.keys(shortcutByMacroIdRef.current).length === 0) {
        const backendMessage = error?.message ? String(error.message) : 'Unexpected error.';
        const message = `Shortcut refresh failed: ${backendMessage}`;
        if (message !== shortcutLoadErrorRef.current) {
          setActionState('error');
          setActionMessage(message);
          shortcutLoadErrorRef.current = message;
        }
      }
    } finally {
      shortcutAuditInFlight.current = false;
    }
  }, [
    searchData.macros,
    searchData.status,
    searchData.workbook?.name,
    searchData.workbook?.path,
    SHORTCUT_REFRESH_TTL_MS
  ]);

  useEffect(() => {
    if (searchData.status !== 'ready') {
      shortcutSnapshotRef.current = '';
      shortcutSnapshotTimestampRef.current = 0;
      shortcutLoadErrorRef.current = '';
      setShortcutByMacroId({});
      setShortcutDraftByMacroId({});
      setShortcutSavingMacroId(null);
      shortcutByMacroIdRef.current = {};
      shortcutDraftByMacroIdRef.current = {};
      return;
    }

    loadMacroShortcuts();
  }, [
    loadMacroShortcuts,
    searchData.macros,
    searchData.status,
    searchData.workbook?.name,
    searchData.workbook?.path
  ]);

  useEffect(() => {
    shortcutByMacroIdRef.current = shortcutByMacroId;
  }, [shortcutByMacroId]);

  useEffect(() => {
    shortcutDraftByMacroIdRef.current = shortcutDraftByMacroId;
  }, [shortcutDraftByMacroId]);

  // Handle keyboard shortcuts
  useEffect(() => {
    const handleKeyDown = (e) => {
      // Tab to toggle between Search and Build modes
      if (e.key === 'Tab' && !e.ctrlKey && !e.shiftKey && !e.altKey) {
        // Only toggle if not in an input field or if in our specific inputs
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
        if (settingsOpen) {
          setSettingsOpen(false);
        } else if (mode === 'explorer' || mode === 'edit') {
          setMode('search');
        } else if (mode === 'build') {
          setMode('search');
        }
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [mode, settingsOpen]);

  // Handle close app
  const handleClose = () => {
    // In Electron, this would close the window
    if (window.excel?.app?.close) {
      window.excel.app.close();
    } else {
      console.log('Close app');
    }
  };

  // Handle quit from settings
  const handleQuit = () => {
    handleClose();
  };

  // Handle file click (open explorer)
  const handleFileClick = (file) => {
    console.log('File clicked:', file);
    setMode('explorer');
  };

  // Handle macro run from Search mode
  const handleRunMacro = useCallback(async (macro) => {
    if (!macro || macroRunInFlight.current || runState === 'running') {
      return;
    }

    setSelectedMacro(macro);
    const macroName = macro?.fullName || macro?.runTarget || macro?.name || '';
    if (!macroName) {
      setRunState('error');
      setActionState('error');
      setActionMessage('Run failed: Macro identity is missing.');
      return;
    }

    const runApi = window.excel?.vba?.run;
    if (!runApi) {
      setRunState('error');
      setActionState('error');
      setActionMessage('Run failed: Excel VBA run API is unavailable.');
      return;
    }

    macroRunInFlight.current = true;
    setRunState('running');
    setActionState('running');
    setActionMessage(`Running ${macro.name || macroName}...`);

    try {
      const result = await runApi({ macroName });
      if (result?.success) {
        const backendMessage = result?.message || `Executed "${macroName}"`;
        setRunState('success');
        setActionState('success');
        setActionMessage(`Run succeeded: ${backendMessage}`);
      } else {
        const backendMessage = result?.message || 'Unknown error.';
        setRunState('error');
        setActionState('error');
        setActionMessage(`Run failed: ${backendMessage}`);
      }
    } catch (error) {
      const backendMessage = error?.message ? String(error.message) : 'Unexpected error.';
      setRunState('error');
      setActionState('error');
      setActionMessage(`Run failed: ${backendMessage}`);
    } finally {
      macroRunInFlight.current = false;
      await loadSearchData({ silent: true });
    }
  }, [loadSearchData, runState]);

  const handleShortcutDraftChange = useCallback((macroId, value) => {
    setShortcutDraftByMacroId((previous) => ({
      ...previous,
      [macroId]: value
    }));
  }, []);

  const handleShortcutCommit = useCallback(async (macro, _trigger) => {
    if (!macro || shortcutSavingMacroId) {
      return;
    }

    const savedShortcut = shortcutByMacroId[macro.id] || '';
    const draftShortcut = Object.prototype.hasOwnProperty.call(shortcutDraftByMacroId, macro.id)
      ? shortcutDraftByMacroId[macro.id]
      : savedShortcut;
    const normalizedShortcut = normalizeShortcutKey(draftShortcut);

    if (!normalizedShortcut) {
      setShortcutDraftByMacroId((previous) => ({
        ...previous,
        [macro.id]: savedShortcut
      }));
      return;
    }

    if (normalizedShortcut === savedShortcut) {
      if (draftShortcut !== normalizedShortcut) {
        setShortcutDraftByMacroId((previous) => ({
          ...previous,
          [macro.id]: normalizedShortcut
        }));
      }
      return;
    }

    const macroName = macro?.fullName || macro?.runTarget || macro?.name || '';
    if (!macroName) {
      setActionState('error');
      setActionMessage('Shortcut assign failed: Macro identity is missing.');
      return;
    }

    const setShortcutApi = window.excel?.vba?.setShortcut;
    if (!setShortcutApi) {
      setActionState('error');
      setActionMessage('Shortcut assign failed: Excel VBA setShortcut API is unavailable.');
      return;
    }

    setShortcutSavingMacroId(macro.id);
    setActionState('running');
    setActionMessage(`Saving shortcut for ${macro.name}...`);

    try {
      const result = await setShortcutApi({
        macroName,
        shortcutKey: normalizedShortcut
      });

      if (result?.success) {
        const backendMessage = result?.message || `${macroName} -> ${normalizedShortcut}`;
        setShortcutByMacroId((previous) => ({
          ...previous,
          [macro.id]: normalizedShortcut
        }));
        setShortcutDraftByMacroId((previous) => ({
          ...previous,
          [macro.id]: normalizedShortcut
        }));
        setActionState('success');
        setActionMessage(`Shortcut assigned: ${backendMessage}`);
        await loadMacroShortcuts({ force: true });
      } else {
        const backendMessage = result?.message || 'Unknown error.';
        setActionState('error');
        setActionMessage(`Shortcut assign failed: ${backendMessage}`);
      }
    } catch (error) {
      const backendMessage = error?.message ? String(error.message) : 'Unexpected error.';
      setActionState('error');
      setActionMessage(`Shortcut assign failed: ${backendMessage}`);
    } finally {
      setShortcutSavingMacroId(null);
    }
  }, [loadMacroShortcuts, shortcutByMacroId, shortcutDraftByMacroId, shortcutSavingMacroId]);

  // Render current mode content
  const renderContent = () => {
    switch (mode) {
      case 'search':
        return (
          <SearchMode
            searchQuery={searchQuery}
            onSearchChange={setSearchQuery}
            onBuildModeClick={() => setMode('build')}
            onFileClick={handleFileClick}
            onRunMacro={handleRunMacro}
            searchData={searchData}
            selectedMacroId={selectedMacro?.id || null}
            shortcutByMacroId={shortcutByMacroId}
            shortcutDraftByMacroId={shortcutDraftByMacroId}
            shortcutSavingMacroId={shortcutSavingMacroId}
            onShortcutDraftChange={handleShortcutDraftChange}
            onShortcutCommit={handleShortcutCommit}
            actionState={actionState}
            actionMessage={actionMessage}
            onClose={handleClose}
          />
        );

      case 'build':
        return (
          <BuildMode
            onBack={() => setMode('search')}
            onClose={handleClose}
            onEditMode={() => setMode('edit')}
          />
        );

      case 'explorer':
        // MF-103 scope: Explorer shortcut management is deferred until Explorer is live-data backed.
        return (
          <FileExplorer
            onBack={() => setMode('search')}
            onClose={handleClose}
          />
        );

      case 'edit':
        return (
          <ManualEditMode
            onBack={() => setMode('build')}
            onClose={handleClose}
          />
        );

      default:
        return null;
    }
  };

  // Check if we should show the default footer
  const showDefaultFooter = mode === 'search' || mode === 'explorer';

  return (
    <div className="app-container">
      {/* Main Content */}
      {renderContent()}

      {/* Default Footer (for search and explorer modes) */}
      {showDefaultFooter && (
        <footer className="footer">
          <div className="footer-left">
            <div
              className="logo"
              onClick={() => setSettingsOpen(!settingsOpen)}
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
        onClose={() => setSettingsOpen(false)}
        onQuit={handleQuit}
      />
    </div>
  );
}

export default App;
