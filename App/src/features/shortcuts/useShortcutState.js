import { useState, useRef, useCallback, useEffect } from 'react';
import {
  mapAuditShortcutsToMacroIds
} from '../../lib/shortcut-audit';
import {
  normalizeShortcutLetterDraft,
  parseShortcutLetter,
  toExcelShortcutKeyFromLetter
} from '../../lib/shortcut-keybind';
import { SHORTCUT_REFRESH_TTL_MS } from '../search/search-constants';

export function useShortcutState({ searchData, setActionStatus }) {
  const [shortcutByMacroId, setShortcutByMacroId] = useState({});
  const [shortcutDraftByMacroId, setShortcutDraftByMacroId] = useState({});
  const [shortcutSavingMacroId, setShortcutSavingMacroId] = useState(null);

  const shortcutAuditRequestSequence = useRef(0);
  const shortcutAuditInFlight = useRef(false);
  const shortcutSnapshotRef = useRef('');
  const shortcutSnapshotTimestampRef = useRef(0);
  const shortcutByMacroIdRef = useRef({});
  const shortcutDraftByMacroIdRef = useRef({});
  const shortcutSavingMacroIdRef = useRef(null);
  const shortcutLoadErrorRef = useRef('');

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
          setActionStatus('error', message);
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
            setActionStatus('error', message);
            shortcutLoadErrorRef.current = message;
          }
        }
        return;
      }

      const rawShortcutMap = mapAuditShortcutsToMacroIds(result, macros);
      const shortcutMap = {};
      Object.entries(rawShortcutMap).forEach(([macroId, value]) => {
        const parsedLetter = parseShortcutLetter(value);
        if (parsedLetter) {
          shortcutMap[macroId] = parsedLetter;
        }
      });
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
          setActionStatus('error', message);
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
    setActionStatus
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

  const handleShortcutDraftChange = useCallback((macroId, value) => {
    const normalizedLetter = normalizeShortcutLetterDraft(value);
    setShortcutDraftByMacroId((previous) => ({
      ...previous,
      [macroId]: normalizedLetter
    }));
  }, []);

  const handleShortcutCommit = useCallback(async (macro, _trigger) => {
    if (!macro || shortcutSavingMacroIdRef.current) {
      return;
    }

    const savedMap = shortcutByMacroIdRef.current;
    const draftMap = shortcutDraftByMacroIdRef.current;
    const savedShortcut = savedMap[macro.id] || '';
    const draftShortcut = Object.prototype.hasOwnProperty.call(draftMap, macro.id)
      ? draftMap[macro.id]
      : savedShortcut;
    const normalizedShortcut = normalizeShortcutLetterDraft(draftShortcut);

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
      setActionStatus('error', 'Shortcut assign failed: Macro identity is missing.');
      return;
    }

    const setShortcutApi = window.excel?.vba?.setShortcut;
    if (!setShortcutApi) {
      setActionStatus('error', 'Shortcut assign failed: Excel VBA setShortcut API is unavailable.');
      return;
    }

    shortcutSavingMacroIdRef.current = macro.id;
    setShortcutSavingMacroId(macro.id);
    setActionStatus('running', `Saving shortcut for ${macro.name}...`);

    try {
      const excelShortcutKey = toExcelShortcutKeyFromLetter(normalizedShortcut);
      if (!excelShortcutKey) {
        setActionStatus('error', 'Shortcut assign failed: Enter a valid shortcut key.');
        return;
      }

      const result = await setShortcutApi({
        macroName,
        shortcutKey: excelShortcutKey
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
        setActionStatus('success', `Shortcut assigned: ${backendMessage}`);
        await loadMacroShortcuts({ force: true });
      } else {
        const backendMessage = result?.message || 'Unknown error.';
        setActionStatus('error', `Shortcut assign failed: ${backendMessage}`);
      }
    } catch (error) {
      const backendMessage = error?.message ? String(error.message) : 'Unexpected error.';
      setActionStatus('error', `Shortcut assign failed: ${backendMessage}`);
    } finally {
      shortcutSavingMacroIdRef.current = null;
      setShortcutSavingMacroId(null);
    }
  }, [loadMacroShortcuts, setActionStatus]);

  return {
    shortcutByMacroId,
    shortcutDraftByMacroId,
    shortcutSavingMacroId,
    handleShortcutDraftChange,
    handleShortcutCommit
  };
}
