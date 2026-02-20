import { useState, useRef, useCallback, useEffect } from 'react';
import { mapAuditShortcutsToMacroIds } from '../../lib/shortcut-audit.js';
import {
  normalizeShortcutLetterDraft,
  parseShortcutLetter,
  toExcelShortcutKeyFromLetter
} from '../../lib/shortcut-keybind.js';
import { SHORTCUT_REFRESH_TTL_MS } from '../search/search-constants.js';

export function buildWorkbookShortcutSnapshotKey(workbook, macros = []) {
  const workbookKey = String(workbook?.path || workbook?.name || '').trim() || 'workbook';
  const macroIds = (Array.isArray(macros) ? macros : [])
    .map((macro) => String(macro?.id || ''))
    .filter(Boolean)
    .sort()
    .join('|');
  return `${workbookKey}::${macroIds}`;
}

export function useWorkbookShortcutState({
  enabled,
  workbook,
  macros,
  setActionStatus,
  shortcutSaveInFlightRef
}) {
  const [shortcutByMacroId, setShortcutByMacroId] = useState({});
  const [shortcutDraftByMacroId, setShortcutDraftByMacroId] = useState({});
  const [shortcutInputErrorByMacroId, setShortcutInputErrorByMacroId] = useState({});
  const [shortcutSavingMacroId, setShortcutSavingMacroId] = useState(null);

  const shortcutAuditRequestSequence = useRef(0);
  const shortcutAuditInFlight = useRef(false);
  const shortcutSnapshotRef = useRef('');
  const shortcutSnapshotTimestampRef = useRef(0);
  const shortcutByMacroIdRef = useRef({});
  const shortcutDraftByMacroIdRef = useRef({});
  const shortcutCacheBySnapshotRef = useRef(new Map());
  const shortcutSavingMacroIdRef = useRef(null);
  const shortcutLoadErrorRef = useRef('');

  const workbookName = String(workbook?.name || '').trim();
  const workbookPath = String(workbook?.path || '').trim();
  const macrosList = Array.isArray(macros) ? macros : [];

  const resetShortcutState = useCallback(() => {
    shortcutSnapshotRef.current = '';
    shortcutSnapshotTimestampRef.current = 0;
    shortcutLoadErrorRef.current = '';
    setShortcutByMacroId({});
    setShortcutDraftByMacroId({});
    setShortcutInputErrorByMacroId({});
    setShortcutSavingMacroId(null);
    shortcutByMacroIdRef.current = {};
    shortcutDraftByMacroIdRef.current = {};
    shortcutSavingMacroIdRef.current = null;
  }, []);

  const loadMacroShortcuts = useCallback(async ({ force = false } = {}) => {
    if (!enabled || !workbookName) {
      return;
    }

    const snapshotKey = buildWorkbookShortcutSnapshotKey(
      { name: workbookName, path: workbookPath },
      macrosList
    );
    const snapshotUnchanged = snapshotKey === shortcutSnapshotRef.current;
    const snapshotAgeMs = Date.now() - shortcutSnapshotTimestampRef.current;
    const snapshotStillFresh = snapshotAgeMs < SHORTCUT_REFRESH_TTL_MS;
    if (!force && snapshotUnchanged && snapshotStillFresh) {
      return;
    }

    const cachedSnapshot = shortcutCacheBySnapshotRef.current.get(snapshotKey);
    if (!force && cachedSnapshot && (Date.now() - Number(cachedSnapshot.timestamp || 0)) < SHORTCUT_REFRESH_TTL_MS) {
      shortcutSnapshotRef.current = snapshotKey;
      shortcutSnapshotTimestampRef.current = Number(cachedSnapshot.timestamp || Date.now());
      shortcutLoadErrorRef.current = '';
      setShortcutByMacroId(cachedSnapshot.shortcutMap || {});
      setShortcutDraftByMacroId(cachedSnapshot.draftMap || {});
      shortcutByMacroIdRef.current = cachedSnapshot.shortcutMap || {};
      shortcutDraftByMacroIdRef.current = cachedSnapshot.draftMap || {};
      return;
    }

    const auditApi = window.excel?.vba?.auditShortcutsByWorkbook;
    if (!auditApi) {
      if (force || Object.keys(shortcutByMacroIdRef.current).length === 0) {
        const message = 'Shortcut refresh failed: Workbook shortcut audit API is unavailable.';
        if (message !== shortcutLoadErrorRef.current) {
          setActionStatus?.('error', message);
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
      const result = await auditApi({ workbookName, workbookPath });
      if (requestId !== shortcutAuditRequestSequence.current) {
        return;
      }

      if (!result?.success) {
        const backendMessage = result?.message || 'Unknown error.';
        if (force || Object.keys(shortcutByMacroIdRef.current).length === 0) {
          const message = `Shortcut refresh failed: ${backendMessage}`;
          if (message !== shortcutLoadErrorRef.current) {
            setActionStatus?.('error', message);
            shortcutLoadErrorRef.current = message;
          }
        }
        return;
      }

      if (result?.workbookFound === false) {
        resetShortcutState();
        return;
      }

      const rawShortcutMap = mapAuditShortcutsToMacroIds(result, macrosList);
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
      macrosList.forEach((macro) => {
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
      shortcutCacheBySnapshotRef.current.set(snapshotKey, {
        shortcutMap,
        draftMap: nextDraftMap,
        timestamp: shortcutSnapshotTimestampRef.current
      });
      setShortcutByMacroId(shortcutMap);
      setShortcutDraftByMacroId(nextDraftMap);
      shortcutByMacroIdRef.current = shortcutMap;
      shortcutDraftByMacroIdRef.current = nextDraftMap;
    } catch (error) {
      if (force || Object.keys(shortcutByMacroIdRef.current).length === 0) {
        const backendMessage = error?.message ? String(error.message) : 'Unexpected error.';
        const message = `Shortcut refresh failed: ${backendMessage}`;
        if (message !== shortcutLoadErrorRef.current) {
          setActionStatus?.('error', message);
          shortcutLoadErrorRef.current = message;
        }
      }
    } finally {
      shortcutAuditInFlight.current = false;
    }
  }, [enabled, macrosList, resetShortcutState, setActionStatus, workbookName, workbookPath]);

  useEffect(() => {
    if (!enabled || !workbookName) {
      resetShortcutState();
      return;
    }
    loadMacroShortcuts();
  }, [enabled, loadMacroShortcuts, resetShortcutState, workbookName, macrosList]);

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
    setShortcutInputErrorByMacroId((previous) => {
      if (!previous[macroId]) {
        return previous;
      }
      const next = { ...previous };
      delete next[macroId];
      return next;
    });
  }, []);

  const handleShortcutCommit = useCallback(async (macro) => {
    if (!enabled || !workbookName || !macro || shortcutSavingMacroIdRef.current) {
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
      setShortcutInputErrorByMacroId((previous) => {
        if (!previous[macro.id]) {
          return previous;
        }
        const next = { ...previous };
        delete next[macro.id];
        return next;
      });
      return;
    }

    const duplicateMacroId = Object.keys(savedMap).find(
      (macroId) => macroId !== macro.id && savedMap[macroId] === normalizedShortcut
    );
    if (duplicateMacroId) {
      setShortcutInputErrorByMacroId((previous) => ({
        ...previous,
        [macro.id]: true
      }));
      return;
    }

    const macroName = macro?.fullName || macro?.runTarget || macro?.name || '';
    if (!macroName) {
      setActionStatus?.('error', 'Shortcut assign failed: Macro identity is missing.');
      return;
    }

    const setShortcutApi = window.excel?.vba?.setShortcutByWorkbook;
    if (!setShortcutApi) {
      setActionStatus?.('error', 'Shortcut assign failed: Workbook shortcut API is unavailable.');
      return;
    }

    shortcutSavingMacroIdRef.current = macro.id;
    if (shortcutSaveInFlightRef) {
      shortcutSaveInFlightRef.current = true;
    }
    setShortcutSavingMacroId(macro.id);
    setActionStatus?.('running', `Saving shortcut for ${macro.name}...`);

    try {
      const excelShortcutKey = toExcelShortcutKeyFromLetter(normalizedShortcut);
      if (!excelShortcutKey) {
        setActionStatus?.('error', 'Shortcut assign failed: Enter a valid shortcut key.');
        return;
      }

        const result = await setShortcutApi({
          workbookName,
          workbookPath,
          macroName,
          shortcutKey: excelShortcutKey
        });

      if (result?.workbookFound === false) {
        setActionStatus?.('error', result?.message || `Workbook "${workbookName}" is not open.`);
        return;
      }

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
        setShortcutInputErrorByMacroId((previous) => {
          if (!previous[macro.id]) {
            return previous;
          }
          const next = { ...previous };
          delete next[macro.id];
          return next;
        });
        setActionStatus?.('success', `Shortcut assigned: ${backendMessage}`);
        await loadMacroShortcuts({ force: true });
      } else {
        const backendMessage = result?.message || 'Unknown error.';
        setActionStatus?.('error', `Shortcut assign failed: ${backendMessage}`);
      }
    } catch (error) {
      const backendMessage = error?.message ? String(error.message) : 'Unexpected error.';
      setActionStatus?.('error', `Shortcut assign failed: ${backendMessage}`);
    } finally {
      shortcutSavingMacroIdRef.current = null;
      if (shortcutSaveInFlightRef) {
        shortcutSaveInFlightRef.current = false;
      }
      setShortcutSavingMacroId(null);
    }
  }, [enabled, loadMacroShortcuts, setActionStatus, shortcutSaveInFlightRef, workbookName, workbookPath]);

  return {
    shortcutByMacroId,
    shortcutDraftByMacroId,
    shortcutInputErrorByMacroId,
    shortcutSavingMacroId,
    handleShortcutDraftChange,
    handleShortcutCommit
  };
}
