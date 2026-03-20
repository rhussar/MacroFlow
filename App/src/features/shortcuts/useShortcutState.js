import { useState, useRef, useCallback, useEffect, useMemo } from 'react';
import { mapAuditShortcutsToMacroIds } from '../../lib/shortcut-audit.js';
import {
  normalizeShortcutLetterDraft,
  parseShortcutLetter,
  toExcelShortcutKeyFromLetter
} from '../../lib/shortcut-keybind.js';
import { SHORTCUT_REFRESH_TTL_MS } from '../search/search-constants.js';
import {
  SEARCH_INVALIDATION_BUCKETS,
  getShortcutAuditInvalidationScope,
  invalidateSearchBuckets,
  useSearchInvalidationRevision
} from '../search/search-invalidation.js';

const sharedShortcutCacheBySnapshot = new Map();

function buildMacroIdSignature(macros = []) {
  return (Array.isArray(macros) ? macros : [])
    .map((macro) => String(macro?.id || ''))
    .filter(Boolean)
    .sort()
    .join('|');
}

function buildShortcutSnapshotKey({ scope, workbook, macros = [] }) {
  const workbookKey = String(workbook?.path || workbook?.name || '').trim()
    || (scope === 'active' ? 'active-workbook' : 'workbook');
  return `${workbookKey}::${buildMacroIdSignature(macros)}`;
}

function getShortcutApiContext(scope, workbookName, workbookPath) {
  if (scope === 'workbook') {
    return {
      auditApi: window.excel?.vba?.auditShortcutsByWorkbook,
      auditArgs: { workbookName, workbookPath },
      setShortcutApi: window.excel?.vba?.setShortcutByWorkbook,
      buildSetShortcutArgs: ({ macroName, shortcutKey }) => ({
        workbookName,
        workbookPath,
        macroName,
        shortcutKey
      }),
      unavailableAuditMessage: 'Shortcut refresh failed: Workbook shortcut audit API is unavailable.',
      unavailableSetMessage: 'Shortcut assign failed: Workbook shortcut API is unavailable.'
    };
  }

  return {
    auditApi: window.excel?.vba?.auditShortcuts,
    auditArgs: undefined,
    setShortcutApi: window.excel?.vba?.setShortcut,
    buildSetShortcutArgs: ({ macroName, shortcutKey }) => ({
      macroName,
      shortcutKey
    }),
    unavailableAuditMessage: 'Shortcut refresh failed: Excel VBA audit API is unavailable.',
    unavailableSetMessage: 'Shortcut assign failed: Excel VBA setShortcut API is unavailable.'
  };
}

export function shouldClearShortcutState(status) {
  return status === 'no_excel' || status === 'no_workbook' || status === 'multi_instance' || status === 'error';
}

export function buildWorkbookShortcutSnapshotKey(workbook, macros = []) {
  return buildShortcutSnapshotKey({
    scope: 'workbook',
    workbook,
    macros
  });
}

export function shouldDelayInitialShortcutLoad({
  delayMs = 0,
  hasLoadedShortcuts = false,
  hasSeededAudit = false,
  hasFreshCachedSnapshot = false
}) {
  return Number(delayMs) > 0
    && !hasLoadedShortcuts
    && !hasSeededAudit
    && !hasFreshCachedSnapshot;
}

export function useShortcutState({
  scope = 'active',
  enabled = false,
  clearOnDisabled = true,
  workbook = null,
  macros = [],
  seededAudit = null,
  initialLoadDelayMs = 0,
  setActionStatus,
  shortcutSaveInFlightRef
}) {
  const [shortcutByMacroId, setShortcutByMacroId] = useState({});
  const [shortcutDraftByMacroId, setShortcutDraftByMacroId] = useState({});
  const [shortcutInputErrorByMacroId, setShortcutInputErrorByMacroId] = useState({});
  const [shortcutSavingMacroId, setShortcutSavingMacroId] = useState(null);

  const shortcutAuditRequestSequence = useRef(0);
  const shortcutAuditInFlight = useRef(false);
  const lastHandledShortcutInvalidationRef = useRef(0);
  const shortcutSnapshotRef = useRef('');
  const shortcutSnapshotTimestampRef = useRef(0);
  const shortcutByMacroIdRef = useRef({});
  const shortcutDraftByMacroIdRef = useRef({});
  const shortcutSavingMacroIdRef = useRef(null);
  const shortcutLoadErrorRef = useRef('');
  const initialLoadTimerRef = useRef(null);

  const normalizedScope = scope === 'workbook' ? 'workbook' : 'active';
  const workbookName = String(workbook?.name || '').trim();
  const workbookPath = String(workbook?.path || '').trim();
  const shortcutInvalidationScope = getShortcutAuditInvalidationScope({
    scope: normalizedScope,
    workbook: { name: workbookName, path: workbookPath }
  });
  const shortcutInvalidationRevision = useSearchInvalidationRevision(
    SEARCH_INVALIDATION_BUCKETS.SHORTCUT_AUDIT,
    shortcutInvalidationScope
  );
  const macrosList = useMemo(
    () => (Array.isArray(macros) ? macros : []),
    [macros]
  );
  const macroIdSignature = useMemo(
    () => buildMacroIdSignature(macrosList),
    [macrosList]
  );

  const applyShortcutAuditResult = useCallback((auditResult, snapshotKey) => {
    const rawShortcutMap = mapAuditShortcutsToMacroIds(auditResult, macrosList);
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
    sharedShortcutCacheBySnapshot.set(snapshotKey, {
      shortcutMap,
      draftMap: nextDraftMap,
      timestamp: shortcutSnapshotTimestampRef.current
    });

    setShortcutByMacroId(shortcutMap);
    setShortcutDraftByMacroId(nextDraftMap);
    shortcutByMacroIdRef.current = shortcutMap;
    shortcutDraftByMacroIdRef.current = nextDraftMap;
  }, [macrosList]);

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
    const requiresWorkbookIdentity = normalizedScope === 'workbook';
    if (!enabled || (requiresWorkbookIdentity && !workbookName)) {
      return;
    }

    const snapshotKey = buildShortcutSnapshotKey({
      scope: normalizedScope,
      workbook: { name: workbookName, path: workbookPath },
      macros: macrosList
    });
    const snapshotUnchanged = snapshotKey === shortcutSnapshotRef.current;
    const snapshotAgeMs = Date.now() - shortcutSnapshotTimestampRef.current;
    const snapshotStillFresh = snapshotAgeMs < SHORTCUT_REFRESH_TTL_MS;
    if (!force && snapshotUnchanged && snapshotStillFresh) {
      return;
    }

    const cachedSnapshot = sharedShortcutCacheBySnapshot.get(snapshotKey);
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

    if (
      !force &&
      seededAudit &&
      typeof seededAudit === 'object' &&
      seededAudit.success !== false &&
      Array.isArray(seededAudit.shortcuts)
    ) {
      applyShortcutAuditResult(seededAudit, snapshotKey);
      return;
    }

    const {
      auditApi,
      auditArgs,
      unavailableAuditMessage
    } = getShortcutApiContext(normalizedScope, workbookName, workbookPath);

    if (typeof auditApi !== 'function') {
      if (force || Object.keys(shortcutByMacroIdRef.current).length === 0) {
        if (unavailableAuditMessage !== shortcutLoadErrorRef.current) {
          setActionStatus?.('error', unavailableAuditMessage);
          shortcutLoadErrorRef.current = unavailableAuditMessage;
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
      const result = auditArgs ? await auditApi(auditArgs) : await auditApi();
      if (requestId !== shortcutAuditRequestSequence.current) {
        return;
      }

      if (!result?.success) {
        if (force || Object.keys(shortcutByMacroIdRef.current).length === 0) {
          const backendMessage = result?.message || 'Unknown error.';
          const message = `Shortcut refresh failed: ${backendMessage}`;
          if (message !== shortcutLoadErrorRef.current) {
            setActionStatus?.('error', message);
            shortcutLoadErrorRef.current = message;
          }
        }
        return;
      }

      if (normalizedScope === 'workbook' && result?.workbookFound === false) {
        resetShortcutState();
        return;
      }

      applyShortcutAuditResult(result, snapshotKey);
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
  }, [
    applyShortcutAuditResult,
    enabled,
    macrosList,
    normalizedScope,
    resetShortcutState,
    seededAudit,
    setActionStatus,
    workbookName,
    workbookPath
  ]);

  useEffect(() => {
    const requiresWorkbookIdentity = normalizedScope === 'workbook';
    if (!enabled || (requiresWorkbookIdentity && !workbookName)) {
      if (clearOnDisabled) {
        resetShortcutState();
      }
      if (initialLoadTimerRef.current) {
        clearTimeout(initialLoadTimerRef.current);
        initialLoadTimerRef.current = null;
      }
      return;
    }

    const snapshotKey = buildShortcutSnapshotKey({
      scope: normalizedScope,
      workbook: { name: workbookName, path: workbookPath },
      macros: macrosList
    });
    const invalidationChanged =
      shortcutInvalidationRevision > 0
      && shortcutInvalidationRevision !== lastHandledShortcutInvalidationRef.current;
    if (invalidationChanged) {
      lastHandledShortcutInvalidationRef.current = shortcutInvalidationRevision;
      sharedShortcutCacheBySnapshot.delete(snapshotKey);
      shortcutSnapshotRef.current = '';
      shortcutSnapshotTimestampRef.current = 0;
    }
    const cachedSnapshot = sharedShortcutCacheBySnapshot.get(snapshotKey);
    const hasFreshCachedSnapshot = cachedSnapshot
      && (Date.now() - Number(cachedSnapshot.timestamp || 0)) < SHORTCUT_REFRESH_TTL_MS;
    const hasLoadedShortcuts = shortcutSnapshotRef.current === snapshotKey
      && Object.keys(shortcutByMacroIdRef.current).length > 0;
    const hasSeededAudit = Boolean(
      seededAudit &&
      typeof seededAudit === 'object' &&
      seededAudit.success !== false &&
      Array.isArray(seededAudit.shortcuts)
    );

    if (!invalidationChanged && shouldDelayInitialShortcutLoad({
      delayMs: initialLoadDelayMs,
      hasLoadedShortcuts,
      hasSeededAudit,
      hasFreshCachedSnapshot
    })) {
      if (initialLoadTimerRef.current) {
        clearTimeout(initialLoadTimerRef.current);
      }
      initialLoadTimerRef.current = window.setTimeout(() => {
        initialLoadTimerRef.current = null;
        void loadMacroShortcuts();
      }, Number(initialLoadDelayMs));
      return () => {
        if (initialLoadTimerRef.current) {
          clearTimeout(initialLoadTimerRef.current);
          initialLoadTimerRef.current = null;
        }
      };
    }

    loadMacroShortcuts({ force: invalidationChanged });
    return () => {
      if (initialLoadTimerRef.current) {
        clearTimeout(initialLoadTimerRef.current);
        initialLoadTimerRef.current = null;
      }
    };
  }, [
    clearOnDisabled,
    enabled,
    initialLoadDelayMs,
    loadMacroShortcuts,
    macroIdSignature,
    macrosList,
    normalizedScope,
    resetShortcutState,
    seededAudit,
    shortcutInvalidationRevision,
    workbookName,
    workbookPath
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
    const requiresWorkbookIdentity = normalizedScope === 'workbook';
    if (
      !enabled ||
      !macro ||
      shortcutSavingMacroIdRef.current ||
      (requiresWorkbookIdentity && !workbookName)
    ) {
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
      setActionStatus?.('error', 'Macro not found.');
      return;
    }

    const {
      setShortcutApi,
      buildSetShortcutArgs,
      unavailableSetMessage
    } = getShortcutApiContext(normalizedScope, workbookName, workbookPath);

    if (typeof setShortcutApi !== 'function') {
      setActionStatus?.('error', unavailableSetMessage);
      return;
    }

    shortcutSavingMacroIdRef.current = macro.id;
    if (shortcutSaveInFlightRef) {
      shortcutSaveInFlightRef.current = true;
    }
    setShortcutSavingMacroId(macro.id);
    setActionStatus?.('running', 'Saving shortcut...');

    try {
      const excelShortcutKey = toExcelShortcutKeyFromLetter(normalizedShortcut);
      if (!excelShortcutKey) {
        setActionStatus?.('error', 'Enter a valid key.');
        return;
      }

      const result = await setShortcutApi(buildSetShortcutArgs({
        macroName,
        shortcutKey: excelShortcutKey
      }));

      if (normalizedScope === 'workbook' && result?.workbookFound === false) {
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
        setActionStatus?.('success', 'Shortcut assigned.');
        invalidateSearchBuckets([
          {
            bucket: SEARCH_INVALIDATION_BUCKETS.SHORTCUT_AUDIT,
            scope: shortcutInvalidationScope
          },
          ...(normalizedScope === 'active'
            ? [{ bucket: SEARCH_INVALIDATION_BUCKETS.ACTIVE_WORKBOOK }]
            : [])
        ]);
      } else {
        const backendMessage = result?.message || 'Unknown error.';
        setActionStatus?.('error', 'Shortcut failed.');
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
  }, [
    enabled,
    normalizedScope,
    setActionStatus,
    shortcutInvalidationScope,
    shortcutSaveInFlightRef,
    workbookName,
    workbookPath
  ]);

  return {
    shortcutByMacroId,
    shortcutDraftByMacroId,
    shortcutInputErrorByMacroId,
    shortcutSavingMacroId,
    handleShortcutDraftChange,
    handleShortcutCommit
  };
}
