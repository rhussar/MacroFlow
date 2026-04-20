import React, { useState, useEffect, useRef, useCallback, useMemo } from 'react';
import CodePreview from '../components/CodePreview';
import SplitDivider from '../components/SplitDivider';
import {
  BUILD_MODE_SEED_CODE,
  normalizeBuildWorkbook,
  toWorkbookRequest,
  selectNextModuleName,
  extractPrimaryMacroName,
  buildWorkbookQualifiedRunTarget,
  buildSessionMacroTarget,
  findAssignedShortcutLetterForMacro,
  hasShortcutConflictForMacro,
  resolveBuildLaunchMode,
  shouldUseStrictWorkbook,
  resolveExistingModuleName,
  resolveBuildWorkbookTarget,
  shouldSyncOnBuildExit,
  resolveBuildExitAction
} from '../features/build/build-target';
import {
  formatShortcutPrefix,
  normalizeShortcutLetterDraft,
  toExcelShortcutKeyFromLetter
} from '../lib/shortcut-keybind';
import {
  isValidVbaModuleName,
  shouldCommitModuleRename,
  displayMacroName,
  encodeMacroName
} from '../features/search/module-actions';
import { getSearchStatusView } from '../features/search/search-selectors';
import {
  buildWorkbookInvalidationDescriptors,
  invalidateSearchBuckets
} from '../features/search/search-invalidation';
import { PERSONAL_WORKBOOK_NAME } from '../features/search/usePersonalMacros';
import { useSessionHistory } from '../features/build/useSessionHistory';
import { useModuleRename } from '../features/build/useModuleRename';
import { useMacroExecution } from '../features/build/useMacroExecution';
import { useCodeEditing } from '../features/build/useCodeEditing';
import ImageMsoIcon, { isSpriteReady } from '../components/ImageMsoIcon';
import IconPicker from '../components/IconPicker';
import { getMacroIcon, setMacroIcon, renameMacroIcon } from '../features/icons/macroIconStore';
import { ReturnIcon, WorkbookIcon } from '../components/icons';

const DEFAULT_MODULE_LABEL = 'New Module';
const MODULE_PREFIX = 'Module_';
const RESTART_REQUIRED_MESSAGE =
  'Restart required: Close and reopen MacroFlow to enable workbook-scoped VBA sync APIs.';

const INITIALIZATION_STEP_TEMPLATE = [
  { key: 'resolve', text: 'Resolving workbook', status: 'pending' },
  { key: 'create', text: 'Creating module', status: 'pending' }
];

function buildLocation(workbookName, moduleName = DEFAULT_MODULE_LABEL) {
  const workbook = String(workbookName || '').trim() || 'Active Workbook';
  const module = String(moduleName || '').trim() || DEFAULT_MODULE_LABEL;
  return { workbook, module };
}

function createInitializationSteps(overrides = {}) {
  return INITIALIZATION_STEP_TEMPLATE.map((step) => ({
    ...step,
    status: overrides[step.key] || step.status
  }));
}

function toActiveWorkbookModel(workbookInfoResult) {
  const workbook = {
    name: String(workbookInfoResult?.name || '').trim(),
    path: String(workbookInfoResult?.path || '').trim(),
    key: String(workbookInfoResult?.path || workbookInfoResult?.name || '').trim()
  };
  return normalizeBuildWorkbook(workbook);
}

function normalizeExitIntent(intent) {
  return String(intent || '').trim().toLowerCase() === 'close' ? 'close' : 'back';
}

function mapAiGenerationMessage(result, fallbackMessage = '') {
  const reason = String(result?.reason || '').trim();
  const backendMessage = String(result?.message || '').trim();
  if (backendMessage) {
    return backendMessage;
  }

  switch (reason) {
    case 'AI_RUNTIME_MISSING':
      return 'Local AI runtime is not installed yet.';
    case 'AI_MODEL_MISSING':
      return 'The local AI model is not installed yet.';
    case 'AI_NOT_READY':
      return 'Local AI is not ready yet. Finish setup and retry.';
    case 'AI_TIMEOUT':
      return 'Local AI timed out. Please retry.';
    case 'AI_INVALID_PROMPT':
      return 'Prompt is required to generate VBA.';
    case 'AI_INVALID_RESPONSE':
      return 'Local AI returned invalid VBA output.';
    default:
      return fallbackMessage || 'Unable to generate VBA.';
  }
}

function buildAssistantMessageFromResult(result, fallbackText = '') {
  const intent = String(result?.intent || '').trim().toLowerCase();
  const content = String(result?.content || '').trim();

  if (content) {
    return {
      role: 'assistant',
      kind: 'text',
      content,
      timestamp: Date.now()
    };
  }

  return {
    role: 'assistant',
    kind: 'text',
    content: fallbackText || 'Response received.',
    timestamp: Date.now()
  };
}

function inferBuildPromptIntent(promptText) {
  const normalized = String(promptText || '').trim().toLowerCase();
  if (!normalized) {
    return 'edit';
  }

  const questionPrefixes = [
    'what ',
    'what does',
    'what is',
    'why ',
    'how ',
    'can ',
    'could ',
    'should ',
    'does ',
    'is ',
    'are ',
    'explain ',
    'help ',
    'tell me ',
    'walk me through '
  ];

  if (normalized.endsWith('?')) {
    return 'ask';
  }

  if (questionPrefixes.some((prefix) => normalized.startsWith(prefix))) {
    return 'ask';
  }

  return 'edit';
}

const CreatePage = ({
  onBack,
  onClose,
  attemptExitRef: parentAttemptExitRef,
  targetWorkbook,
  launchMode = 'new_module',
  launchModuleName = '',
  launchSource = '',
  chatOpen: chatOpenProp,
  onChatToggle,
  searchData,
  aiStatus,
  onRequestAiSetup,
  onRefreshAiStatus
}) => {
  const [prompt, setPrompt] = useState('');
  const [buildState, setBuildState] = useState(() =>
    shouldUseStrictWorkbook(resolveBuildLaunchMode(launchMode)) ? 'initializing' : 'idle'
  );
  const [steps, setSteps] = useState(() =>
    shouldUseStrictWorkbook(resolveBuildLaunchMode(launchMode))
      ? createInitializationSteps({ resolve: 'loading' })
      : []
  );
  const [errorInfo, setErrorInfo] = useState(null);
  const [chatOpenInternal, setChatOpenInternal] = useState(true);
  const chatOpen = chatOpenProp !== undefined ? chatOpenProp : chatOpenInternal;
  const setChatOpen = onChatToggle
    ? () => onChatToggle()
    : setChatOpenInternal;
  const [isBusy, setIsBusy] = useState(false);
  const [location, setLocation] = useState(buildLocation(targetWorkbook?.name, DEFAULT_MODULE_LABEL));
  const [sessionContext, setSessionContext] = useState(null);
  const [createWorkbook, setCreateWorkbook] = useState({ name: PERSONAL_WORKBOOK_NAME, path: '' });
  const [createWorkbookList, setCreateWorkbookList] = useState([]);
  const [createWorkbookOpen, setCreateWorkbookOpen] = useState(false);
  const [messages, setMessages] = useState([]);
  const [splitPct, setSplitPct] = useState(35);
  const [exitDialog, setExitDialog] = useState(null);
  const [savedShortcutLetter, setSavedShortcutLetter] = useState('');
  const [draftShortcutLetter, setDraftShortcutLetter] = useState('');
  const [shortcutInputError, setShortcutInputError] = useState(false);
  const [shortcutSaving, setShortcutSaving] = useState(false);
  const [showIconPicker, setShowIconPicker] = useState(false);
  const [sessionContextMenu, setSessionContextMenu] = useState(null);
  const [iconVersion, setIconVersion] = useState(0);

  // Sync icon changes from other tabs (Shortcuts, Files)
  useEffect(() => {
    const handler = () => setIconVersion(v => v + 1);
    window.addEventListener('macroflow-icon-change', handler);
    return () => window.removeEventListener('macroflow-icon-change', handler);
  }, []);
  const { sessions, saveSession, updateSession, restoreSession, removeSession } = useSessionHistory();

  const normalizedWorkbook = useMemo(
    () => normalizeBuildWorkbook(targetWorkbook),
    [targetWorkbook?.key, targetWorkbook?.name, targetWorkbook?.path]
  );
  const normalizedLaunchMode = useMemo(() => resolveBuildLaunchMode(launchMode), [launchMode]);
  const strictLaunchMode = useMemo(
    () => shouldUseStrictWorkbook(normalizedLaunchMode),
    [normalizedLaunchMode]
  );
  const requestedLaunchModuleName = useMemo(
    () => String(launchModuleName || '').trim(),
    [launchModuleName]
  );

  const buildStateRef = useRef(buildState);
  const promptRef = useRef(prompt);
  const chatOpenRef = useRef(chatOpen);
  const isBusyRef = useRef(isBusy);
  const sessionContextRef = useRef(sessionContext);
  const savedShortcutLetterRef = useRef(savedShortcutLetter);
  const draftShortcutLetterRef = useRef(draftShortcutLetter);
  const shortcutSavingRef = useRef(shortcutSaving);
  const bootstrapSessionRef = useRef(() => Promise.resolve());
  const attemptExitRef = useRef(() => Promise.resolve());
  const activeBootstrapIdRef = useRef(0);
  const activeShortcutHydrationIdRef = useRef(0);
  const exitInFlightRef = useRef(false);
  const messagesRef = useRef(messages);
  const activeSessionIdRef = useRef(null);
  buildStateRef.current = buildState;
  promptRef.current = prompt;
  chatOpenRef.current = chatOpen;
  isBusyRef.current = isBusy;
  sessionContextRef.current = sessionContext;
  savedShortcutLetterRef.current = savedShortcutLetter;
  draftShortcutLetterRef.current = draftShortcutLetter;
  shortcutSavingRef.current = shortcutSaving;
  messagesRef.current = messages;

  // --- Extracted hooks ---
  const {
    editedCode, setEditedCode, editedCodeRef, hasPendingChanges,
    dirtyLocalRef, setDirtyState, markLocalDirty
  } = useCodeEditing();

  const {
    runOutcome, setRunOutcome, savedMacroName, setSavedMacroName,
    sessionMacroName, sessionMacroTarget, handleRunMacro
  } = useMacroExecution({ sessionContextRef, editedCodeRef, editedCode, sessionContext, isBusyRef, setErrorInfo, setIsBusy });
  const handleLocalAiSetup = useCallback(async () => {
    if (typeof onRequestAiSetup !== 'function') {
      setErrorInfo({
        title: 'Local AI setup API is unavailable. Restart MacroFlow dev mode to load the new preload bridge.',
        line: null
      });
      return;
    }

    try {
      await onRequestAiSetup();
    } catch (error) {
      const message = String(error?.message || 'Unable to start local AI setup.');
      setErrorInfo({ title: message, line: null });
    }
  }, [onRequestAiSetup]);

  const aiChecking = aiStatus == null;
  const aiReady = Boolean(aiStatus?.ready);
  const aiSetupInProgress = Boolean(aiStatus?.setupInProgress);
  const aiRemoveInProgress = Boolean(aiStatus?.removeInProgress);
  const aiOperationInProgress = aiChecking || aiSetupInProgress || aiRemoveInProgress;
  const aiProgressPercent = typeof aiStatus?.progress === 'number'
    ? Math.max(0, Math.min(100, Math.round(aiStatus.progress * 100)))
    : null;

  // Auto-start Ollama when Create page mounts and runtime is installed but not running.
  useEffect(() => {
    if (
      aiStatus?.runtimeInstalled &&
      !aiStatus?.serverReachable &&
      !aiStatus?.setupInProgress &&
      !aiStatus?.removeInProgress
    ) {
      window.excel?.ai?.ensureReady?.().catch(() => {});
    }
  }, [aiStatus?.runtimeInstalled, aiStatus?.serverReachable, aiStatus?.setupInProgress, aiStatus?.removeInProgress]);


  const setStepStatus = useCallback((stepKey, status, textOverride = null) => {
    setSteps((previous) =>
      previous.map((step) =>
        step.key === stepKey
          ? {
              ...step,
              status,
              text: textOverride ? String(textOverride) : step.text
            }
          : step
      )
    );
  }, []);

  const stopWithError = useCallback((message, options = {}) => {
    setBuildState('error');
    setIsBusy(false);
    setErrorInfo({
      title: String(message || '').trim() || 'Build failed.',
      line: null,
      canRebind: Boolean(options.canRebind),
      restartRequired: Boolean(options.restartRequired)
    });
    setSavedShortcutLetter('');
    setDraftShortcutLetter('');
    setShortcutInputError(false);
    setShortcutSaving(false);
  }, []);

  const ensureBuildApis = useCallback(() => {
    const apis = {
      modulesByWorkbook: window.excel?.vba?.modulesByWorkbook,
      injectByWorkbook: window.excel?.vba?.injectByWorkbook,
      moduleCodeByWorkbook: window.excel?.vba?.moduleCodeByWorkbook,
      setModuleCodeByWorkbook: window.excel?.vba?.setModuleCodeByWorkbook,
      workbookInfo: window.excel?.workbook?.info
    };

    const missingBoundaryApis = ['moduleCodeByWorkbook', 'setModuleCodeByWorkbook']
      .filter((apiKey) => typeof apis[apiKey] !== 'function');
    if (missingBoundaryApis.length > 0) {
      return {
        ok: false,
        restartRequired: true,
        message: RESTART_REQUIRED_MESSAGE
      };
    }

    const requiredCoreApis = ['modulesByWorkbook', 'injectByWorkbook', 'workbookInfo'];
    const missingCoreApis = requiredCoreApis.filter((apiKey) => typeof apis[apiKey] !== 'function');
    if (missingCoreApis.length > 0) {
      return {
        ok: false,
        restartRequired: false,
        message: `Build APIs are unavailable: ${missingCoreApis.join(', ')}.`
      };
    }

    return {
      ok: true,
      restartRequired: false,
      apis
    };
  }, []);

  const invalidateWorkbookMutation = useCallback((workbook, options = {}) => {
    const normalizedTargetWorkbook = normalizeBuildWorkbook(workbook);
    if (!normalizedTargetWorkbook) {
      return;
    }

    const activeWorkbook = normalizeBuildWorkbook(searchData?.workbook);
    const affectsActiveWorkbook = Boolean(
      normalizedTargetWorkbook?.key
      && activeWorkbook?.key
      && normalizedTargetWorkbook.key === activeWorkbook.key
    );
    const affectsPersonalWorkbook =
      String(normalizedTargetWorkbook?.name || '').trim().toUpperCase() === PERSONAL_WORKBOOK_NAME;

    invalidateSearchBuckets(buildWorkbookInvalidationDescriptors({
      workbook: normalizedTargetWorkbook,
      includeActiveWorkbook: options.includeActiveWorkbook === true && affectsActiveWorkbook,
      includeWorkbookList: options.includeWorkbookList === true,
      includeExplorerAllFiles: options.includeExplorerAllFiles === true,
      includePersonalMacros: options.includePersonalMacros === true && affectsPersonalWorkbook,
      includeShortcutAudit: options.includeShortcutAudit === true,
      includeWorkbookScopedData: options.includeWorkbookScopedData === true
    }));
  }, [searchData?.workbook]);

  const finalizeExit = useCallback((intent) => {
    if (sessionContextRef.current && messagesRef.current.length > 0) {
      const payload = {
        sessionContext: sessionContextRef.current,
        messages: messagesRef.current,
        editedCode: editedCodeRef.current
      };
      if (activeSessionIdRef.current && updateSession(activeSessionIdRef.current, payload)) {
        // updated in place
      } else {
        activeSessionIdRef.current = saveSession(payload);
      }
    }
    // Clear refs so no other code path can re-save the same session
    activeSessionIdRef.current = null;
    sessionContextRef.current = null;
    messagesRef.current = [];
    const normalizedIntent = normalizeExitIntent(intent);
    if (normalizedIntent === 'close') {
      onClose?.();
      return;
    }
    onBack?.();
  }, [onBack, onClose, saveSession, updateSession]);

  const persistUnsyncedChanges = useCallback(async () => {
    const setModuleCodeByWorkbookApi = window.excel?.vba?.setModuleCodeByWorkbook;
    if (typeof setModuleCodeByWorkbookApi !== 'function') {
      throw new Error(RESTART_REQUIRED_MESSAGE);
    }

    const session = sessionContextRef.current;
    if (!session) {
      return;
    }

    const writeResult = await setModuleCodeByWorkbookApi({
      workbookName: session.workbook.name,
      workbookPath: session.workbook.path,
      moduleName: session.moduleName,
      code: String(editedCodeRef.current || ''),
      createIfMissing: true
    });

    if (!writeResult?.success) {
      throw new Error(writeResult?.message || 'Unable to save changes before exit.');
    }
    if (writeResult?.workbookFound === false) {
      const workbookLabel = session.workbook.path || session.workbook.name || 'target workbook';
      throw new Error(`Workbook "${workbookLabel}" is no longer open.`);
    }

    const savedWorkbook = normalizeBuildWorkbook(writeResult?.workbook || session.workbook) || session.workbook;
    const savedModuleName = String(writeResult?.moduleName || session.moduleName).trim() || session.moduleName;
    const updatedSession = {
      ...session,
      workbook: savedWorkbook,
      moduleName: savedModuleName
    };
    sessionContextRef.current = updatedSession;
    setSessionContext(updatedSession);
    setLocation(buildLocation(savedWorkbook.name, savedModuleName));

    setDirtyState(false);
    invalidateWorkbookMutation(savedWorkbook, {
      includeActiveWorkbook: true,
      includeExplorerAllFiles: true,
      includePersonalMacros: true,
      includeShortcutAudit: true,
      includeWorkbookScopedData: true
    });
  }, [invalidateWorkbookMutation, setDirtyState]);

  const cleanupEmptyModule = useCallback(() => {
    const session = sessionContextRef.current;
    if (!session) return;
    const deleteApi = window.excel?.vba?.deleteModuleByWorkbook;
    if (typeof deleteApi !== 'function') return;
    const moduleName = session.moduleName;
    if (!moduleName) return;
    // Only clean up if the module has no real macro (no Sub/Function saved)
    const hasMacro = Boolean(extractPrimaryMacroName(editedCodeRef.current));
    if (hasMacro) return;
    void deleteApi({
      workbookName: session.workbook.name,
      workbookPath: session.workbook.path,
      moduleName
    }).then(() => {
      invalidateWorkbookMutation(session.workbook, {
        includeActiveWorkbook: true,
        includeExplorerAllFiles: true,
        includePersonalMacros: true,
        includeWorkbookScopedData: true
      });
    }).catch(() => {});
  }, [invalidateWorkbookMutation]);

  const pendingInternalActionHandlerRef = useRef(null);

  const runPendingInternalAction = useCallback(() => {
    const pending = pendingInternalActionRef.current;
    pendingInternalActionRef.current = null;
    if (!pending) return false;
    const handler = pendingInternalActionHandlerRef.current;
    if (handler) {
      handler(pending);
    }
    return true;
  }, []);

  const confirmAndSaveExit = useCallback(async (intent) => {
    const normalizedIntent = normalizeExitIntent(intent);
    exitInFlightRef.current = true;
    setIsBusy(true);

    try {
      await persistUnsyncedChanges();
      setExitDialog(null);
      if (!runPendingInternalAction()) {
        finalizeExit(normalizedIntent);
      }
    } catch (error) {
      const message = String(error?.message || 'Unable to save changes before exit.');
      setExitDialog({
        intent: normalizedIntent,
        message,
        type: 'error'
      });
    } finally {
      exitInFlightRef.current = false;
      setIsBusy(false);
    }
  }, [finalizeExit, persistUnsyncedChanges, runPendingInternalAction]);

  const attemptExit = useCallback(async (intent = 'back') => {
    if (exitInFlightRef.current) {
      return;
    }

    const normalizedIntent = normalizeExitIntent(intent);
    const hasSession = Boolean(sessionContextRef.current);
    const shouldSyncOnExit = shouldSyncOnBuildExit({
      hasSession,
      hasPendingChanges: dirtyLocalRef.current
    });
    if (!shouldSyncOnExit) {
      finalizeExit(normalizedIntent);
      return;
    }

    setExitDialog({ intent: normalizedIntent, type: 'confirm' });
  }, [finalizeExit]);

  attemptExitRef.current = attemptExit;
  if (parentAttemptExitRef) {
    parentAttemptExitRef.current = attemptExit;
  }

  useEffect(() => {
    return () => {
      if (parentAttemptExitRef) {
        parentAttemptExitRef.current = null;
      }
    };
  }, [parentAttemptExitRef]);

  useEffect(() => {
    const setSelectedWorkbookApi = window.excel?.security?.setSelectedWorkbook;
    if (typeof setSelectedWorkbookApi !== 'function') {
      return;
    }

    const workbookName = String(sessionContext?.workbook?.name || '').trim();
    const workbookPath = String(sessionContext?.workbook?.path || '').trim();
    if (!workbookName && !workbookPath) {
      return;
    }

    void setSelectedWorkbookApi({ workbookName, workbookPath });
  }, [sessionContext?.workbook?.name, sessionContext?.workbook?.path]);

  const bootstrapSession = useCallback(async () => {
    const strictWorkbookMode = shouldUseStrictWorkbook(normalizedLaunchMode);
    const requestedModuleName = requestedLaunchModuleName;
    const bootstrapId = activeBootstrapIdRef.current + 1;
    activeBootstrapIdRef.current = bootstrapId;

    setBuildState('initializing');
    setIsBusy(true);
    setErrorInfo(null);
    setExitDialog(null);
    setSessionContext(null);
    setRunOutcome('idle');
    setMessages([]);
    setPrompt('');
    setEditedCode(BUILD_MODE_SEED_CODE);
    setSavedMacroName('');
    setSteps(createInitializationSteps({ resolve: 'loading' }));
    setLocation(
      buildLocation(
        normalizedWorkbook?.name,
        strictWorkbookMode && requestedModuleName ? requestedModuleName : DEFAULT_MODULE_LABEL
      )
    );
    setSavedShortcutLetter('');
    setDraftShortcutLetter('');
    setShortcutInputError(false);
    setShortcutSaving(false);

    setDirtyState(false);

    const apiCheck = ensureBuildApis();
    if (!apiCheck.ok) {
      stopWithError(apiCheck.message, { restartRequired: apiCheck.restartRequired });
      return;
    }
    const {
      modulesByWorkbook: modulesByWorkbookApi,
      injectByWorkbook: injectByWorkbookApi,
      moduleCodeByWorkbook: moduleCodeByWorkbookApi,
      workbookInfo: workbookInfoApi
    } = apiCheck.apis;

    try {
      let selectedResolution = null;
      let activeResolution = null;
      let selectedUnavailable = false;

      const selectedCandidate = normalizeBuildWorkbook(normalizedWorkbook);
      if (selectedCandidate) {
        const selectedModulesResult = await modulesByWorkbookApi(toWorkbookRequest(selectedCandidate));
        if (activeBootstrapIdRef.current !== bootstrapId) {
          return;
        }
        if (!selectedModulesResult?.success) {
          throw new Error(selectedModulesResult?.message || 'Unable to inspect selected workbook.');
        }
        if (selectedModulesResult?.workbookFound === false) {
          selectedUnavailable = true;
        } else {
          selectedResolution = {
            workbook:
              normalizeBuildWorkbook(selectedModulesResult?.workbook || selectedCandidate) || selectedCandidate,
            modules: Array.isArray(selectedModulesResult?.modules) ? selectedModulesResult.modules : []
          };
        }
      }

      if (!selectedResolution && !strictWorkbookMode) {
        const workbookInfoResult = await workbookInfoApi();
        if (activeBootstrapIdRef.current !== bootstrapId) {
          return;
        }
        if (!workbookInfoResult?.success) {
          throw new Error(workbookInfoResult?.message || 'No active workbook is available.');
        }

        const activeCandidate = toActiveWorkbookModel(workbookInfoResult);
        if (!activeCandidate) {
          throw new Error('No active workbook is available.');
        }

        const activeModulesResult = await modulesByWorkbookApi(toWorkbookRequest(activeCandidate));
        if (activeBootstrapIdRef.current !== bootstrapId) {
          return;
        }
        if (!activeModulesResult?.success) {
          throw new Error(activeModulesResult?.message || 'Unable to inspect active workbook.');
        }
        if (activeModulesResult?.workbookFound === false) {
          throw new Error('No valid workbook is available for Build Mode.');
        }

        activeResolution = {
          workbook: normalizeBuildWorkbook(activeModulesResult?.workbook || activeCandidate) || activeCandidate,
          modules: Array.isArray(activeModulesResult?.modules) ? activeModulesResult.modules : []
        };
      }

      if (strictWorkbookMode && normalizedLaunchMode === 'existing_module') {
        if (!selectedCandidate) {
          throw new Error('No workbook was provided for the selected module.');
        }
        if (!selectedResolution || selectedUnavailable) {
          const workbookLabel = selectedCandidate.path || selectedCandidate.name || 'selected workbook';
          throw new Error(`Workbook "${workbookLabel}" is not open.`);
        }
        if (!requestedModuleName) {
          throw new Error('No module was provided for Build Mode.');
        }
      }

      const resolvedWorkbook = (strictWorkbookMode && normalizedLaunchMode === 'existing_module')
        ? normalizeBuildWorkbook(selectedResolution?.workbook || selectedCandidate)
        : resolveBuildWorkbookTarget({
            selectedWorkbook: selectedResolution?.workbook,
            selectedWorkbookFound: Boolean(selectedResolution),
            activeWorkbook: activeResolution?.workbook
          });
      if (!resolvedWorkbook) {
        throw new Error('No valid workbook is available for Build Mode.');
      }

      // Set security boundary before any high-risk operations (inject/code write)
      const setSelectedWorkbookApi = window.excel?.security?.setSelectedWorkbook;
      if (typeof setSelectedWorkbookApi === 'function') {
        await setSelectedWorkbookApi({
          workbookName: resolvedWorkbook.name,
          workbookPath: resolvedWorkbook.path
        });
      }

      const resolvedModules = selectedResolution?.modules || activeResolution?.modules || [];
      const resolveStepText = strictWorkbookMode
        ? `Located workbook: ${resolvedWorkbook.name}`
        : selectedUnavailable && activeResolution
          ? `Located workbook: ${resolvedWorkbook.name} (active fallback)`
          : `Located workbook: ${resolvedWorkbook.name}`;
      setStepStatus('resolve', 'complete', resolveStepText);

      let sessionWorkbook = resolvedWorkbook;
      let sessionModuleName = '';
      let pulledCode = BUILD_MODE_SEED_CODE;

      if (strictWorkbookMode && normalizedLaunchMode === 'existing_module' && requestedModuleName) {
        setStepStatus('create', 'loading', `Locating module: ${requestedModuleName}`);

        const existingModuleName = resolveExistingModuleName(resolvedModules, requestedModuleName);
        if (!existingModuleName) {
          throw new Error(`Module "${requestedModuleName}" was not found in workbook "${resolvedWorkbook.name}".`);
        }

        setStepStatus('create', 'complete', `Located module: ${existingModuleName}`);
        setLocation(buildLocation(resolvedWorkbook.name, existingModuleName));

        const codeResult = await moduleCodeByWorkbookApi({
          workbookName: resolvedWorkbook.name,
          workbookPath: resolvedWorkbook.path,
          moduleName: existingModuleName
        });
        if (activeBootstrapIdRef.current !== bootstrapId) {
          return;
        }
        if (!codeResult?.success) {
          throw new Error(codeResult?.message || 'Unable to load selected module code.');
        }
        if (codeResult?.workbookFound === false) {
          const workbookLabel = resolvedWorkbook.path || resolvedWorkbook.name || 'selected workbook';
          throw new Error(`Workbook "${workbookLabel}" is not open.`);
        }
        if (codeResult?.moduleFound === false) {
          throw new Error(`Module "${existingModuleName}" was not found in workbook "${resolvedWorkbook.name}".`);
        }

        pulledCode = String(codeResult?.code || BUILD_MODE_SEED_CODE);
        sessionWorkbook = normalizeBuildWorkbook(codeResult?.workbook || resolvedWorkbook) || resolvedWorkbook;
        sessionModuleName = String(codeResult?.moduleName || existingModuleName).trim() || existingModuleName;
      } else {
        const nextModuleName = selectNextModuleName(resolvedModules, MODULE_PREFIX);
        setStepStatus('create', 'loading', `Creating module: ${nextModuleName}`);
        setLocation(buildLocation(resolvedWorkbook.name, nextModuleName));

        const injectResult = await injectByWorkbookApi({
          workbookName: resolvedWorkbook.name,
          workbookPath: resolvedWorkbook.path,
          moduleName: nextModuleName,
          code: BUILD_MODE_SEED_CODE,
          createIfMissing: true
        });
        if (activeBootstrapIdRef.current !== bootstrapId) {
          return;
        }
        if (!injectResult?.success) {
          throw new Error(injectResult?.message || 'Unable to create the Build Mode module.');
        }
        if (injectResult?.workbookFound === false) {
          const workbookLabel = resolvedWorkbook.path || resolvedWorkbook.name;
          throw new Error(`Workbook "${workbookLabel}" is not open.`);
        }

        const createdWorkbook =
          normalizeBuildWorkbook(injectResult?.workbook || resolvedWorkbook) || resolvedWorkbook;
        const createdModuleName = String(injectResult?.moduleName || nextModuleName).trim() || nextModuleName;
        setStepStatus('create', 'complete', `Created: ${createdModuleName}`);
        setLocation(buildLocation(createdWorkbook.name, createdModuleName));
        invalidateWorkbookMutation(createdWorkbook, {
          includeActiveWorkbook: true,
          includeExplorerAllFiles: true,
          includePersonalMacros: true,
          includeWorkbookScopedData: true
        });

        const codeResult = await moduleCodeByWorkbookApi({
          workbookName: createdWorkbook.name,
          workbookPath: createdWorkbook.path,
          moduleName: createdModuleName
        });
        if (activeBootstrapIdRef.current !== bootstrapId) {
          return;
        }
        if (!codeResult?.success) {
          throw new Error(codeResult?.message || 'Unable to load module code on entry.');
        }
        if (codeResult?.workbookFound === false || codeResult?.moduleFound === false) {
          throw new Error('Unable to load module code on entry.');
        }

        pulledCode = String(codeResult?.code || BUILD_MODE_SEED_CODE);
        sessionWorkbook = normalizeBuildWorkbook(codeResult?.workbook || createdWorkbook) || createdWorkbook;
        sessionModuleName = String(codeResult?.moduleName || createdModuleName).trim() || createdModuleName;
      }

      const nextSession = {
        workbook: sessionWorkbook,
        moduleName: sessionModuleName,
        createdAt: Date.now()
      };
      sessionContextRef.current = nextSession;
      setSessionContext(nextSession);
      setEditedCode(pulledCode);
      setSavedMacroName(extractPrimaryMacroName(pulledCode));
      setLocation(buildLocation(sessionWorkbook.name, sessionModuleName));
      setBuildState('ready');
      setIsBusy(false);
      if (normalizedLaunchMode === 'new_module_immediate') {
        setChatOpen(true);
      }
    } catch (error) {
      if (activeBootstrapIdRef.current !== bootstrapId) {
        return;
      }
      const message = String(error?.message || 'Unable to initialize Build Mode.');
      stopWithError(`Build session failed: ${message}`);
    } finally {
      if (activeBootstrapIdRef.current === bootstrapId) {
        setIsBusy(false);
      }
    }
  }, [
    ensureBuildApis,
    invalidateWorkbookMutation,
    normalizedLaunchMode,
    normalizedWorkbook,
    requestedLaunchModuleName,
    setStepStatus,
    stopWithError,
    setDirtyState
  ]);

  bootstrapSessionRef.current = bootstrapSession;

  // Wrap hook's markLocalDirty to also reset runOutcome
  const markLocalDirtyFull = useCallback(() => {
    markLocalDirty();
    setRunOutcome('idle');
  }, [markLocalDirty, setRunOutcome]);

  const handleCodeChange = useCallback((newCode) => {
    const nextCode = String(newCode || '');
    setEditedCode(nextCode);
    setShortcutInputError(false);
    markLocalDirtyFull();
  }, [markLocalDirtyFull, setEditedCode]);

  const handleSave = useCallback(async () => {
    if (isBusyRef.current || !dirtyLocalRef.current) {
      return;
    }
    if (!sessionContextRef.current) {
      return;
    }

    setErrorInfo(null);
    setIsBusy(true);
    isBusyRef.current = true;

    try {
      await persistUnsyncedChanges();
      setSavedMacroName(extractPrimaryMacroName(editedCodeRef.current));
      setRunOutcome('idle');
    } catch (error) {
      const message = String(error?.message || 'Unable to save changes.');
      setRunOutcome('error');
      setErrorInfo({
        title: `Save failed: ${message}`,
        line: null
      });
    } finally {
      setIsBusy(false);
      isBusyRef.current = false;
    }
  }, [persistUnsyncedChanges]);

  const handleShortcutDraftChange = useCallback((value) => {
    const normalizedDraft = normalizeShortcutLetterDraft(value);
    setDraftShortcutLetter(normalizedDraft);
    setShortcutInputError(false);
  }, []);

  const commitShortcutDraft = useCallback(async () => {
    if (shortcutSavingRef.current) {
      return;
    }

    const session = sessionContextRef.current;
    if (!session) {
      return;
    }

    const moduleName = String(session?.moduleName || '').trim();
    const macroName = extractPrimaryMacroName(editedCodeRef.current);
    const macroTarget = buildSessionMacroTarget(moduleName, macroName);
    const savedLetter = String(savedShortcutLetterRef.current || '');
    const draftLetter = normalizeShortcutLetterDraft(draftShortcutLetterRef.current);

    if (!macroTarget) {
      setDraftShortcutLetter(savedLetter);
      setShortcutInputError(Boolean(draftLetter));
      return;
    }

    if (!draftLetter) {
      setDraftShortcutLetter(savedLetter);
      setShortcutInputError(false);
      return;
    }

    if (draftLetter === savedLetter) {
      setDraftShortcutLetter(draftLetter);
      setShortcutInputError(false);
      return;
    }

    const workbookName = String(session?.workbook?.name || '').trim();
    const workbookPath = String(session?.workbook?.path || '').trim();
    const auditShortcutsApi = window.excel?.vba?.auditShortcutsByWorkbook;
    const setModuleCodeByWorkbookApi = window.excel?.vba?.setModuleCodeByWorkbook;
    const setShortcutByWorkbookApi = window.excel?.vba?.setShortcutByWorkbook;

    if (
      typeof auditShortcutsApi !== 'function' ||
      typeof setModuleCodeByWorkbookApi !== 'function' ||
      typeof setShortcutByWorkbookApi !== 'function'
    ) {
      setShortcutInputError(true);
      return;
    }

    shortcutSavingRef.current = true;
    setShortcutSaving(true);
    setShortcutInputError(false);

    try {
      const auditResult = await auditShortcutsApi({
        workbookName,
        workbookPath
      });
      if (!auditResult?.success || auditResult?.workbookFound === false) {
        throw new Error('Unable to audit workbook shortcuts.');
      }

      if (hasShortcutConflictForMacro(auditResult, macroTarget, draftLetter)) {
        setShortcutInputError(true);
        return;
      }

      const writeResult = await setModuleCodeByWorkbookApi({
        workbookName,
        workbookPath,
        moduleName,
        code: String(editedCodeRef.current || ''),
        createIfMissing: true
      });
      if (!writeResult?.success || writeResult?.workbookFound === false) {
        throw new Error('Unable to persist module code before shortcut assignment.');
      }

      const excelShortcutKey = toExcelShortcutKeyFromLetter(draftLetter);
      if (!excelShortcutKey) {
        setShortcutInputError(true);
        return;
      }

      const assignResult = await setShortcutByWorkbookApi({
        workbookName,
        workbookPath,
        macroName: macroTarget,
        shortcutKey: excelShortcutKey
      });

      if (!assignResult?.success || assignResult?.workbookFound === false) {
        throw new Error('Unable to assign workbook shortcut.');
      }

      setSavedShortcutLetter(draftLetter);
      setDraftShortcutLetter(draftLetter);
      setShortcutInputError(false);
      invalidateWorkbookMutation(session.workbook, {
        includeActiveWorkbook: true,
        includeExplorerAllFiles: true,
        includePersonalMacros: true,
        includeShortcutAudit: true,
        includeWorkbookScopedData: true
      });
    } catch (_error) {
      setShortcutInputError(true);
    } finally {
      shortcutSavingRef.current = false;
      setShortcutSaving(false);
    }
  }, [invalidateWorkbookMutation]);

  const toggleChat = useCallback(() => {
    if (onChatToggle) {
      onChatToggle();
    } else {
      setChatOpenInternal((previous) => !previous);
    }
  }, [onChatToggle]);

  const handleSubmit = useCallback(async () => {
    const submittedPrompt = String(promptRef.current || '').trim();
    if (isBusyRef.current || !submittedPrompt || !sessionContextRef.current) {
      return;
    }

    if (!aiStatus?.ready) {
      setErrorInfo({
        title: String(aiStatus?.statusText || 'Local AI setup is required before generating VBA.'),
        line: null
      });
      return;
    }

    const generateVbaApi = window.excel?.ai?.generateVba;
    if (typeof generateVbaApi !== 'function') {
      setErrorInfo({
        title: 'Generation failed: AI generation API is unavailable.',
        line: null
      });
      return;
    }

    setMessages((prev) => [...prev, { role: 'user', content: submittedPrompt, timestamp: Date.now() }]);
    setPrompt('');
    setRunOutcome('idle');
    setErrorInfo(null);
    setIsBusy(true);
    isBusyRef.current = true;

    try {
      const session = sessionContextRef.current;
      const intent = inferBuildPromptIntent(submittedPrompt);
      const isCodeIntent = intent === 'create' || intent === 'edit';

      // For ask intent: stream tokens into a chat bubble
      // For code intents: no streaming — loading dots show activity, code goes straight to editor when done
      let placeholderId = null;
      let unsubTokens = () => {};

      if (!isCodeIntent) {
        placeholderId = `stream_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
        setMessages((prev) => [...prev, { role: 'assistant', kind: 'text', content: '', _id: placeholderId, streaming: true, timestamp: Date.now() }]);
        unsubTokens = window.excel?.ai?.onGenerateToken?.((token) => {
          setMessages((prev) => {
            const last = prev[prev.length - 1];
            if (last?._id === placeholderId && last?.streaming) {
              const updated = [...prev];
              updated[updated.length - 1] = { ...last, content: last.content + token };
              return updated;
            }
            return prev;
          });
        }) || (() => {});
      }

      const request = {
        prompt: submittedPrompt,
        intent,
        workbookName: String(session?.workbook?.name || '').trim(),
        workbookPath: String(session?.workbook?.path || '').trim(),
        moduleName: String(session?.moduleName || '').trim(),
        includeCurrentCode: true,
        currentCode: String(editedCodeRef.current || '')
      };

      const result = await generateVbaApi(request);
      unsubTokens();

      if (!result?.success) {
        if (placeholderId) setMessages((prev) => prev.filter((m) => m._id !== placeholderId));
        throw new Error(mapAiGenerationMessage(result));
      }

      if (!isCodeIntent) {
        // Ask intent: finalize the streamed chat message
        setMessages((prev) => prev.map((m) =>
          m._id === placeholderId ? { ...buildAssistantMessageFromResult(result, 'Response received.'), _id: placeholderId } : m
        ));
        return;
      }

      // Code intent: put clean extracted code into the editor
      const generatedCode = String(result?.code || '').trim();
      if (!generatedCode) {
        throw new Error('Local AI returned empty VBA output.');
      }

      setEditedCode(generatedCode);
      setSavedMacroName(extractPrimaryMacroName(generatedCode));
      setMessages((prev) => [...prev, { role: 'assistant', kind: 'text', content: 'Done — code updated in the editor.', timestamp: Date.now() }]);
      markLocalDirtyFull();
    } catch (error) {
      const message = mapAiGenerationMessage(null, String(error?.message || 'Unable to generate VBA.'));
      setMessages((prev) => [...prev, { role: 'assistant', kind: 'text', content: `Error: ${message}`, timestamp: Date.now() }]);
      setErrorInfo({
        title: `Generation failed: ${message}`,
        line: null
      });
    } finally {
      setIsBusy(false);
      isBusyRef.current = false;
    }
  }, [markLocalDirtyFull]);

  const openSidebar = useCallback(() => {
    if (!chatOpenRef.current) {
      if (onChatToggle) {
        onChatToggle();
      } else {
        setChatOpenInternal(true);
      }
    }
  }, [onChatToggle]);

  const handleFirstSubmit = useCallback(async () => {
    const submittedPrompt = String(promptRef.current || '').trim();
    if (isBusyRef.current || !submittedPrompt) return;

    if (!aiStatus?.ready) {
      setErrorInfo({
        title: String(aiStatus?.statusText || 'Local AI setup is required before generating VBA.'),
        line: null
      });
      return;
    }

    const generateVbaApi = window.excel?.ai?.generateVba;
    if (typeof generateVbaApi !== 'function') {
      setErrorInfo({ title: 'Generation failed: AI generation API is unavailable.', line: null });
      return;
    }

    setMessages([{ role: 'user', content: submittedPrompt, timestamp: Date.now() }]);
    setPrompt('');
    setBuildState('initializing');
    setIsBusy(true);
    openSidebar();
    isBusyRef.current = true;
    setErrorInfo(null);
    setRunOutcome('idle');

    try {
      const apiCheck = ensureBuildApis();
      if (!apiCheck.ok) {
        stopWithError(apiCheck.message, { restartRequired: apiCheck.restartRequired });
        return;
      }
      const {
        modulesByWorkbook: modulesByWorkbookApi,
        injectByWorkbook: injectByWorkbookApi,
        workbookInfo: workbookInfoApi
      } = apiCheck.apis;

      // Phase 1: Resolve workbook (use dropdown selection if available)
      let selectedResolution = null;
      let activeResolution = null;

      const selectedCandidate = normalizeBuildWorkbook(createWorkbook || normalizedWorkbook);
      if (selectedCandidate) {
        const selectedModulesResult = await modulesByWorkbookApi(toWorkbookRequest(selectedCandidate));
        if (!selectedModulesResult?.success) {
          throw new Error(selectedModulesResult?.message || 'Unable to inspect selected workbook.');
        }
        if (selectedModulesResult?.workbookFound !== false) {
          selectedResolution = {
            workbook: normalizeBuildWorkbook(selectedModulesResult?.workbook || selectedCandidate) || selectedCandidate,
            modules: Array.isArray(selectedModulesResult?.modules) ? selectedModulesResult.modules : []
          };
        }
      }

      if (!selectedResolution) {
        const workbookInfoResult = await workbookInfoApi();
        if (!workbookInfoResult?.success) {
          throw new Error(workbookInfoResult?.message || 'No active workbook is available.');
        }
        const activeCandidate = toActiveWorkbookModel(workbookInfoResult);
        if (!activeCandidate) {
          throw new Error('No active workbook is available.');
        }
        const activeModulesResult = await modulesByWorkbookApi(toWorkbookRequest(activeCandidate));
        if (!activeModulesResult?.success) {
          throw new Error(activeModulesResult?.message || 'Unable to inspect active workbook.');
        }
        if (activeModulesResult?.workbookFound === false) {
          throw new Error('No valid workbook is available for Build Mode.');
        }
        activeResolution = {
          workbook: normalizeBuildWorkbook(activeModulesResult?.workbook || activeCandidate) || activeCandidate,
          modules: Array.isArray(activeModulesResult?.modules) ? activeModulesResult.modules : []
        };
      }

      const resolvedWorkbook = resolveBuildWorkbookTarget({
        selectedWorkbook: selectedResolution?.workbook,
        selectedWorkbookFound: Boolean(selectedResolution),
        activeWorkbook: activeResolution?.workbook
      });
      if (!resolvedWorkbook) {
        throw new Error('No valid workbook is available for Build Mode.');
      }
      const resolvedModules = selectedResolution?.modules || activeResolution?.modules || [];

      // Phase 2: Create module
      const nextModuleName = selectNextModuleName(resolvedModules, MODULE_PREFIX);
      setLocation(buildLocation(resolvedWorkbook.name, nextModuleName));

      const injectResult = await injectByWorkbookApi({
        workbookName: resolvedWorkbook.name,
        workbookPath: resolvedWorkbook.path,
        moduleName: nextModuleName,
        code: BUILD_MODE_SEED_CODE,
        createIfMissing: true
      });
      if (!injectResult?.success) {
        throw new Error(injectResult?.message || 'Unable to create the Build Mode module.');
      }
      if (injectResult?.workbookFound === false) {
        const workbookLabel = resolvedWorkbook.path || resolvedWorkbook.name;
        throw new Error(`Workbook "${workbookLabel}" is not open.`);
      }

      const createdWorkbook = normalizeBuildWorkbook(injectResult?.workbook || resolvedWorkbook) || resolvedWorkbook;
      const createdModuleName = String(injectResult?.moduleName || nextModuleName).trim() || nextModuleName;
      setLocation(buildLocation(createdWorkbook.name, createdModuleName));
      invalidateWorkbookMutation(createdWorkbook, {
        includeActiveWorkbook: true,
        includeExplorerAllFiles: true,
        includePersonalMacros: true,
        includeWorkbookScopedData: true
      });

      // Phase 3: Set session context
      const nextSession = {
        workbook: createdWorkbook,
        moduleName: createdModuleName,
        createdAt: Date.now()
      };
      sessionContextRef.current = nextSession;
      setSessionContext(nextSession);
      openSidebar();

      // Phase 4: Generate VBA (loading dots show activity, code goes to editor when done)

      try {
        const result = await generateVbaApi({
          prompt: submittedPrompt,
          intent: 'create',
          workbookName: createdWorkbook.name,
          workbookPath: createdWorkbook.path,
          moduleName: createdModuleName,
          includeCurrentCode: false
        });
        if (!result?.success) {
          throw new Error(mapAiGenerationMessage(result));
        }
        const generatedCode = String(result?.code || '').trim();
        if (!generatedCode) {
          throw new Error('Local AI returned empty VBA output.');
        }

        setEditedCode(generatedCode);
        setSavedMacroName(extractPrimaryMacroName(generatedCode));
        setMessages((prev) => [...prev, { role: 'assistant', kind: 'text', content: 'Done — macro generated in the editor.', timestamp: Date.now() }]);
        setBuildState('ready');
        markLocalDirtyFull();
      } catch (streamError) {
        throw streamError;
      }
    } catch (error) {
      const message = String(error?.message || 'Unable to generate VBA.');
      setMessages((prev) => [...prev, { role: 'assistant', kind: 'text', content: `Error: ${message}`, timestamp: Date.now() }]);
      setBuildState('idle');
      setErrorInfo({ title: `Generation failed: ${message}`, line: null });
    } finally {
      setIsBusy(false);
      isBusyRef.current = false;
    }
  }, [ensureBuildApis, invalidateWorkbookMutation, normalizedWorkbook, markLocalDirtyFull, stopWithError, openSidebar]);

  const dispatchSubmit = useCallback(() => {
    if (sessionContextRef.current) {
      // Session exists — open sidebar immediately, it already shows the module chat
      openSidebar();
      void handleSubmit();
    } else {
      // First submit — sidebar will be opened after session is created in handleFirstSubmit
      void handleFirstSubmit();
    }
  }, [handleSubmit, handleFirstSubmit, openSidebar]);

  const doRestoreSession = useCallback((id) => {
    const saved = restoreSession(id);
    if (!saved) return;

    activeSessionIdRef.current = id;
    setSessionContext(saved.sessionContext);
    sessionContextRef.current = saved.sessionContext;
    setMessages(saved.messages);
    messagesRef.current = saved.messages;
    setEditedCode(saved.editedCode);
    editedCodeRef.current = saved.editedCode;
    setSavedMacroName(extractPrimaryMacroName(saved.editedCode));
    setLocation(buildLocation(saved.sessionContext.workbook.name, saved.sessionContext.moduleName));
    setBuildState('ready');
    setErrorInfo(null);
    setRunOutcome('idle');
    setDirtyState(false);
    setPrompt('');
    setSavedShortcutLetter('');
    setDraftShortcutLetter('');
    setShortcutInputError(false);
    setShortcutSaving(false);
  }, [restoreSession, setDirtyState]);

  const pendingInternalActionRef = useRef(null);

  const handleRestoreSession = useCallback((id) => {
    if (dirtyLocalRef.current && sessionContextRef.current) {
      pendingInternalActionRef.current = { type: 'restore', id };
      setExitDialog({ intent: 'back', type: 'confirm' });
      return;
    }
    doRestoreSession(id);
  }, [doRestoreSession]);

  const doBackToHistory = useCallback(() => {
    if (sessionContextRef.current && messagesRef.current.length > 0) {
      const payload = {
        sessionContext: sessionContextRef.current,
        messages: messagesRef.current,
        editedCode: editedCodeRef.current
      };
      if (activeSessionIdRef.current && updateSession(activeSessionIdRef.current, payload)) {
        // updated in place
      } else {
        activeSessionIdRef.current = saveSession(payload);
      }
    }
    activeSessionIdRef.current = null;
    setSessionContext(null);
    sessionContextRef.current = null;
    setMessages([]);
    messagesRef.current = [];
    setEditedCode(BUILD_MODE_SEED_CODE);
    editedCodeRef.current = BUILD_MODE_SEED_CODE;
    setBuildState('idle');
    setErrorInfo(null);
    setRunOutcome('idle');
    setDirtyState(false);
    setPrompt('');
    setSavedMacroName('');
    setSavedShortcutLetter('');
    setDraftShortcutLetter('');
    setShortcutInputError(false);
    setShortcutSaving(false);
    setLocation(buildLocation(normalizedWorkbook?.name, DEFAULT_MODULE_LABEL));
  }, [saveSession, updateSession, setDirtyState, normalizedWorkbook?.name]);

  pendingInternalActionHandlerRef.current = (pending) => {
    if (pending.type === 'restore') {
      doRestoreSession(pending.id);
    } else if (pending.type === 'history') {
      doBackToHistory();
    }
  };

  const handleBackToHistory = useCallback(() => {
    if (dirtyLocalRef.current && sessionContextRef.current) {
      pendingInternalActionRef.current = { type: 'history' };
      setExitDialog({ intent: 'back', type: 'confirm' });
      return;
    }
    doBackToHistory();
  }, [doBackToHistory]);

  // --- Module rename hook (needs invalidateWorkbookMutation defined above) ---
  const {
    moduleRenameActive, setModuleRenameActive, moduleRenameDraft, setModuleRenameDraft,
    moduleRenameInputRef, moduleRenameRestoreValueRef, cancelModuleRename, commitModuleRename
  } = useModuleRename({ sessionContextRef, location, buildLocationFn: buildLocation, invalidateWorkbookMutation, setSessionContext, setLocation, setErrorInfo, setIsBusy });



  const shortcutPrefix = formatShortcutPrefix(draftShortcutLetter);

  // Build macro ID for icon lookup — the store normalizes keys so kind/scope suffix doesn't matter
  const currentMacroId = useMemo(() => {
    const mod = sessionContext?.moduleName;
    const name = sessionMacroName;
    if (!mod || !name) return null;
    const workbookKey = String(sessionContext?.workbook?.path || sessionContext?.workbook?.name || '').trim();
    return workbookKey
      ? `${workbookKey}::macro::${mod}::${name}`
      : `${mod}::${name}`;
  }, [sessionContext?.moduleName, sessionMacroName, sessionContext?.workbook?.path, sessionContext?.workbook?.name]);

  const currentMacroIcon = currentMacroId ? getMacroIcon(currentMacroId) : null;

  const prevMacroIdRef = useRef(currentMacroId);
  useEffect(() => {
    const prevId = prevMacroIdRef.current;
    prevMacroIdRef.current = currentMacroId;
    if (prevId && currentMacroId && prevId !== currentMacroId) {
      renameMacroIcon(prevId, currentMacroId);
    }
  }, [currentMacroId]);

  useEffect(() => {
    if (!sessionContext || !sessionMacroTarget) {
      setSavedShortcutLetter('');
      setDraftShortcutLetter('');
      setShortcutInputError(false);
      return;
    }

    const auditShortcutsApi = window.excel?.vba?.auditShortcutsByWorkbook;
    if (typeof auditShortcutsApi !== 'function') {
      setSavedShortcutLetter('');
      setDraftShortcutLetter('');
      setShortcutInputError(false);
      return;
    }

    const requestId = activeShortcutHydrationIdRef.current + 1;
    activeShortcutHydrationIdRef.current = requestId;
    let cancelled = false;

    const loadAssignedShortcut = async () => {
      try {
        const result = await auditShortcutsApi({
          workbookName: sessionContext.workbook.name,
          workbookPath: sessionContext.workbook.path
        });
        if (cancelled || requestId !== activeShortcutHydrationIdRef.current) {
          return;
        }
        if (!result?.success || result?.workbookFound === false) {
          setSavedShortcutLetter('');
          setDraftShortcutLetter('');
          setShortcutInputError(false);
          return;
        }

        const hydratedLetter = findAssignedShortcutLetterForMacro(result, sessionMacroTarget);
        setSavedShortcutLetter(hydratedLetter);
        setDraftShortcutLetter(hydratedLetter);
        setShortcutInputError(false);
      } catch (_error) {
        if (cancelled || requestId !== activeShortcutHydrationIdRef.current) {
          return;
        }
        setSavedShortcutLetter('');
        setDraftShortcutLetter('');
        setShortcutInputError(false);
      }
    };

    void loadAssignedShortcut();

    return () => {
      cancelled = true;
    };
  }, [
    sessionContext?.moduleName,
    sessionContext?.workbook?.name,
    sessionContext?.workbook?.path,
    sessionMacroTarget
  ]);

  useEffect(() => {
    if (!strictLaunchMode) {
      setBuildState('idle');
      setIsBusy(false);
      isBusyRef.current = false;
      return;
    }

    const bootstrapTimerId = window.setTimeout(() => {
      void bootstrapSessionRef.current();
    }, 0);

    return () => {
      window.clearTimeout(bootstrapTimerId);
      activeBootstrapIdRef.current += 1;
    };
  }, [
    strictLaunchMode,
    normalizedLaunchMode,
    normalizedWorkbook?.key,
    normalizedWorkbook?.name,
    normalizedWorkbook?.path,
    requestedLaunchModuleName
  ]);

  useEffect(() => {
    if (moduleRenameActive) {
      return;
    }
    const currentName = String(sessionContext?.moduleName || location.module || '').trim();
    setModuleRenameDraft(currentName);
  }, [location.module, moduleRenameActive, sessionContext?.moduleName]);

  useEffect(() => {
    if (!moduleRenameActive) {
      return undefined;
    }

    const timerId = window.setTimeout(() => {
      const input = moduleRenameInputRef.current;
      if (!input) {
        return;
      }
      input.focus();
      const selection = window.getSelection();
      if (!selection) {
        return;
      }
      const range = document.createRange();
      range.selectNodeContents(input);
      range.collapse(false);
      selection.removeAllRanges();
      selection.addRange(range);
    }, 0);

    return () => window.clearTimeout(timerId);
  }, [moduleRenameActive]);

  // Keyboard shortcut handler uses refs so the listener is registered once.
  useEffect(() => {
    const handleKeyDown = (event) => {
      if (event.ctrlKey && event.key === 's' && !chatOpenRef.current && sessionContextRef.current) {
        event.preventDefault();
        void handleSave();
        return;
      }

      if (event.key === 'Escape') {
        event.preventDefault();
        void attemptExitRef.current('back');
        return;
      }

      if (!event.ctrlKey && !event.altKey && !event.metaKey && !event.shiftKey && event.key === 'Tab') {
        // Don't intercept Tab when user is in the code editor — let it indent
        if (event.target?.isContentEditable) return;
        event.preventDefault();
        void attemptExitRef.current('back');
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [handleSave]);

  const getFooterContent = () => {
    switch (runOutcome) {
      case 'success':
        return {
          logoClass: 'success',
          text: 'Macro ran successfully!',
          textClass: 'success'
        };
      case 'error':
        return {
          logoClass: 'error',
          text: 'Run failed',
          textClass: 'error'
        };
      default:
        return {
          logoClass: '',
          text: null
        };
    }
  };

  const footerContent = getFooterContent();
  const showCodePanel = Boolean(sessionContext);
  const shortcutInputDisabled =
    !showCodePanel || !sessionMacroTarget || isBusy || shortcutSaving;
  const primaryActionLabel = hasPendingChanges ? 'Save changes' : 'Run macro';
  const codeStatus =
    runOutcome === 'success'
      ? 'success'
      : runOutcome === 'error' || buildState === 'error'
        ? 'error'
        : 'normal';
  const currentCode = String(editedCode || BUILD_MODE_SEED_CODE);
  const renderModuleBreadcrumb = () => {
    if (!showCodePanel) {
      return displayMacroName(location.module);
    }

    if (moduleRenameActive) {
      return (
        <span
          ref={moduleRenameInputRef}
          className="build-module-rename-editable"
          contentEditable={!isBusy}
          suppressContentEditableWarning
          role="textbox"
          aria-label="Rename module"
          onBlur={() => {
            const nextValue = String(moduleRenameInputRef.current?.textContent || '').trim();
            void commitModuleRename(nextValue);
          }}
          onKeyDown={(event) => {
            if (event.key === 'Enter') {
              event.preventDefault();
              const nextValue = String(moduleRenameInputRef.current?.textContent || '').trim();
              void commitModuleRename(nextValue);
              return;
            }
            if (event.key === 'Escape') {
              event.preventDefault();
              if (moduleRenameInputRef.current) {
                moduleRenameInputRef.current.textContent = moduleRenameRestoreValueRef.current;
              }
              cancelModuleRename();
            }
          }}
          onPaste={(e) => {
            e.preventDefault();
            const text = e.clipboardData.getData('text/plain');
            document.execCommand('insertText', false, text);
          }}
        >
          {moduleRenameDraft}
        </span>
      );
    }

    return (
      <button
        type="button"
        className="code-breadcrumb-module-btn"
        onClick={() => {
          if (!isBusy) {
            const currentName = String(location.module || '').trim();
            moduleRenameRestoreValueRef.current = currentName;
            setModuleRenameDraft(displayMacroName(currentName));
            setModuleRenameActive(true);
          }
        }}
        disabled={isBusy}
      >
        {displayMacroName(location.module)}
      </button>
    );
  };

  const conversationEndRef = useRef(null);

  useEffect(() => {
    conversationEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages.length, isBusy]);

  const aiGateOverlay = !aiReady ? (
    <div className="ai-gate-overlay" role="dialog" aria-modal="true" aria-label="Local AI setup required">
      <div className="ai-gate-card">
        <div className="ai-gate-description">
          {aiChecking
            ? 'Checking local AI...'
            : aiSetupInProgress
              ? String(aiStatus?.statusText || 'Setting up local AI...')
              : aiRemoveInProgress
                ? String(aiStatus?.statusText || 'Removing local AI...')
                : 'Create runs entirely on-device. Install the local AI model to get started.'}
        </div>

        {aiProgressPercent !== null && (
          <div className="ai-gate-progress-section">
            <div className="ai-gate-progress-label">
              <span>{aiSetupInProgress ? 'Installing' : 'Working'}</span>
              <span>{aiProgressPercent}%</span>
            </div>
            <div className="ai-gate-progress">
              <div className="ai-gate-progress-bar" style={{ width: `${aiProgressPercent}%` }} />
            </div>
          </div>
        )}

        {aiChecking && aiProgressPercent === null && (
          <div className="ai-gate-progress">
            <div className="ai-gate-progress-bar ai-gate-progress-indeterminate" />
          </div>
        )}

        {aiStatus?.lastError && (
          <div className="ai-gate-error">{aiStatus.lastError}</div>
        )}

        {!aiOperationInProgress && (
          <button
            type="button"
            className="ai-gate-download-btn"
            onClick={() => { void handleLocalAiSetup(); }}
          >
            Install Local AI
          </button>
        )}

        <div className="ai-gate-badges">
          <span className="ai-gate-badge">
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="3" y="11" width="18" height="11" rx="2" /><path d="M7 11V7a5 5 0 0110 0v4" /></svg>
            Private
          </span>
          <span className="ai-gate-badge">
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="2" y="2" width="20" height="20" rx="2" /><path d="M7 12h10M12 7v10" /></svg>
            On-device
          </span>
        </div>
      </div>
    </div>
  ) : null;


  const conversationPanel = (
    <div className="conversation-panel">
      {messages.map((msg, i) => (
        <div
          key={i}
          className={msg.role === 'user' ? 'user-prompt' : 'assistant-message'}
        >
          {msg.content}
        </div>
      ))}

      {isBusy && (
        <div className="ai-loading-indicator">
          <span className="ai-loading-dot" />
          <span className="ai-loading-dot" />
          <span className="ai-loading-dot" />
        </div>
      )}

      {errorInfo?.restartRequired && (
        <div className="completion-message">
          Close MacroFlow and reopen it to load the workbook-scoped sync APIs.
        </div>
      )}
      <div ref={conversationEndRef} />
    </div>
  );

  const searchStatus = searchData?.status || 'idle';
  const searchNotReady = searchStatus !== 'ready';

  const sidebarContent = showCodePanel ? (
    <>
      <div className="build-chat-header">
        <div className="build-chat-header-left">
          <button
            type="button"
            className="build-chat-back-btn"
            onClick={handleBackToHistory}
            disabled={isBusy}
            title="Back to sessions"
          >
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <polyline points="15,18 9,12 15,6" />
            </svg>
          </button>
          <span className="build-chat-module-name">{displayMacroName(sessionMacroName || location.module)}</span>
        </div>
      </div>
      {conversationPanel}
      <div className="build-prompt-input-wrap">
        <div className="build-prompt-input-box">
          <textarea
            className="build-prompt-input"
            placeholder="Ask for follow-up changes"
            rows={2}
            value={prompt}
            onChange={(event) => setPrompt(event.target.value)}
            disabled={isBusy || buildState === 'initializing' || Boolean(errorInfo?.restartRequired) || !aiReady}
            onKeyDown={(event) => {
              if (event.key === 'Enter' && !event.shiftKey && prompt.trim()) {
                event.preventDefault();
                void dispatchSubmit();
              }
            }}
          />
          <button
            type="button"
            className="build-prompt-send-btn"
            onClick={() => { if (prompt.trim()) void dispatchSubmit(); }}
            disabled={isBusy || !prompt.trim() || buildState === 'initializing' || !aiReady}
            title="Send"
          >
            <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
              <line x1="12" y1="19" x2="12" y2="5" />
              <polyline points="5,12 12,5 19,12" />
            </svg>
          </button>
        </div>
      </div>
    </>
  ) : (
    <div className="build-sidebar-history">
      <div className="build-chat-header">
        <div className="build-chat-header-left">
          <span className="build-chat-module-name">Sessions</span>
        </div>
      </div>
      <div className="build-session-history-list">
        {sessions.length > 0 ? (
          sessions
            .filter((s) => !normalizedWorkbook?.key || s.sessionContext?.workbook?.key === normalizedWorkbook.key)
            .slice().reverse().map((s) => (
            <div key={s.id} style={{ position: 'relative' }}>
              <button
                className="build-session-history-item"
                onClick={() => { setSessionContextMenu(null); handleRestoreSession(s.id); }}
                onContextMenu={(e) => {
                  e.preventDefault();
                  setSessionContextMenu((prev) => prev === s.id ? null : s.id);
                }}
              >
                <span className="build-session-history-module">{displayMacroName(extractPrimaryMacroName(s.editedCode) || s.sessionContext.moduleName)}</span>
                <span className="build-session-history-time">
                  {new Date(s.savedAt).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                </span>
              </button>
              {sessionContextMenu === s.id && (
                <button
                  className="session-context-menu-delete"
                  onClick={() => { setSessionContextMenu(null); removeSession(s.id); }}
                  onBlur={() => setSessionContextMenu(null)}
                  autoFocus
                >
                  Delete
                </button>
              )}
            </div>
          ))
        ) : (
          <div className="build-session-history-empty">No previous sessions</div>
        )}
      </div>
    </div>
  );

  const mainContent = searchNotReady ? (() => {
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
  })() : !showCodePanel ? (
    buildState === 'idle' || (buildState === 'initializing' && !strictLaunchMode) ? (
      <div className="build-empty-state build-idle-state">
        <h1 className="build-empty-title">Create a Macro</h1>
        <p className="build-empty-subtitle">
          Describe what you want your macro to do and MacroFlow will generate the VBA locally.
        </p>
        {errorInfo && <div className="error-title">{errorInfo.title}</div>}
        <div className="build-prompt-input-wrap">
            <div className="build-prompt-input-box">
              <div className="build-prompt-context-bar">
                <div className="build-prompt-workbook-picker">
                  <button
                    type="button"
                    className="build-prompt-workbook-btn"
                    onClick={async () => {
                      if (createWorkbookOpen) {
                        setCreateWorkbookOpen(false);
                        return;
                      }
                      try {
                        const listApi = window.excel?.workbook?.list || window.excel?.workbook?.listContext;
                        if (typeof listApi === 'function') {
                          const result = await listApi();
                          if (result?.success && Array.isArray(result.workbooks)) {
                            setCreateWorkbookList(result.workbooks);
                          }
                        }
                      } catch { /* ignore */ }
                      setCreateWorkbookOpen(true);
                    }}
                    disabled={isBusy}
                  >
                    <WorkbookIcon size={11} />
                    <span>{createWorkbook?.name || PERSONAL_WORKBOOK_NAME}</span>
                    <svg width="8" height="8" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                      <polyline points="6,9 12,15 18,9" />
                    </svg>
                  </button>
                  {createWorkbookOpen && (
                    <>
                      <div className="build-prompt-workbook-backdrop" onClick={() => setCreateWorkbookOpen(false)} />
                      <div className="build-prompt-workbook-dropdown">
                        {[{ name: PERSONAL_WORKBOOK_NAME, path: '' }, ...createWorkbookList.filter(
                          (wb) => wb.name?.toUpperCase() !== PERSONAL_WORKBOOK_NAME
                        )].map((wb) => (
                          <button
                            key={wb.path || wb.name}
                            type="button"
                            className={`build-prompt-workbook-option${(createWorkbook?.name || '') === wb.name ? ' active' : ''}`}
                            onClick={() => {
                              setCreateWorkbook({ name: wb.name, path: wb.path || '' });
                              setLocation(buildLocation(wb.name, DEFAULT_MODULE_LABEL));
                              setCreateWorkbookOpen(false);
                            }}
                          >
                            <WorkbookIcon size={11} />
                            <span>{wb.name}</span>
                          </button>
                        ))}
                      </div>
                    </>
                  )}
                </div>
              </div>
              <textarea
                className="build-prompt-input"
                placeholder="Describe the macro you want"
                rows={2}
                value={prompt}
                onChange={(event) => setPrompt(event.target.value)}
                disabled={isBusy}
                onKeyDown={(event) => {
                  if (event.key === 'Enter' && !event.shiftKey && prompt.trim()) {
                    event.preventDefault();
                    void dispatchSubmit();
                  }
                }}
              />
              <div className="build-prompt-input-footer">
                <div />
                <button
                  type="button"
                  className="build-prompt-send-btn"
                  onClick={() => { if (prompt.trim()) void dispatchSubmit(); }}
                  disabled={isBusy || !prompt.trim()}
                  title="Send"
                >
                  <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                    <line x1="12" y1="19" x2="12" y2="5" />
                    <polyline points="5,12 12,5 19,12" />
                  </svg>
                </button>
              </div>
            </div>
          </div>
      </div>
    ) : (
      <div className="build-empty-state">
        {buildState === 'initializing' ? (
          <span className="status-spinner" />
        ) : (
          <>
            <h1 className="build-empty-title">Build Session Unavailable</h1>
            {errorInfo && <div className="error-title">{errorInfo.title}</div>}
          </>
        )}
      </div>
    )
  ) : (
    <div className="build-code-fullwidth">
      <div className="build-code-bar">
        {currentMacroId && (
          <button
            type="button"
            className="build-chat-icon-btn"
            title="Change icon"
            onClick={() => setShowIconPicker(true)}
          >
            {isSpriteReady() ? (
              <ImageMsoIcon name={currentMacroIcon || 'FileSaveAs'} size={22} title="" />
            ) : (
              <ReturnIcon size={22} />
            )}
          </button>
        )}
        {buildState !== 'initializing' && (
          <button
            type="button"
            className="build-run-action"
            onClick={() => {
              if (!isBusy) {
                if (hasPendingChanges) {
                  void handleSave();
                  return;
                }
                void handleRunMacro();
              }
            }}
            disabled={isBusy || (!hasPendingChanges && !sessionMacroName)}
          >
            {primaryActionLabel}
          </button>
        )}
      </div>
      {showIconPicker && currentMacroId && (
        <IconPicker
          currentIcon={currentMacroIcon}
          onSelect={(iconName) => {
            setMacroIcon(currentMacroId, iconName);
            setIconVersion(v => v + 1);
          }}
          onClose={() => setShowIconPicker(false)}
        />
      )}
      <CodePreview
        code={currentCode}
        showHeader={false}
        editable={true}
        onChange={handleCodeChange}
        status={codeStatus}
        errorLine={errorInfo?.line}
      />
    </div>
  );

  return (
    <>
      <main className="main-content">
        <div className={!aiReady ? 'ai-gate-bg-blur' : undefined} style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
          <div className="split-view">
            <div className="split-left build-chat-panel" style={{ width: chatOpen ? `${splitPct}%` : '0%', display: chatOpen ? undefined : 'none' }}>
              {sidebarContent}
            </div>
            {chatOpen && <SplitDivider onResize={setSplitPct} />}
            <div className={chatOpen ? 'split-right' : 'split-right split-right--full'}>
              {mainContent}
            </div>
          </div>
        </div>
        {aiGateOverlay}
      </main>

      {exitDialog && exitDialog.type === 'confirm' && (
        <div className="build-exit-overlay" role="dialog" aria-modal="true" aria-label="Unsaved changes">
          <div className="build-exit-dialog">
            <div className="build-exit-title">Unsaved changes</div>
            <div className="build-exit-message">Do you want to save your changes before leaving?</div>
            <div className="build-exit-actions">
              <button
                type="button"
                className="build-exit-btn primary"
                onClick={() => void confirmAndSaveExit(exitDialog.intent)}
              >
                Save
              </button>
              <button
                type="button"
                className="build-exit-btn danger"
                onClick={() => {
                  cleanupEmptyModule();
                  setDirtyState(false);
                  setExitDialog(null);
                  if (!runPendingInternalAction()) {
                    finalizeExit(exitDialog.intent);
                  }
                }}
              >
                Discard
              </button>
              <button
                type="button"
                className="build-exit-btn"
                onClick={() => {
                  pendingInternalActionRef.current = null;
                  setExitDialog(null);
                }}
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {exitDialog && exitDialog.type === 'error' && (
        <div className="build-exit-overlay" role="dialog" aria-modal="true" aria-label="Exit build mode save conflict">
          <div className="build-exit-dialog">
            <div className="build-exit-title">Unable to save</div>
            <div className="build-exit-message">{exitDialog.message}</div>
            <div className="build-exit-actions">
              <button
                type="button"
                className="build-exit-btn primary"
                onClick={() => void confirmAndSaveExit(exitDialog.intent)}
              >
                Retry
              </button>
              <button
                type="button"
                className="build-exit-btn danger"
                onClick={() => {
                  cleanupEmptyModule();
                  setDirtyState(false);
                  setExitDialog(null);
                  if (!runPendingInternalAction()) {
                    finalizeExit(exitDialog.intent);
                  }
                }}
              >
                Exit without save
              </button>
              <button
                type="button"
                className="build-exit-btn"
                onClick={() => {
                  pendingInternalActionRef.current = null;
                  setExitDialog(null);
                }}
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
};

export default CreatePage;
