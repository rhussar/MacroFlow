import React, { useState, useEffect, useRef, useCallback, useMemo } from 'react';
import { ArrowLeftIcon, CloseIcon, CheckIcon, MacroFlowLogo, DocumentIcon, SidebarIcon } from './icons';
import CodePreview from './CodePreview';
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

// Placeholder generated VBA code until AI codegen is wired in.
const mockGeneratedCode = `Sub CleanData()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' 1. Remove Empty Rows
    On Error Resume Next
    ws.Columns("A:A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete

    ' 2. Trim Whitespace
    For Each cell In ws.Range("B1:B150")
        cell.Value = Trim(cell.Value)
    Next cell

    ' 3. Fix Date Format
    ws.Columns("C:C").NumberFormat = "mm/dd/yyyy"

    MsgBox "Cleanup Complete!"
End Sub`;

const DEFAULT_MODULE_LABEL = 'New Module';
const MODULE_PREFIX = 'MacroFlowModule';
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

const BuildMode = ({
  onBack,
  onClose,
  targetWorkbook,
  launchMode = 'new_module',
  launchModuleName = '',
  launchSource = '',
  onRefreshSearchData
}) => {
  const [prompt, setPrompt] = useState('');
  const [buildState, setBuildState] = useState('initializing');
  const [steps, setSteps] = useState(() => createInitializationSteps({ resolve: 'loading' }));
  const [errorInfo, setErrorInfo] = useState(null);
  const [chatOpen, setChatOpen] = useState(true);
  const [editedCode, setEditedCode] = useState(BUILD_MODE_SEED_CODE);
  const [isBusy, setIsBusy] = useState(false);
  const [location, setLocation] = useState(buildLocation(targetWorkbook?.name, DEFAULT_MODULE_LABEL));
  const [sessionContext, setSessionContext] = useState(null);
  const [runOutcome, setRunOutcome] = useState('idle');
  const [savedMacroName, setSavedMacroName] = useState('');
  const [lastPrompt, setLastPrompt] = useState('');
  const [exitDialog, setExitDialog] = useState(null);
  const [savedShortcutLetter, setSavedShortcutLetter] = useState('');
  const [draftShortcutLetter, setDraftShortcutLetter] = useState('');
  const [shortcutInputError, setShortcutInputError] = useState(false);
  const [shortcutSaving, setShortcutSaving] = useState(false);

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
  const editedCodeRef = useRef(editedCode);
  const isBusyRef = useRef(isBusy);
  const sessionContextRef = useRef(sessionContext);
  const savedShortcutLetterRef = useRef(savedShortcutLetter);
  const draftShortcutLetterRef = useRef(draftShortcutLetter);
  const shortcutSavingRef = useRef(shortcutSaving);
  const bootstrapSessionRef = useRef(() => Promise.resolve());
  const attemptExitRef = useRef(() => Promise.resolve());
  const activeBootstrapIdRef = useRef(0);
  const activeShortcutHydrationIdRef = useRef(0);
  const dirtyLocalRef = useRef(false);
  const exitInFlightRef = useRef(false);

  buildStateRef.current = buildState;
  promptRef.current = prompt;
  chatOpenRef.current = chatOpen;
  editedCodeRef.current = editedCode;
  isBusyRef.current = isBusy;
  sessionContextRef.current = sessionContext;
  savedShortcutLetterRef.current = savedShortcutLetter;
  draftShortcutLetterRef.current = draftShortcutLetter;
  shortcutSavingRef.current = shortcutSaving;

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

  const finalizeExit = useCallback((intent) => {
    const normalizedIntent = normalizeExitIntent(intent);
    if (normalizedIntent === 'close') {
      onClose?.();
      return;
    }
    onBack?.();
  }, [onBack, onClose]);

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

    dirtyLocalRef.current = false;

    if (typeof onRefreshSearchData === 'function') {
      await onRefreshSearchData({ silent: true });
    }
  }, [onRefreshSearchData]);

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

    exitInFlightRef.current = true;
    setIsBusy(true);

    try {
      await persistUnsyncedChanges();
      setExitDialog(null);
      finalizeExit(normalizedIntent);
    } catch (error) {
      const message = String(error?.message || 'Unable to save changes before exit.');
      setExitDialog({
        intent: normalizedIntent,
        message
      });
    } finally {
      exitInFlightRef.current = false;
      setIsBusy(false);
    }
  }, [finalizeExit, persistUnsyncedChanges]);

  attemptExitRef.current = attemptExit;

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
    setLastPrompt('');
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

    dirtyLocalRef.current = false;

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

      if (strictWorkbookMode) {
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

      const resolvedWorkbook = strictWorkbookMode
        ? normalizeBuildWorkbook(selectedResolution?.workbook || selectedCandidate)
        : resolveBuildWorkbookTarget({
            selectedWorkbook: selectedResolution?.workbook,
            selectedWorkbookFound: Boolean(selectedResolution),
            activeWorkbook: activeResolution?.workbook
          });
      if (!resolvedWorkbook) {
        throw new Error('No valid workbook is available for Build Mode.');
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

      if (strictWorkbookMode) {
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

      if (typeof onRefreshSearchData === 'function') {
        await onRefreshSearchData({ silent: true });
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
    normalizedLaunchMode,
    normalizedWorkbook,
    onRefreshSearchData,
    requestedLaunchModuleName,
    setStepStatus,
    stopWithError
  ]);

  bootstrapSessionRef.current = bootstrapSession;

  const markLocalDirty = useCallback(() => {
    dirtyLocalRef.current = true;
    setRunOutcome('idle');
  }, []);

  const handleCodeChange = useCallback((newCode) => {
    const nextCode = String(newCode || '');
    setEditedCode(nextCode);
    setShortcutInputError(false);
    markLocalDirty();
  }, [markLocalDirty]);

  const handleSave = useCallback(() => {
    if (!dirtyLocalRef.current) {
      return;
    }
  }, []);

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
    } catch (_error) {
      setShortcutInputError(true);
    } finally {
      shortcutSavingRef.current = false;
      setShortcutSaving(false);
    }
  }, []);

  const toggleChat = useCallback(() => setChatOpen((previous) => !previous), []);

  const handleSubmit = useCallback(() => {
    const submittedPrompt = String(promptRef.current || '').trim();
    if (isBusyRef.current || !submittedPrompt || !sessionContextRef.current) {
      return;
    }

    setLastPrompt(submittedPrompt);
    setPrompt('');
    setRunOutcome('idle');
    setErrorInfo(null);

    const generatedCode = String(mockGeneratedCode);
    setEditedCode(generatedCode);
    setSavedMacroName(extractPrimaryMacroName(generatedCode));
    markLocalDirty();
  }, [markLocalDirty]);

  const handleRunMacro = useCallback(async () => {
    if (isBusyRef.current) {
      return;
    }

    const runApi = window.excel?.vba?.run;
    if (typeof runApi !== 'function') {
      setRunOutcome('error');
      setErrorInfo({
        title: 'Run failed: VBA run API is unavailable.',
        line: null
      });
      return;
    }

    const session = sessionContextRef.current;
    if (!session) {
      setRunOutcome('error');
      setErrorInfo({
        title: 'Run failed: Build session is not ready.',
        line: null
      });
      return;
    }

    const moduleName = String(session?.moduleName || '').trim();
    const workbookName = String(session?.workbook?.name || '').trim();
    const codeText = String(editedCodeRef.current || '');
    const macroName = String(extractPrimaryMacroName(codeText) || savedMacroName).trim();

    if (!moduleName || !macroName) {
      setRunOutcome('error');
      setErrorInfo({
        title: 'Run failed: Add a Sub procedure before running.',
        line: null
      });
      return;
    }

    const runTarget = buildWorkbookQualifiedRunTarget(workbookName, moduleName, macroName);
    if (!runTarget) {
      setRunOutcome('error');
      setErrorInfo({
        title: 'Run failed: Unable to resolve macro target.',
        line: null
      });
      return;
    }

    setErrorInfo(null);
    setIsBusy(true);
    isBusyRef.current = true;

    try {
      const result = await runApi({ macroName: runTarget });
      if (!result?.success) {
        throw new Error(result?.message || 'Macro execution failed.');
      }
      setRunOutcome('success');
    } catch (error) {
      const message = error?.message ? String(error.message) : 'Macro execution failed.';
      setRunOutcome('error');
      setErrorInfo({
        title: `Run failed: ${message}`,
        line: null
      });
    } finally {
      setIsBusy(false);
      isBusyRef.current = false;
    }
  }, [savedMacroName]);

  const sessionMacroName = useMemo(() => extractPrimaryMacroName(editedCode), [editedCode]);
  const sessionMacroTarget = useMemo(
    () => buildSessionMacroTarget(sessionContext?.moduleName, sessionMacroName),
    [sessionContext?.moduleName, sessionMacroName]
  );
  const shortcutPrefix = formatShortcutPrefix(draftShortcutLetter);

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
    const bootstrapTimerId = window.setTimeout(() => {
      void bootstrapSessionRef.current();
    }, 0);

    return () => {
      window.clearTimeout(bootstrapTimerId);
      activeBootstrapIdRef.current += 1;
    };
  }, [
    normalizedLaunchMode,
    normalizedWorkbook?.key,
    normalizedWorkbook?.name,
    normalizedWorkbook?.path,
    requestedLaunchModuleName
  ]);

  // Keyboard shortcut handler uses refs so the listener is registered once.
  useEffect(() => {
    const handleKeyDown = (event) => {
      if (event.ctrlKey && event.key === 's' && !chatOpenRef.current && sessionContextRef.current) {
        event.preventDefault();
        handleSave();
        return;
      }

      if (event.key === 'Escape') {
        event.preventDefault();
        void attemptExitRef.current('back');
        return;
      }

      if (!event.ctrlKey && !event.altKey && !event.metaKey && !event.shiftKey && event.key === 'Tab') {
        event.preventDefault();
        void attemptExitRef.current('back');
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
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
  const codeStatus =
    runOutcome === 'success'
      ? 'success'
      : runOutcome === 'error' || buildState === 'error'
        ? 'error'
        : 'normal';
  const currentCode = String(editedCode || BUILD_MODE_SEED_CODE);

  const conversationPanel = (
    <div className="conversation-panel">
      <div className="user-prompt">
        {lastPrompt || 'Build session is ready. Describe what macro you want to generate.'}
      </div>

      {errorInfo && (
        <div className="error-title">{errorInfo.title}</div>
      )}

      <div className="status-steps">
        {steps.map((step) => (
          <div
            key={step.key}
            className={`status-step ${step.status === 'complete' ? 'completed' : ''}`}
          >
            {step.status === 'complete' ? (
              <span className="status-checkbox checked">
                <CheckIcon size={12} />
              </span>
            ) : step.status === 'loading' ? (
              <span className="status-spinner" />
            ) : (
              <span className="status-checkbox" />
            )}
            <span>{step.text}</span>
          </div>
        ))}
      </div>

      {errorInfo?.restartRequired && (
        <div className="completion-message">
          Close MacroFlow and reopen it to load the workbook-scoped sync APIs.
        </div>
      )}
    </div>
  );

  return (
    <>
      <header className="header">
        <div className="drag-region" />
        <button
          className="header-back-btn"
          onClick={() => {
            void attemptExitRef.current('back');
          }}
        >
          <ArrowLeftIcon size={20} />
        </button>

        {showCodePanel && (
          <button className="chat-toggle-icon-btn" onClick={toggleChat} title={chatOpen ? 'Close Chat' : 'Open Chat'}>
            <SidebarIcon size={18} />
          </button>
        )}

        <div className="search-input-wrapper">
          <input
            type="text"
            className="search-input"
            placeholder={
              showCodePanel
                ? 'Describe the macro change...'
                : buildState === 'initializing'
                  ? 'Preparing build session...'
                  : 'Build session unavailable'
            }
            value={prompt}
            onChange={(event) => setPrompt(event.target.value)}
            disabled={!showCodePanel || buildState === 'initializing' || Boolean(errorInfo?.restartRequired)}
            onKeyDown={(event) => {
              if (event.key === 'Enter' && showCodePanel && prompt.trim()) {
                event.preventDefault();
                handleSubmit();
              }
            }}
          />
        </div>

        <div className="header-actions">
          <button
            className="close-btn"
            onClick={() => {
              void attemptExitRef.current('close');
            }}
          >
            <CloseIcon />
          </button>
        </div>
      </header>

      <main className="main-content">
        {!showCodePanel ? (
          <div className="build-empty-state">
            <h1 className="build-empty-title">
              {buildState === 'initializing' ? 'Preparing Build Session' : 'Build Session Unavailable'}
            </h1>
            <p className="build-empty-subtitle">
              {buildState === 'initializing'
                ? strictLaunchMode
                  ? 'Locating the selected module and loading entry sync...'
                  : 'Creating a workbook-bound module and loading entry sync...'
                : 'Fix the error and retry build session initialization.'}
            </p>
            <div className="conversation-panel">
              {errorInfo && <div className="error-title">{errorInfo.title}</div>}
              <div className="status-steps">
                {steps.map((step) => (
                  <div
                    key={step.key}
                    className={`status-step ${step.status === 'complete' ? 'completed' : ''}`}
                  >
                    {step.status === 'complete' ? (
                      <span className="status-checkbox checked">
                        <CheckIcon size={12} />
                      </span>
                    ) : step.status === 'loading' ? (
                      <span className="status-spinner" />
                    ) : (
                      <span className="status-checkbox" />
                    )}
                    <span>{step.text}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>
        ) : chatOpen ? (
          <div className="split-view">
            <div className="split-left build-chat-panel">
              {conversationPanel}
            </div>
            <div className="split-right">
              <div className="build-code-bar">
                <div className="code-breadcrumb">
                  <DocumentIcon size={14} className="code-breadcrumb-icon" />
                  <span className="code-breadcrumb-text">
                    {location.workbook}
                    <span className="code-breadcrumb-separator"> &gt; </span>
                    {location.module}
                  </span>
                </div>
              </div>
              <CodePreview
                code={currentCode}
                showHeader={false}
                editable={true}
                onChange={handleCodeChange}
                status={codeStatus}
                errorLine={errorInfo?.line}
              />
            </div>
          </div>
        ) : (
          <div className="build-code-fullwidth">
            <div className="build-code-bar">
              <div className="code-breadcrumb">
                <DocumentIcon size={14} className="code-breadcrumb-icon" />
                <span className="code-breadcrumb-text">
                  {location.workbook}
                  <span className="code-breadcrumb-separator"> &gt; </span>
                  {location.module}
                </span>
              </div>
            </div>
            <CodePreview
              code={currentCode}
              showHeader={false}
              editable={true}
              onChange={handleCodeChange}
              status={codeStatus}
              errorLine={errorInfo?.line}
            />
          </div>
        )}
      </main>

      {showCodePanel ? (
        <footer className="footer">
          <div className="footer-left">
            <div className={`logo ${footerContent.logoClass || ''}`}>
              <MacroFlowLogo size={20} />
            </div>
            {footerContent.text && (
              <span className={`footer-text ${footerContent.textClass || ''}`}>
                {footerContent.text}
              </span>
            )}
          </div>
          <div className="footer-right build-footer-actions">
            <button
              type="button"
              className="build-run-action"
              onClick={() => {
                if (!isBusy) {
                  void handleRunMacro();
                }
              }}
              disabled={isBusy}
            >
              Run macro
            </button>
            <div
              className={`shortcut-binding build-shortcut-binding ${shortcutSaving ? 'saving' : ''}`}
              onClick={(event) => event.stopPropagation()}
            >
              <span className="shortcut-prefix">Ctrl +</span>
              {shortcutPrefix.includes('Shift') && (
                <span className="shortcut-shift">Shift +</span>
              )}
              <input
                type="text"
                className={`shortcut-keycap-input ${draftShortcutLetter ? '' : 'is-empty'} ${shortcutInputError ? 'has-error' : ''}`}
                value={draftShortcutLetter}
                placeholder=""
                maxLength={1}
                autoCapitalize="off"
                autoComplete="off"
                spellCheck={false}
                aria-label="Excel shortcut letter for the current Build Mode macro"
                onChange={(event) => handleShortcutDraftChange(event.target.value)}
                onBlur={() => {
                  void commitShortcutDraft();
                }}
                onKeyDown={(event) => {
                  if (event.key === 'Enter') {
                    event.preventDefault();
                    event.stopPropagation();
                    event.currentTarget.blur();
                  } else if (event.key === 'Escape') {
                    event.preventDefault();
                    event.stopPropagation();
                    setDraftShortcutLetter(savedShortcutLetterRef.current || '');
                    setShortcutInputError(false);
                    event.currentTarget.blur();
                  }
                }}
                onClick={(event) => event.stopPropagation()}
                disabled={shortcutInputDisabled}
              />
            </div>
          </div>
        </footer>
      ) : (
        <footer className="footer">
          <div className="footer-left">
            <div className="logo">
              <MacroFlowLogo size={20} />
            </div>
          </div>
        </footer>
      )}

      {exitDialog && (
        <div className="build-exit-overlay" role="dialog" aria-modal="true" aria-label="Exit build mode save conflict">
          <div className="build-exit-dialog">
            <div className="build-exit-title">Unable to save before exit</div>
            <div className="build-exit-message">{exitDialog.message}</div>
            <div className="build-exit-actions">
              <button
                type="button"
                className="build-exit-btn primary"
                onClick={() => {
                  const action = resolveBuildExitAction('retry');
                  if (action === 'retry') {
                    void attemptExitRef.current(exitDialog.intent);
                  }
                }}
              >
                Retry
              </button>
              <button
                type="button"
                className="build-exit-btn danger"
                onClick={() => {
                  const action = resolveBuildExitAction('exit_without_save');
                  if (action === 'exit_without_save') {
                    dirtyLocalRef.current = false;
                    setExitDialog(null);
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
                  const action = resolveBuildExitAction('cancel');
                  if (action === 'cancel') {
                    setExitDialog(null);
                  }
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

export default BuildMode;
