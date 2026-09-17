import { useState, useMemo, useCallback } from 'react';
import {
  extractPrimaryMacroName,
  buildWorkbookQualifiedRunTarget,
  buildSessionMacroTarget
} from './build-target';

export function useMacroExecution({
  sessionContextRef,
  editedCodeRef,
  editedCode,
  sessionContext,
  isBusyRef,
  setErrorInfo,
  setIsBusy
}) {
  const [runOutcome, setRunOutcome] = useState('idle');
  const [savedMacroName, setSavedMacroName] = useState('');

  const sessionMacroName = useMemo(() => extractPrimaryMacroName(editedCode), [editedCode]);
  const sessionMacroTarget = useMemo(
    () => buildSessionMacroTarget(sessionContext?.moduleName, sessionMacroName),
    [sessionContext?.moduleName, sessionMacroName]
  );

  const handleRunMacro = useCallback(async () => {
    if (isBusyRef.current) return;

    const runApi = window.excel?.vba?.run;
    if (typeof runApi !== 'function') {
      setRunOutcome('error');
      setErrorInfo({ title: 'Run failed: VBA run API is unavailable.', line: null });
      return;
    }

    const session = sessionContextRef.current;
    if (!session) {
      setRunOutcome('error');
      setErrorInfo({ title: 'Run failed: Build session is not ready.', line: null });
      return;
    }

    const moduleName = String(session?.moduleName || '').trim();
    const workbookName = String(session?.workbook?.name || '').trim();
    const codeText = String(editedCodeRef.current || '');
    const macroName = String(extractPrimaryMacroName(codeText) || savedMacroName).trim();

    if (!moduleName || !macroName) {
      setRunOutcome('error');
      setErrorInfo({ title: 'Run failed: Add a Sub procedure before running.', line: null });
      return;
    }

    const runTarget = buildWorkbookQualifiedRunTarget(workbookName, moduleName, macroName);
    if (!runTarget) {
      setRunOutcome('error');
      setErrorInfo({ title: 'Run failed: Unable to resolve macro target.', line: null });
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
      setErrorInfo({ title: `Run failed: ${message}`, line: null });
    } finally {
      setIsBusy(false);
      isBusyRef.current = false;
    }
  }, [savedMacroName, sessionContextRef, editedCodeRef, isBusyRef, setErrorInfo, setIsBusy]);

  return {
    runOutcome, setRunOutcome,
    savedMacroName, setSavedMacroName,
    sessionMacroName, sessionMacroTarget,
    handleRunMacro
  };
}
