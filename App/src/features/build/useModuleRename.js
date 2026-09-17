import { useState, useRef, useCallback } from 'react';
import {
  normalizeBuildWorkbook
} from './build-target';
import {
  isValidVbaModuleName,
  shouldCommitModuleRename,
  encodeMacroName
} from '../search/module-actions';

export function useModuleRename({
  sessionContextRef,
  location,
  buildLocationFn,
  invalidateWorkbookMutation,
  setSessionContext,
  setLocation,
  setErrorInfo,
  setIsBusy
}) {
  const [moduleRenameActive, setModuleRenameActive] = useState(false);
  const [moduleRenameDraft, setModuleRenameDraft] = useState('');
  const moduleRenameInputRef = useRef(null);
  const moduleRenameRestoreValueRef = useRef('');
  const moduleRenameCommitInFlightRef = useRef(false);

  const cancelModuleRename = useCallback(() => {
    const restoreName = String(
      moduleRenameRestoreValueRef.current || sessionContextRef.current?.moduleName || location.module || ''
    ).trim();
    setModuleRenameDraft(restoreName);
    setModuleRenameActive(false);
  }, [location.module, sessionContextRef]);

  const commitModuleRename = useCallback(async (nextModuleNameOverride = null) => {
    if (moduleRenameCommitInFlightRef.current) return;

    const renameApi = window.excel?.vba?.renameModuleByWorkbook;
    if (typeof renameApi !== 'function') {
      setErrorInfo({ title: 'Rename failed: workbook module rename API is unavailable.', line: null });
      setModuleRenameActive(false);
      return;
    }

    const session = sessionContextRef.current;
    if (!session) {
      setModuleRenameActive(false);
      return;
    }

    const currentModuleName = String(session.moduleName || '').trim();
    const nextModuleName = encodeMacroName(String(
      nextModuleNameOverride ?? moduleRenameInputRef.current?.textContent ?? moduleRenameDraft
    ).trim());

    if (!shouldCommitModuleRename({ currentName: currentModuleName, nextName: nextModuleName })) {
      setModuleRenameDraft(currentModuleName);
      setModuleRenameActive(false);
      return;
    }

    if (!isValidVbaModuleName(nextModuleName)) {
      setErrorInfo({ title: 'Rename failed: module names must start with a letter and use only letters, numbers, or underscores.', line: null });
      return;
    }

    moduleRenameCommitInFlightRef.current = true;
    setIsBusy(true);

    try {
      const result = await renameApi({
        workbookName: session.workbook.name,
        workbookPath: session.workbook.path,
        moduleName: currentModuleName,
        nextModuleName
      });

      if (!result?.success || result?.workbookFound === false || result?.moduleFound === false || result?.renamed === false) {
        throw new Error(result?.message || 'Unable to rename module.');
      }

      const updatedWorkbook = normalizeBuildWorkbook(result?.workbook || session.workbook) || session.workbook;
      const updatedModuleName = String(result?.moduleName || nextModuleName).trim() || nextModuleName;
      const updatedSession = { ...session, workbook: updatedWorkbook, moduleName: updatedModuleName };

      sessionContextRef.current = updatedSession;
      setSessionContext(updatedSession);
      setLocation(buildLocationFn(updatedWorkbook.name, updatedModuleName));
      setModuleRenameDraft(updatedModuleName);
      setModuleRenameActive(false);
      setErrorInfo(null);
      invalidateWorkbookMutation(updatedWorkbook, {
        includeActiveWorkbook: true,
        includeExplorerAllFiles: true,
        includePersonalMacros: true,
        includeShortcutAudit: true,
        includeWorkbookScopedData: true
      });
    } catch (error) {
      setErrorInfo({ title: `Rename failed: ${String(error?.message || 'Unable to rename module.')}`, line: null });
    } finally {
      moduleRenameCommitInFlightRef.current = false;
      setIsBusy(false);
    }
  }, [invalidateWorkbookMutation, moduleRenameDraft, sessionContextRef, buildLocationFn, setSessionContext, setLocation, setErrorInfo, setIsBusy]);

  return {
    moduleRenameActive, setModuleRenameActive,
    moduleRenameDraft, setModuleRenameDraft,
    moduleRenameInputRef, moduleRenameRestoreValueRef,
    cancelModuleRename, commitModuleRename
  };
}
