import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  getSearchStatusView,
  selectActiveWorkbookMacros,
  selectPersonalGlobalMacros,
  selectPersonalGlobalSectionModel
} from '../features/search/search-selectors';
import { usePersonalMacros, PERSONAL_WORKBOOK_NAME } from '../features/search/usePersonalMacros';
import { useWorkbookPickerData } from '../features/search/useWorkbookPickerData';
import { useShortcutState } from '../features/shortcuts/useShortcutState';
import PersonalMacrosSection from './shortcuts/PersonalMacrosSection';
import ShortcutGrid from './shortcuts/ShortcutGrid';
import WorkbookPicker from './shortcuts/WorkbookPicker';

const defaultSearchData = {
  status: 'idle',
  workbook: null,
  modules: [],
  macros: [],
  shortcutAudit: null,
  error: null
};

const ShortcutsPage = ({
  onBuildModeClick,
  onRunMacro,
  searchData = defaultSearchData,
  selectedMacroId = null,
  shortcutByMacroId = {},
  shortcutDraftByMacroId = {},
  shortcutInputErrorByMacroId = {},
  shortcutSavingMacroId = null,
  onShortcutDraftChange,
  onShortcutCommit,
  onActionStatus,
  shortcutSaveInFlightRef,
  selectedWorkbookForBuild = null,
  onSelectedWorkbookForBuildChange
}) => {
  const status = searchData.status || 'idle';
  const workbookPickerState = useWorkbookPickerData(searchData, {
    preferredWorkbookKey: selectedWorkbookForBuild?.key || ''
  });
  const [isWorkbookMenuOpen, setWorkbookMenuOpen] = useState(false);
  const [personalActionInFlight, setPersonalActionInFlight] = useState(false);
  const [personalInfoHover, setPersonalInfoHover] = useState(false);
  const [personalInfoPinned, setPersonalInfoPinned] = useState(false);
  const workbookMenuRef = useRef(null);
  const personalInfoRef = useRef(null);
  const showPersonalInfo = personalInfoHover || personalInfoPinned;

  const selectedWorkbook = workbookPickerState.selectedWorkbook;
  const selectedWorkbookLabel = selectedWorkbook?.name || 'Active Workbook';
  const selectedWorkbookData = workbookPickerState.selectedWorkbookData;
  const selectedWorkbookErrorMessage = selectedWorkbookData.error?.message
    ? String(selectedWorkbookData.error.message)
    : 'Unable to load workbook data.';
  const usingActiveWorkbookShortcuts = workbookPickerState.isSelectedActiveWorkbook;

  const displayedWorkbookData = useMemo(() => {
    if (status !== 'ready') {
      return {
        status: 'idle',
        workbook: null,
        modules: [],
        macros: [],
        error: null
      };
    }

    if (usingActiveWorkbookShortcuts) {
      return {
        status: 'ready',
        workbook: searchData?.workbook || selectedWorkbook,
        modules: Array.isArray(searchData?.modules) ? searchData.modules : [],
        macros: Array.isArray(searchData?.macros) ? searchData.macros : [],
        error: null
      };
    }

    return selectedWorkbookData;
  }, [
    searchData?.macros,
    searchData?.modules,
    searchData?.workbook,
    selectedWorkbook,
    selectedWorkbookData,
    status,
    usingActiveWorkbookShortcuts
  ]);

  const workbookShortcutState = useShortcutState({
    scope: 'workbook',
    enabled: status === 'ready' && !usingActiveWorkbookShortcuts && Boolean(selectedWorkbook?.name),
    workbook: selectedWorkbook,
    macros: displayedWorkbookData.macros,
    setActionStatus: onActionStatus,
    shortcutSaveInFlightRef
  });

  const effectiveShortcutState = usingActiveWorkbookShortcuts
    ? {
        shortcutByMacroId,
        shortcutDraftByMacroId,
        shortcutInputErrorByMacroId,
        shortcutSavingMacroId,
        handleShortcutDraftChange: onShortcutDraftChange,
        handleShortcutCommit: onShortcutCommit
      }
    : workbookShortcutState;

  const activeMacroRows = useMemo(
    () => selectActiveWorkbookMacros(
      displayedWorkbookData.macros,
      '',
      effectiveShortcutState.shortcutByMacroId
    ),
    [displayedWorkbookData.macros, effectiveShortcutState.shortcutByMacroId]
  );

  const personalMacrosState = usePersonalMacros(searchData, workbookPickerState.workbookListSignature);
  const personalSectionVisible = String(selectedWorkbook?.name || '').trim().toUpperCase() !== PERSONAL_WORKBOOK_NAME;
  const personalShortcutState = useShortcutState({
    scope: 'workbook',
    enabled: status === 'ready'
      && personalSectionVisible
      && personalMacrosState.workbookFound
      && personalMacrosState.macros.length > 0,
    workbook: personalMacrosState.workbook || {
      name: PERSONAL_WORKBOOK_NAME,
      path: personalMacrosState.workbookPath
    },
    macros: personalMacrosState.macros,
    seededAudit: personalMacrosState.shortcutAudit,
    setActionStatus: onActionStatus,
    shortcutSaveInFlightRef
  });

  const personalMacroRows = useMemo(
    () => selectPersonalGlobalMacros(
      personalMacrosState.macros,
      '',
      personalShortcutState.shortcutByMacroId
    ),
    [personalMacrosState.macros, personalShortcutState.shortcutByMacroId]
  );

  const personalSectionModel = useMemo(
    () => selectPersonalGlobalSectionModel({
      workbookName: selectedWorkbook?.name,
      rows: personalMacroRows,
      totalMacros: personalMacrosState.macros.length,
      status: personalMacrosState.status,
      workbookFound: personalMacrosState.workbookFound,
      fileExists: personalMacrosState.fileExists,
      error: personalMacrosState.error
    }),
    [
      personalMacroRows,
      personalMacrosState.error,
      personalMacrosState.fileExists,
      personalMacrosState.macros.length,
      personalMacrosState.status,
      personalMacrosState.workbookFound,
      selectedWorkbook?.name
    ]
  );

  const handlePersonalAction = useCallback(async (action) => {
    if (!action || personalActionInFlight) {
      return;
    }

    const workbookPath = String(
      personalMacrosState.workbook?.path || personalMacrosState.workbookPath || ''
    ).trim();
    const workbookTarget = {
      name: PERSONAL_WORKBOOK_NAME,
      path: workbookPath,
      key: workbookPath || PERSONAL_WORKBOOK_NAME
    };

    if (action === 'create_global_macro') {
      onBuildModeClick?.(workbookTarget);
      return;
    }

    const personalApi = window.excel?.personal;
    if (!personalApi) {
      onActionStatus?.('error', 'PERSONAL.XLSB actions are unavailable.');
      return;
    }

    const runAction = action === 'create_file'
      ? personalApi.create
      : personalApi.open;
    if (typeof runAction !== 'function') {
      onActionStatus?.('error', 'PERSONAL.XLSB actions are unavailable.');
      return;
    }

    setPersonalActionInFlight(true);
    onActionStatus?.(
      'running',
      action === 'create_file' ? 'Creating PERSONAL.xlsb...' : 'Opening PERSONAL.xlsb...'
    );

    try {
      const result = await runAction();
      if (!result?.success) {
        const message = String(
          result?.message ||
          (action === 'create_file'
            ? 'Unable to create PERSONAL.xlsb.'
            : 'Unable to open PERSONAL.xlsb.')
        );
        onActionStatus?.('error', message);
        return;
      }

      await Promise.allSettled([
        Promise.resolve(personalMacrosState.refresh?.()),
        Promise.resolve(workbookPickerState.refreshWorkbooks?.({ silent: true }))
      ]);

      onActionStatus?.(
        'success',
        action === 'create_file'
          ? 'Created and opened PERSONAL.xlsb.'
          : (result.alreadyOpen ? 'PERSONAL.xlsb is already open.' : 'Opened PERSONAL.xlsb.')
      );
    } catch (error) {
      const message = error?.message
        ? String(error.message)
        : (action === 'create_file'
          ? 'Unable to create PERSONAL.xlsb.'
          : 'Unable to open PERSONAL.xlsb.');
      onActionStatus?.('error', message);
    } finally {
      setPersonalActionInFlight(false);
    }
  }, [
    onActionStatus,
    onBuildModeClick,
    personalActionInFlight,
    personalMacrosState.refresh,
    personalMacrosState.workbook?.path,
    personalMacrosState.workbookPath,
    workbookPickerState.refreshWorkbooks
  ]);

  const toggleWorkbookMenu = useCallback(() => {
    setWorkbookMenuOpen((previous) => {
      const next = !previous;
      if (next) {
        workbookPickerState.refreshWorkbooks({ silent: true });
      }
      return next;
    });
  }, [workbookPickerState]);

  const handleWorkbookSelection = useCallback((workbookKey) => {
    workbookPickerState.setSelectedWorkbookKey(workbookKey);
    setWorkbookMenuOpen(false);
  }, [workbookPickerState]);

  useEffect(() => {
    if (!onSelectedWorkbookForBuildChange) {
      return;
    }
    onSelectedWorkbookForBuildChange(selectedWorkbook || null);
  }, [
    onSelectedWorkbookForBuildChange,
    selectedWorkbook?.key,
    selectedWorkbook?.name,
    selectedWorkbook?.path
  ]);

  useEffect(() => {
    if (!isWorkbookMenuOpen) {
      return;
    }

    const handlePointerDown = (event) => {
      if (workbookMenuRef.current && !workbookMenuRef.current.contains(event.target)) {
        setWorkbookMenuOpen(false);
      }
    };

    const handleEscape = (event) => {
      if (event.key === 'Escape') {
        setWorkbookMenuOpen(false);
      }
    };

    document.addEventListener('mousedown', handlePointerDown);
    window.addEventListener('keydown', handleEscape);
    return () => {
      document.removeEventListener('mousedown', handlePointerDown);
      window.removeEventListener('keydown', handleEscape);
    };
  }, [isWorkbookMenuOpen]);

  useEffect(() => {
    if (!personalInfoPinned) {
      return;
    }

    const handlePointerDown = (event) => {
      if (personalInfoRef.current && !personalInfoRef.current.contains(event.target)) {
        setPersonalInfoPinned(false);
      }
    };

    const handleEscape = (event) => {
      if (event.key === 'Escape') {
        setPersonalInfoPinned(false);
      }
    };

    document.addEventListener('mousedown', handlePointerDown);
    window.addEventListener('keydown', handleEscape);
    return () => {
      document.removeEventListener('mousedown', handlePointerDown);
      window.removeEventListener('keydown', handleEscape);
    };
  }, [personalInfoPinned]);

  useEffect(() => {
    if (status !== 'ready') {
      setWorkbookMenuOpen(false);
      setPersonalInfoHover(false);
      setPersonalInfoPinned(false);
    }
  }, [status]);

  const renderNonReadyState = () => {
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
  };

  const workbookDataIsLoading = displayedWorkbookData.status === 'loading';
  const workbookDataHasError = displayedWorkbookData.status === 'error';
  const workbookDataIsReady = displayedWorkbookData.status === 'ready';
  const canRenderMacroRows = workbookDataIsReady || (workbookDataIsLoading && activeMacroRows.length > 0);
  const workbookMacroCount = Array.isArray(displayedWorkbookData.macros)
    ? displayedWorkbookData.macros.length
    : 0;
  const macrosEmptyMessage = workbookDataIsLoading
    ? `Loading macros from ${selectedWorkbookLabel}...`
    : workbookDataHasError
      ? selectedWorkbookErrorMessage
      : workbookMacroCount < 1
        ? 'This workbook has no macros.'
        : 'No macros match this search.';

  return (
    <main className="main-content">
      {status !== 'ready' ? renderNonReadyState() : (
        <>
          <section className="search-ready-section search-ready-section-first">
            <WorkbookPicker
              menuRef={workbookMenuRef}
              isOpen={isWorkbookMenuOpen}
              selectedWorkbookLabel={selectedWorkbookLabel}
              pickerStatus={workbookPickerState.pickerStatus}
              pickerError={workbookPickerState.pickerError}
              workbooks={workbookPickerState.workbooks}
              selectedWorkbookKey={workbookPickerState.selectedWorkbookKey}
              onToggle={toggleWorkbookMenu}
              onSelectWorkbook={handleWorkbookSelection}
            />

            <ShortcutGrid
              rows={canRenderMacroRows ? activeMacroRows : []}
              shortcutState={effectiveShortcutState}
              selectedMacroId={usingActiveWorkbookShortcuts ? selectedMacroId : null}
              onRunMacro={onRunMacro}
              emptyMessage={macrosEmptyMessage}
              showEmptyState={workbookDataHasError || activeMacroRows.length === 0}
            />
          </section>

          <PersonalMacrosSection
            sectionModel={personalSectionModel}
            infoRef={personalInfoRef}
            showInfo={showPersonalInfo}
            onInfoHoverChange={setPersonalInfoHover}
            onInfoToggle={() => setPersonalInfoPinned((previous) => !previous)}
            actionInFlight={personalActionInFlight}
            onAction={handlePersonalAction}
            rows={personalMacroRows}
            shortcutState={personalShortcutState}
            selectedMacroId={selectedMacroId}
            onRunMacro={onRunMacro}
          />
        </>
      )}
    </main>
  );
};

export default ShortcutsPage;
