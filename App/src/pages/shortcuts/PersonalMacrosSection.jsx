import ShortcutGrid from './ShortcutGrid';

function getPersonalActionButtonLabel(action) {
  if (action === 'create_global_macro') {
    return 'Create a macro +';
  }
  if (action === 'create_file') {
    return 'Create PERSONAL.xlsb';
  }
  if (action === 'open_file') {
    return 'Open PERSONAL.xlsb';
  }
  return '';
}

function PersonalMacrosSection({
  sectionModel,
  actionInFlight,
  onAction,
  rows,
  shortcutState,
  selectedMacroId,
  onRunMacro,
  onEditMacro,
  onActionStatus,
  onRenameComplete,
  workbook
}) {
  if (sectionModel.hidden) {
    return null;
  }

  return (
    <section className="search-ready-section">
      <div className="personal-picker-wrap">
        <div className="personal-picker-group">
          <div className="personal-picker">
            <span className="personal-picker-label">Personal Macros</span>
          </div>
        </div>
      </div>

      {sectionModel.isEmpty ? (
        <div className="search-empty-state global-macros-empty">
          {sectionModel.isLoading ? (
            <div className="ai-loading-indicator">
              <span className="ai-loading-dot" />
              <span className="ai-loading-dot" />
              <span className="ai-loading-dot" />
            </div>
          ) : sectionModel.action ? (
            <button
              type="button"
              className="global-macros-empty-action"
              onClick={() => onAction(sectionModel.action)}
              disabled={actionInFlight}
            >
              {getPersonalActionButtonLabel(sectionModel.action)}
            </button>
          ) : (
            <span className="global-macros-empty-text">
              {sectionModel.emptyMessage}
            </span>
          )}
        </div>
      ) : (
        <ShortcutGrid
          rows={rows}
          shortcutState={shortcutState}
          selectedMacroId={selectedMacroId}
          onRunMacro={onRunMacro}
          onEditMacro={onEditMacro}
          onActionStatus={onActionStatus}
          onRenameComplete={onRenameComplete}
          workbook={workbook}
        />
      )}
    </section>
  );
}

export default PersonalMacrosSection;
