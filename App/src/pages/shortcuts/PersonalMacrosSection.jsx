import React from 'react';
import { InfoIcon } from '../../components/icons';
import ShortcutGrid from './ShortcutGrid';

function getPersonalActionButtonLabel(action) {
  if (action === 'create_global_macro') {
    return 'Create a macro +';
  }
  if (action === 'create_file') {
    return 'Create file +';
  }
  if (action === 'open_file') {
    return 'Open file +';
  }
  return '';
}

function PersonalMacrosSection({
  sectionModel,
  infoRef,
  showInfo,
  onInfoHoverChange,
  onInfoToggle,
  actionInFlight,
  onAction,
  rows,
  shortcutState,
  selectedMacroId,
  onRunMacro
}) {
  if (sectionModel.hidden) {
    return null;
  }

  return (
    <section className="search-ready-section">
      <div className="personal-picker-wrap">
        <div ref={infoRef} className="personal-picker-group">
          <div className="personal-picker">
            <span className="personal-picker-label">Global Macros</span>
          </div>
          <button
            type="button"
            className="personal-info-btn"
            aria-label="About PERSONAL.XLSB"
            aria-expanded={showInfo}
            onMouseEnter={() => onInfoHoverChange(true)}
            onMouseLeave={() => onInfoHoverChange(false)}
            onFocus={() => onInfoHoverChange(true)}
            onBlur={() => onInfoHoverChange(false)}
            onClick={(event) => {
              event.stopPropagation();
              onInfoToggle();
            }}
          >
            <InfoIcon size={16} />
          </button>
          {showInfo && (
            <div className="personal-info-tooltip" role="tooltip">
              PERSONAL.XLSB is a hidden workbook that opens automatically with Excel. Macros stored here are available globally across all workbooks.
            </div>
          )}
        </div>
      </div>

      {sectionModel.isEmpty ? (
        <div className="search-empty-state global-macros-empty">
          {sectionModel.action ? (
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
        />
      )}
    </section>
  );
}

export default PersonalMacrosSection;
