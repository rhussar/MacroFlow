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
  visibilityControl,
  onToggleVisibility,
  onOpenFolder,
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

  const handleInfoBlur = (event) => {
    const nextTarget = event.relatedTarget;
    if (
      infoRef?.current &&
      nextTarget &&
      typeof infoRef.current.contains === 'function' &&
      infoRef.current.contains(nextTarget)
    ) {
      return;
    }
    onInfoHoverChange(false);
  };

  return (
    <section className="search-ready-section">
      <div className="personal-picker-wrap">
        <div
          ref={infoRef}
          className="personal-picker-group"
          onMouseEnter={() => onInfoHoverChange(true)}
          onMouseLeave={() => onInfoHoverChange(false)}
          onFocusCapture={() => onInfoHoverChange(true)}
          onBlurCapture={handleInfoBlur}
        >
          <div className="personal-picker">
            <span className="personal-picker-label">Global Macros</span>
          </div>
          <button
            type="button"
            className="personal-info-btn"
            aria-label="About PERSONAL.XLSB"
            aria-expanded={showInfo}
            onClick={(event) => {
              event.stopPropagation();
              onInfoToggle();
            }}
          >
            <InfoIcon size={16} />
          </button>
          {showInfo && (
            <div
              className="personal-info-tooltip"
              role="tooltip"
              onMouseEnter={() => onInfoHoverChange(true)}
              onMouseLeave={() => onInfoHoverChange(false)}
            >
              <div className="personal-info-text">Global macros in PERSONAL.XLSB are available across all workbooks.</div>
              <div className="personal-visibility-row">
                <div className="personal-visibility-copy">
                  <span className="personal-visibility-label">
                    {visibilityControl?.label || 'Show workbook'}
                  </span>
                </div>
                {visibilityControl?.showSwitch ? (
                  <button
                    type="button"
                    role="switch"
                    aria-checked={visibilityControl.checked}
                    aria-label="Toggle PERSONAL.XLSB visibility"
                    className={`personal-visibility-switch ${visibilityControl.checked ? 'on' : ''}`}
                    onClick={(event) => {
                      event.stopPropagation();
                      onToggleVisibility?.();
                    }}
                    disabled={visibilityControl.disabled}
                  >
                    <span className="personal-visibility-thumb" />
                  </button>
                ) : null}
              </div>
            </div>
          )}
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
        />
      )}
    </section>
  );
}

export default PersonalMacrosSection;
