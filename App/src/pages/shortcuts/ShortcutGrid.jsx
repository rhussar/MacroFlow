import React from 'react';
import { ReturnIcon } from '../../components/icons';
import { formatShortcutPrefix } from '../../lib/shortcut-keybind';

function ShortcutGrid({
  rows = [],
  shortcutState,
  selectedMacroId = null,
  onRunMacro,
  emptyMessage = '',
  showEmptyState = false
}) {
  const shortcutByMacroId = shortcutState?.shortcutByMacroId || {};
  const shortcutDraftByMacroId = shortcutState?.shortcutDraftByMacroId || {};
  const shortcutInputErrorByMacroId = shortcutState?.shortcutInputErrorByMacroId || {};
  const shortcutSavingMacroId = shortcutState?.shortcutSavingMacroId || null;
  const handleShortcutDraftChange = shortcutState?.handleShortcutDraftChange;
  const handleShortcutCommit = shortcutState?.handleShortcutCommit;
  const selectShortcutInputValue = (event) => {
    event.currentTarget.select();
  };

  return (
    <div className="shortcuts-grid">
      {showEmptyState && (
        <div className="search-empty-state search-empty-state-grid">{emptyMessage}</div>
      )}

      {rows.map((row) => {
        const macro = row.macro;
        const currentShortcutLetter =
          shortcutDraftByMacroId[macro.id] ?? shortcutByMacroId[macro.id] ?? '';
        const hasInputError = Boolean(shortcutInputErrorByMacroId[macro.id]);
        const isSaving = shortcutSavingMacroId === macro.id;
        const shortcutPrefix = formatShortcutPrefix(currentShortcutLetter);

        return (
          <div
            key={row.uiId}
            className={`shortcut-item ${selectedMacroId === macro.id ? 'selected' : ''} ${isSaving ? 'saving' : ''}`}
          >
            <div className="shortcut-run-target">
              <button
                type="button"
                className="shortcut-icon-btn"
                onClick={() => onRunMacro?.(macro)}
              >
                <ReturnIcon size={20} />
              </button>
              <span className="shortcut-name">{macro.name}</span>
            </div>
            <div className="shortcut-binding" onClick={(event) => event.stopPropagation()}>
              <span className="shortcut-prefix">Ctrl +</span>
              {shortcutPrefix.includes('Shift') && (
                <span className="shortcut-shift">Shift +</span>
              )}
              <input
                type="text"
                className={`shortcut-keycap-input ${currentShortcutLetter ? '' : 'is-empty'} ${hasInputError ? 'has-error' : ''}`}
                value={currentShortcutLetter}
                placeholder=""
                maxLength={1}
                autoCapitalize="off"
                autoComplete="off"
                spellCheck={false}
                aria-label={`Shortcut letter for ${macro.name}`}
                onChange={(event) => handleShortcutDraftChange?.(macro.id, event.target.value)}
                onFocus={selectShortcutInputValue}
                onBlur={() => handleShortcutCommit?.(macro, 'blur')}
                onKeyDown={(event) => {
                  if (event.key === 'Enter') {
                    event.preventDefault();
                    event.currentTarget.blur();
                  } else if (event.key === 'Escape') {
                    event.preventDefault();
                    handleShortcutDraftChange?.(macro.id, shortcutByMacroId[macro.id] || '');
                    event.currentTarget.blur();
                  }
                }}
                onClick={(event) => {
                  event.stopPropagation();
                  selectShortcutInputValue(event);
                }}
                disabled={isSaving}
              />
            </div>
          </div>
        );
      })}
    </div>
  );
}

export default ShortcutGrid;
