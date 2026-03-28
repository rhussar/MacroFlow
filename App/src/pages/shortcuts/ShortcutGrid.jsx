import React, { useState, useCallback } from 'react';
import { ReturnIcon } from '../../components/icons';
import { formatShortcutPrefix } from '../../lib/shortcut-keybind';
import ImageMsoIcon, { isSpriteReady } from '../../components/ImageMsoIcon';
import MacroContextMenu from '../../components/MacroContextMenu';
import IconPicker from '../../components/IconPicker';
import { getMacroIcon, setMacroIcon, removeMacroIcon } from '../../features/icons/macroIconStore';

function ShortcutGrid({
  rows = [],
  shortcutState,
  selectedMacroId = null,
  onRunMacro,
  emptyMessage = '',
  showEmptyState = false,
  isLoading = false
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

  // Context menu state
  const [contextMenu, setContextMenu] = useState(null);
  // Icon picker state
  const [pickerMacroId, setPickerMacroId] = useState(null);
  // Force re-render after icon changes
  const [iconVersion, setIconVersion] = useState(0);

  const handleContextMenu = useCallback((e, macroId) => {
    e.preventDefault();
    e.stopPropagation();
    setContextMenu({ x: e.clientX, y: e.clientY, macroId });
  }, []);

  const handleAssignIcon = useCallback(() => {
    if (contextMenu) {
      setPickerMacroId(contextMenu.macroId);
    }
  }, [contextMenu]);

  const handleRemoveIcon = useCallback(() => {
    if (contextMenu) {
      removeMacroIcon(contextMenu.macroId);
      setIconVersion(v => v + 1);
    }
  }, [contextMenu]);

  const handleIconSelect = useCallback((iconName) => {
    if (pickerMacroId) {
      setMacroIcon(pickerMacroId, iconName);
      setIconVersion(v => v + 1);
    }
  }, [pickerMacroId]);

  return (
    <div className="shortcuts-grid">
      {showEmptyState && (
        <div className="search-empty-state search-empty-state-grid">
          {isLoading ? (
            <div className="ai-loading-indicator">
              <span className="ai-loading-dot" />
              <span className="ai-loading-dot" />
              <span className="ai-loading-dot" />
            </div>
          ) : emptyMessage}
        </div>
      )}

      {rows.map((row) => {
        const macro = row.macro;
        const currentShortcutLetter =
          shortcutDraftByMacroId[macro.id] ?? shortcutByMacroId[macro.id] ?? '';
        const hasInputError = Boolean(shortcutInputErrorByMacroId[macro.id]);
        const isSaving = shortcutSavingMacroId === macro.id;
        const shortcutPrefix = formatShortcutPrefix(currentShortcutLetter);
        const assignedIcon = getMacroIcon(macro.id);

        return (
          <div
            key={row.uiId}
            className={`shortcut-item ${selectedMacroId === macro.id ? 'selected' : ''} ${isSaving ? 'saving' : ''}`}
            onContextMenu={(e) => handleContextMenu(e, macro.id)}
          >
            <div className="shortcut-run-target">
              <button
                type="button"
                className="shortcut-icon-btn"
                title="Run macro"
                onClick={() => onRunMacro?.(macro)}
              >
                {isSpriteReady() ? (
                  <ImageMsoIcon name={assignedIcon || 'MacroRecord'} size={22} />
                ) : (
                  <ReturnIcon size={20} />
                )}
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

      {/* Right-click context menu */}
      {contextMenu && (
        <MacroContextMenu
          x={contextMenu.x}
          y={contextMenu.y}
          hasIcon={Boolean(getMacroIcon(contextMenu.macroId))}
          onAssignIcon={handleAssignIcon}
          onRemoveIcon={handleRemoveIcon}
          onClose={() => setContextMenu(null)}
        />
      )}

      {/* Icon picker modal */}
      {pickerMacroId && (
        <IconPicker
          currentIcon={getMacroIcon(pickerMacroId)}
          onSelect={handleIconSelect}
          onClose={() => setPickerMacroId(null)}
        />
      )}
    </div>
  );
}

export default ShortcutGrid;
