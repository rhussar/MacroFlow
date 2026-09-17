import React, { useState, useCallback, useEffect, useRef } from 'react';
import { ReturnIcon } from '../../components/icons';
import { formatShortcutPrefix } from '../../lib/shortcut-keybind';
import ImageMsoIcon, { isSpriteReady } from '../../components/ImageMsoIcon';
import MacroContextMenu from '../../components/MacroContextMenu';
import IconPicker from '../../components/IconPicker';
import { getMacroIcon, setMacroIcon, removeMacroIcon, renameMacroIcon } from '../../features/icons/macroIconStore';
import { displayMacroName, buildMacroRenameRequest, isValidVbaModuleName, encodeMacroName } from '../../features/search/module-actions';

function ShortcutGrid({
  rows = [],
  shortcutState,
  selectedMacroId = null,
  onRunMacro,
  onEditMacro,
  onActionStatus,
  onRenameComplete,
  workbook = null,
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
  // Force re-render after icon changes (local or from other components)
  const [iconVersion, setIconVersion] = useState(0);

  useEffect(() => {
    const handler = () => setIconVersion(v => v + 1);
    window.addEventListener('macroflow-icon-change', handler);
    return () => window.removeEventListener('macroflow-icon-change', handler);
  }, []);

  const handleContextMenu = useCallback((e, macro) => {
    e.preventDefault();
    e.stopPropagation();
    setContextMenu({ x: e.clientX, y: e.clientY, macroId: macro.id, macro });
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

  const handleEditMacro = useCallback(() => {
    if (contextMenu?.macro && onEditMacro) {
      onEditMacro(contextMenu.macro);
    }
  }, [contextMenu, onEditMacro]);

  const handleIconSelect = useCallback((iconName) => {
    if (pickerMacroId) {
      setMacroIcon(pickerMacroId, iconName);
      setIconVersion(v => v + 1);
    }
  }, [pickerMacroId]);

  // Rename state
  const [renameMacro, setRenameMacro] = useState(null);
  const [renameDraft, setRenameDraft] = useState('');
  const renameInFlightRef = useRef(false);

  const handleStartRename = useCallback(() => {
    if (contextMenu?.macro) {
      setRenameMacro(contextMenu.macro);
      setRenameDraft(displayMacroName(contextMenu.macro.name));
    }
  }, [contextMenu]);

  const handleCommitRename = useCallback(async () => {
    if (!renameMacro || renameInFlightRef.current) return;
    const nextName = encodeMacroName(renameDraft.trim());
    if (!nextName || nextName === renameMacro.name) {
      setRenameMacro(null);
      return;
    }
    if (!isValidVbaModuleName(nextName)) {
      onActionStatus?.('error', 'Invalid name.');
      return;
    }
    const renameApi = window.excel?.vba?.renameMacroByWorkbook;
    if (typeof renameApi !== 'function') {
      onActionStatus?.('error', 'Rename unavailable.');
      return;
    }
    const request = buildMacroRenameRequest(renameMacro, workbook);
    if (!request.workbookName && !request.workbookPath) {
      onActionStatus?.('error', 'Workbook not found.');
      return;
    }
    renameInFlightRef.current = true;
    try {
      const result = await renameApi({ ...request, nextMacroName: nextName });
      if (!result?.success) {
        onActionStatus?.('error', String(result?.message || 'Unable to rename macro.'));
        return;
      }
      renameMacroIcon(renameMacro.id, renameMacro.id.replace(`::${renameMacro.name}::`, `::${nextName}::`));
      setRenameMacro(null);
      onRenameComplete?.(workbook);
    } catch (error) {
      onActionStatus?.('error', error?.message ? String(error.message) : 'Unable to rename macro.');
    } finally {
      renameInFlightRef.current = false;
    }
  }, [renameMacro, renameDraft, workbook, onActionStatus]);

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
            onContextMenu={(e) => handleContextMenu(e, macro)}
          >
            <div className="shortcut-run-target">
              <button
                type="button"
                className="shortcut-icon-btn"
                title="Run macro"
                onClick={() => onRunMacro?.(macro)}
              >
                {isSpriteReady() ? (
                  <ImageMsoIcon name={assignedIcon || 'FileSaveAs'} size={22} title="" />
                ) : (
                  <ReturnIcon size={20} />
                )}
              </button>
              {renameMacro?.id === macro.id ? (
                <input
                  type="text"
                  className="shortcut-rename-input"
                  value={renameDraft}
                  autoFocus
                  spellCheck={false}
                  maxLength={80}
                  onChange={(e) => setRenameDraft(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') { e.preventDefault(); void handleCommitRename(); }
                    else if (e.key === 'Escape') { e.preventDefault(); setRenameMacro(null); }
                  }}
                  onBlur={() => void handleCommitRename()}
                  onClick={(e) => e.stopPropagation()}
                />
              ) : (
                <span className="shortcut-name">{displayMacroName(macro.name)}</span>
              )}
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
                aria-label={`Shortcut letter for ${displayMacroName(macro.name)}`}
                onChange={(event) => {
                  const input = event.target;
                  handleShortcutDraftChange?.(macro.id, input.value);
                  if (input.value) {
                    requestAnimationFrame(() => input.blur());
                  }
                }}
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
          onRename={handleStartRename}
          onEdit={onEditMacro ? handleEditMacro : undefined}
          onRun={onRunMacro ? () => onRunMacro(contextMenu.macro) : undefined}
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
