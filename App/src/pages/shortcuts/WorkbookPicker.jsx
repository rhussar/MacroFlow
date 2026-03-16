import React from 'react';
import { ChevronDownIcon, WorkbookIcon } from '../../components/icons';

function WorkbookPicker({
  menuRef,
  isOpen,
  selectedWorkbookLabel,
  pickerStatus,
  pickerError,
  workbooks = [],
  selectedWorkbookKey = '',
  onToggle,
  onSelectWorkbook
}) {
  return (
    <div className="macro-workbook-picker-wrap" ref={menuRef}>
      <div className="macro-workbook-picker-group">
        <button
          type="button"
          className={`macro-workbook-picker ${isOpen ? 'open' : ''}`}
          aria-label="Select workbook"
          aria-expanded={isOpen}
          onClick={onToggle}
        >
          <span className="macro-workbook-picker-label">{selectedWorkbookLabel}</span>
        </button>
        <button
          type="button"
          className={`macro-workbook-picker-chevron ${isOpen ? 'open' : ''}`}
          aria-label="Toggle workbook menu"
          onClick={onToggle}
        >
          <ChevronDownIcon size={16} />
        </button>
      </div>

      {isOpen && (
        <div className="macro-workbook-menu" role="listbox" aria-label="Open workbooks">
          <div className="macro-workbook-menu-title">Select open workbook</div>
          <div className="macro-workbook-menu-divider" />

          {pickerStatus === 'loading' && workbooks.length === 0 && (
            <div className="macro-workbook-menu-state">Loading open workbooks...</div>
          )}

          {pickerStatus === 'error' && (
            <div className="macro-workbook-menu-state error">
              {pickerError?.message || 'Unable to load open workbooks.'}
            </div>
          )}

          {workbooks.map((workbook) => (
            <button
              key={workbook.key}
              type="button"
              className={`macro-workbook-option ${workbook.key === selectedWorkbookKey ? 'selected' : ''}`}
              onClick={() => onSelectWorkbook(workbook.key)}
            >
              <span className="macro-workbook-option-icon">
                <WorkbookIcon size={14} />
              </span>
              <span className="macro-workbook-option-label">{workbook.name}</span>
            </button>
          ))}

          {pickerStatus === 'ready' && workbooks.length === 0 && (
            <div className="macro-workbook-menu-state">No open workbooks found.</div>
          )}
        </div>
      )}
    </div>
  );
}

export default WorkbookPicker;
