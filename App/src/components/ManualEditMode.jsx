import React, { useState, useEffect } from 'react';
import { ArrowLeftIcon, CloseIcon, MacroFlowLogo } from './icons';
import { highlightVBA } from './CodePreview';

// Mock VBA code
const mockCode = `Sub CleanData()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' 1. Remove Empty Rows
    On Error Resume Next
    ws.Columns("A:A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
    ' 2. Trim Whitespace
    For Each cell In ws.Range("B1:B150")
        cell.Value = Trim(cell.Value)
    Next cell
    
    ' 3. Fix Date Format
    ws.Columns("C:C").NumberFormat = "mm/dd/yyyy"
    
    MsgBox "Cleanup Complete!"
End Sub`;

// Edit states: 'editing', 'unsaved', 'saved', 'success', 'error'
const ManualEditMode = ({ onBack, onClose }) => {
  const [code, setCode] = useState(mockCode);
  const [originalCode, setOriginalCode] = useState(mockCode);
  const [editState, setEditState] = useState('saved');
  const [followUp, setFollowUp] = useState('');

  // Track changes
  useEffect(() => {
    if (code !== originalCode) {
      setEditState('unsaved');
    }
  }, [code, originalCode]);

  // Handle save
  const handleSave = () => {
    setOriginalCode(code);
    setEditState('saved');
  };

  // Handle run
  const handleRun = () => {
    // Simulate success or error randomly for demo
    const isSuccess = Math.random() > 0.3;
    setEditState(isSuccess ? 'success' : 'error');
  };

  // Keyboard shortcuts
  useEffect(() => {
    const handleKeyDown = (e) => {
      // Ctrl+S to save
      if (e.ctrlKey && e.key === 's') {
        e.preventDefault();
        handleSave();
      }
      // Shift+K to run
      if (e.shiftKey && e.key === 'K') {
        e.preventDefault();
        handleRun();
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [code, originalCode]);

  // Get footer content based on state
  const getFooterContent = () => {
    switch (editState) {
      case 'unsaved':
        return {
          logoClass: '',
          saveText: 'Save changes',
          saveClass: 'warning',
          action: 'Run macro',
        };
      case 'saved':
        return {
          logoClass: '',
          saveText: 'Saved',
          saveClass: 'success',
          action: 'Run macro',
        };
      case 'success':
        return {
          logoClass: 'success',
          text: 'Macro ran successfully!',
          textClass: 'success',
          action: 'Run again',
        };
      case 'error':
        return {
          logoClass: 'error',
          text: 'Error occurred',
          textClass: 'error',
          action: 'Run again',
        };
      default:
        return {
          logoClass: '',
          saveText: 'Saved',
          saveClass: 'success',
          action: 'Run macro',
        };
    }
  };

  const footerContent = getFooterContent();

  // Get code container class based on state
  const getContainerClass = () => {
    let classes = 'code-editor-container';
    if (editState === 'success') classes += ' success';
    if (editState === 'error') classes += ' error';
    return classes;
  };

  return (
    <>
      {/* Header */}
      <header className="header">
        <button className="header-back-btn" onClick={onBack}>
          <ArrowLeftIcon size={20} />
        </button>

        <div className="search-input-wrapper">
          <input
            type="text"
            className="search-input"
            placeholder="Ask follow-up..."
            value={followUp}
            onChange={(e) => setFollowUp(e.target.value)}
          />
        </div>

        <div className="header-actions">
          <button className="close-btn" onClick={onClose}>
            <CloseIcon />
          </button>
        </div>
      </header>

      {/* Code Editor */}
      <div className={getContainerClass()}>
        <div className="code-preview-header">
          <span className="code-preview-title"></span>
          <div className="code-preview-actions">
            <button className="code-action-btn" onClick={() => console.log('Open Chat')}>
              ○ Open Chat
            </button>
          </div>
        </div>
        <div
          className="code-editor-body"
          contentEditable
          suppressContentEditableWarning
          onInput={(e) => setCode(e.currentTarget.textContent || '')}
          dangerouslySetInnerHTML={{ __html: highlightVBA(code) }}
        />
      </div>

      {/* Footer */}
      <footer className="footer">
        <div className="footer-left">
          <div className={`logo ${footerContent.logoClass || ''}`}>
            <MacroFlowLogo size={20} />
          </div>
          {footerContent.text && (
            <span className={`footer-text ${footerContent.textClass || ''}`}>
              {footerContent.text}
            </span>
          )}
        </div>
        <div className="footer-right">
          {footerContent.saveText && !footerContent.text && (
            <>
              <span
                className={`footer-text ${footerContent.saveClass || ''}`}
                onClick={editState === 'unsaved' ? handleSave : undefined}
                style={editState === 'unsaved' ? { cursor: 'pointer' } : {}}
              >
                {footerContent.saveText}
              </span>
              <span className="footer-separator">|</span>
            </>
          )}
          <span className="footer-action" onClick={handleRun} style={{ cursor: 'pointer' }}>
            {footerContent.action}
            <span className="kbd">Shift</span>
            <span className="kbd">K</span>
          </span>
        </div>
      </footer>
    </>
  );
};

export default ManualEditMode;