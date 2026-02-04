import React, { useState, useEffect } from 'react';
import { ArrowLeftIcon, CloseIcon, CheckIcon, MacroFlowLogo } from './icons';
import CodePreview from './CodePreview';

// Mock generated VBA code
const mockGeneratedCode = `Sub CleanData()
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

// Build states: 'empty', 'typing', 'processing', 'complete', 'success', 'error', 'fixing', 'fixed'
const BuildMode = ({ onBack, onClose, onEditMode }) => {
  const [prompt, setPrompt] = useState('');
  const [followUp, setFollowUp] = useState('');
  const [buildState, setBuildState] = useState('empty');
  const [steps, setSteps] = useState([]);
  const [errorInfo, setErrorInfo] = useState(null);

  // Handle prompt submission
  const handleSubmit = () => {
    if (!prompt.trim()) return;

    setBuildState('processing');
    setSteps([
      { text: 'Scanned Active Workbook', status: 'loading' },
    ]);

    // Simulate processing steps
    setTimeout(() => {
      setSteps([
        { text: 'Scanned Active Workbook', status: 'complete' },
        { text: 'Scanned Active Workbook', status: 'loading' },
      ]);
    }, 800);

    setTimeout(() => {
      setSteps([
        { text: 'Scanned Active Workbook', status: 'complete' },
        { text: 'Detected Range F1:C140', status: 'complete' },
        { text: 'Built custom macro', status: 'loading' },
      ]);
    }, 1600);

    setTimeout(() => {
      setSteps([
        { text: 'Scanned Active Workbook', status: 'complete' },
        { text: 'Detected Range F1:C140', status: 'complete' },
        { text: 'Built custom macro', status: 'complete' },
      ]);
      setBuildState('complete');
    }, 2400);
  };

  // Handle running macro
  const handleRunMacro = () => {
    // Simulate success or error randomly for demo
    const isSuccess = Math.random() > 0.3;

    if (isSuccess) {
      setBuildState('success');
    } else {
      setBuildState('error');
      setErrorInfo({
        title: 'Run Failed: Runtime Error 1004',
        line: 12,
      });
      setSteps([
        { text: 'Found mistake on line 12', status: 'complete' },
        { text: 'Attempting to fix', status: 'loading' },
      ]);

      // Simulate fixing
      setTimeout(() => {
        setSteps([
          { text: 'Found mistake on line 12', status: 'complete' },
          { text: 'Fixed error', status: 'complete' },
        ]);
        setBuildState('fixed');
        setErrorInfo(null);
      }, 2000);
    }
  };

  // Keyboard shortcut handler
  useEffect(() => {
    const handleKeyDown = (e) => {
      // Tab to switch modes (handled by parent)
      // Enter to submit prompt
      if (e.key === 'Enter' && buildState === 'empty' && prompt.trim()) {
        e.preventDefault();
        handleSubmit();
      }
      // Shift+K to run macro
      if (e.shiftKey && e.key === 'K' && ['complete', 'success', 'fixed'].includes(buildState)) {
        e.preventDefault();
        handleRunMacro();
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [buildState, prompt]);

  // Get footer content based on state
  const getFooterContent = () => {
    switch (buildState) {
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
          text: 'Error Occurred',
          textClass: 'error',
          action: 'Run again',
        };
      case 'complete':
      case 'fixed':
        return {
          logoClass: '',
          text: null,
          action: 'Run macro',
        };
      default:
        return {
          logoClass: '',
          text: null,
          action: null,
        };
    }
  };

  const footerContent = getFooterContent();
  const showCodePanel = ['processing', 'complete', 'success', 'error', 'fixing', 'fixed'].includes(buildState);

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
            placeholder={showCodePanel ? 'Ask follow-up...' : 'Ask AI to build any macro...'}
            value={showCodePanel ? followUp : prompt}
            onChange={(e) => showCodePanel ? setFollowUp(e.target.value) : setPrompt(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && !showCodePanel && prompt.trim()) {
                handleSubmit();
              }
            }}
          />
        </div>

        <div className="header-actions">
          <button className="close-btn" onClick={onClose}>
            <CloseIcon />
          </button>
        </div>
      </header>

      {/* Main Content */}
      <main className="main-content">
        {buildState === 'empty' ? (
          /* Empty State */
          <div className="build-empty-state">
            <h1 className="build-empty-title">Ask AI to build any macro</h1>
            <p className="build-empty-subtitle">MacroFlow's AI agent is here to help</p>
          </div>
        ) : (
          /* Split View with Steps and Code */
          <div className="split-view">
            {/* Left Panel - Conversation */}
            <div className="split-left">
              <div className="conversation-panel">
                {/* User Prompt */}
                <div className="user-prompt">{prompt}</div>

                {/* Error Title (if error) */}
                {errorInfo && (
                  <div className="error-title">{errorInfo.title}</div>
                )}

                {/* Status Steps */}
                <div className="status-steps">
                  {steps.map((step, index) => (
                    <div
                      key={index}
                      className={`status-step ${step.status === 'complete' ? 'completed' : ''}`}
                    >
                      {step.status === 'complete' ? (
                        <span className="status-checkbox checked">
                          <CheckIcon size={12} />
                        </span>
                      ) : step.status === 'loading' ? (
                        <span className="status-spinner" />
                      ) : (
                        <span className="status-checkbox" />
                      )}
                      <span>{step.text}</span>
                    </div>
                  ))}
                </div>

                {/* Completion Message */}
                {buildState === 'complete' && (
                  <div className="completion-message">
                    Done! Press Shift + K to run this macro. Would you like to make any changes?
                  </div>
                )}

                {/* Fix Message */}
                {buildState === 'fixed' && (
                  <div className="fix-message">
                    Error Fixed, the macro had an incorrect function name. Try running again!
                  </div>
                )}
              </div>
            </div>

            {/* Right Panel - Code */}
            <div className="split-right">
              <CodePreview
                code={mockGeneratedCode}
                title="Preview"
                showHeader={true}
                showEdit={true}
                onEdit={onEditMode}
                status={
                  buildState === 'success'
                    ? 'success'
                    : buildState === 'error'
                    ? 'error'
                    : 'normal'
                }
                errorLine={errorInfo?.line}
              />
            </div>
          </div>
        )}
      </main>

      {/* Custom Footer for Build Mode */}
      {showCodePanel && footerContent.action && (
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
            <span className="footer-action" onClick={handleRunMacro} style={{ cursor: 'pointer' }}>
              {footerContent.action}
              <span className="kbd">Shift</span>
              <span className="kbd">K</span>
            </span>
          </div>
        </footer>
      )}

      {/* Footer for empty state (just logo) */}
      {!showCodePanel && (
        <footer className="footer">
          <div className="footer-left">
            <div className="logo">
              <MacroFlowLogo size={20} />
            </div>
          </div>
        </footer>
      )}
    </>
  );
};

export default BuildMode;