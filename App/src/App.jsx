import React, { useState, useRef } from 'react';
import Editor from '@monaco-editor/react';
import { X, Play } from 'lucide-react';

const DEFAULT_CODE = `' Write your macro here
Sub MyMacro()
    MsgBox "Hello from MacroFlow"
End Sub`;

function App() {
  const [code, setCode] = useState(DEFAULT_CODE);
  const [status, setStatus] = useState({ show: false, message: '', type: '' });
  const [isLoading, setIsLoading] = useState(false);
  const editorRef = useRef(null);

  const handleEditorDidMount = (editor, monaco) => {
    editorRef.current = editor;
  };

  const handlePushToExcel = async () => {
    setIsLoading(true);
    setStatus({ show: true, message: 'Injecting code into Excel...', type: '' });

    try {
      const result = await window.electronAPI.injectCode(code);

      if (result.success) {
        setStatus({
          show: true,
          message: `✓ ${result.message}`,
          type: 'success'
        });
      } else {
        setStatus({
          show: true,
          message: `✗ ${result.message}`,
          type: 'error'
        });
      }
    } catch (error) {
      setStatus({
        show: true,
        message: `✗ Failed: ${error.message}`,
        type: 'error'
      });
    } finally {
      setIsLoading(false);
    }
  };

  const handleClose = () => {
    if (window.electronAPI?.closeApp) {
      window.electronAPI.closeApp();
    } else {
      window.close();
    }
  };

  return (
    <div className="app-root">
      {/* Header */}
      <div className="header">
        <div className="header-title">MacroFlow Editor</div>
        <button className="header-close-button" onClick={handleClose} title="Close">
          <X size={16} />
        </button>
      </div>

      {/* Editor */}
      <div className="editor-container">
        <Editor
          height="100%"
          language="vb"
          theme="vs-dark"
          value={code}
          onChange={(value) => setCode(value || '')}
          onMount={handleEditorDidMount}
          options={{
            minimap: { enabled: false },
            fontSize: 14,
            lineNumbers: 'on',
            scrollBeyondLastLine: false,
            automaticLayout: true,
            tabSize: 4,
            wordWrap: 'on'
          }}
        />
      </div>

      {/* Footer */}
      <div className="footer">
        {status.show && (
          <div className={`footer-status ${status.type}`}>
            {status.message}
          </div>
        )}
        <button
          className="push-button"
          onClick={handlePushToExcel}
          disabled={isLoading}
        >
          <Play size={20} />
          {isLoading ? 'Pushing...' : 'Push to Excel'}
        </button>
      </div>
    </div>
  );
}

export default App;
