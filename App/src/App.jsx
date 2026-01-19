import React, { useState } from 'react';

function App() {
  const [status, setStatus] = useState({ show: false, message: '', type: '' });
  const [isLoading, setIsLoading] = useState(false);

  const handleTestConnection = async () => {
    setIsLoading(true);
    setStatus({ show: true, message: 'Testing connection...', type: '' });

    try {
      const result = await window.electronAPI.testExcelConnection();

      if (result.success) {
        setStatus({
          show: true,
          message: `✓ ${result.message}`,
          type: 'success'
        });
      } else {
        setStatus({
          show: true,
          message: `✗ Error: ${result.message}`,
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

  return (
    <div className="app-container">
      <div className="content">
        <h1>MacroFlow Electron</h1>
        <h2>Proof of Life - Excel COM Integration</h2>

        <div className="test-section">
          <button
            className="test-button"
            onClick={handleTestConnection}
            disabled={isLoading}
          >
            {isLoading ? 'Testing...' : 'Test Connection'}
          </button>

          {status.show && (
            <div className={`status ${status.type}`}>
              {status.message}
            </div>
          )}
        </div>

        <div className="info">
          <h3>What this does:</h3>
          <ol>
            <li>Launches Microsoft Excel via Windows COM</li>
            <li>Creates a new workbook</li>
            <li>Writes "Hello World" to cell A1</li>
            <li>Formats the text (bold, red, size 14)</li>
          </ol>
        </div>
      </div>
    </div>
  );
}

export default App;
