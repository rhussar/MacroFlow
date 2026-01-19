import React, { useState } from 'react'

function App() {
  const [status, setStatus] = useState<string>('')
  const [loading, setLoading] = useState<boolean>(false)

  const handleTestConnection = async () => {
    setLoading(true)
    setStatus('Testing connection...')

    try {
      const result = await window.electronAPI.testExcelConnection()

      if (result.success) {
        setStatus(`✓ ${result.message}`)
      } else {
        setStatus(`✗ Error: ${result.message}`)
      }
    } catch (error) {
      setStatus(`✗ Failed: ${error instanceof Error ? error.message : 'Unknown error'}`)
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="app-container">
      <div className="content">
        <h1>MacroFlow Electron</h1>
        <h2>Proof of Life - Excel COM Integration</h2>

        <div className="test-section">
          <button
            onClick={handleTestConnection}
            disabled={loading}
            className="test-button"
          >
            {loading ? 'Testing...' : 'Test Connection'}
          </button>

          {status && (
            <div className={`status ${status.startsWith('✓') ? 'success' : 'error'}`}>
              {status}
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
  )
}

export default App
