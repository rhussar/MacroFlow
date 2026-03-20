import React, { useState, useEffect } from 'react';

/* global __APP_VERSION__ */

/**
 * LicenseGate — shown instead of the main app when no valid license is found.
 * Once the user enters a valid key, calls onLicensed() to reveal the app.
 */
export default function LicenseGate({ onLicensed }) {
  const [key, setKey] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  // On mount, check if the main process already validated a cached license.
  useEffect(() => {
    window.excel?.license?.getStatus?.().then((status) => {
      if (status?.valid) {
        onLicensed(status.licenseData);
      }
    }).catch(() => {});
  }, [onLicensed]);

  const handleActivate = async () => {
    const trimmed = key.trim();
    if (!trimmed) {
      setError('Please enter a license key.');
      return;
    }

    setLoading(true);
    setError('');

    try {
      const result = await window.excel.license.activate(trimmed);
      if (result.success) {
        onLicensed(result.licenseData);
      } else {
        setError(result.message || 'Activation failed.');
      }
    } catch {
      setError('Something went wrong. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  const handleKeyDown = (e) => {
    if (e.key === 'Enter' && !loading) {
      handleActivate();
    }
  };

  return (
    <div className="app-container">
      <div className="license-gate">
        <div className="license-gate-header">
          <div className="license-gate-title">MacroFlow</div>
          <div className="license-gate-version">v{__APP_VERSION__}</div>
        </div>

        <div className="license-gate-body">
          <div className="license-gate-label">Enter your license key</div>
          <input
            className={`license-gate-input ${error ? 'has-error' : ''}`}
            type="text"
            value={key}
            onChange={(e) => { setKey(e.target.value); setError(''); }}
            onKeyDown={handleKeyDown}
            placeholder="XXXX-XXXX-XXXX-XXXX"
            disabled={loading}
            autoFocus
            spellCheck={false}
          />
          {error && <div className="license-gate-error">{error}</div>}
          <button
            className="license-gate-btn"
            onClick={handleActivate}
            disabled={loading}
          >
            {loading ? 'Activating...' : 'Activate'}
          </button>
        </div>

        <div className="license-gate-footer">
          <span className="license-gate-footer-text">
            Need a license? Visit <button
              type="button"
              className="license-gate-link"
              onClick={() => {
                try { require('electron')?.shell?.openExternal?.('https://macroflow.com'); } catch { /* renderer */ }
              }}
            >macroflow.com</button>
          </span>
        </div>

        {/* Window controls */}
        <div className="license-gate-controls">
          <button
            className="window-control-btn close"
            onClick={() => window.excel?.app?.close?.()}
            title="Close"
          >
            &times;
          </button>
        </div>
      </div>
    </div>
  );
}
