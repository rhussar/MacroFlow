import React, { useState, useEffect } from 'react';

/* global __APP_VERSION__ */

/**
 * LicenseGate — shown instead of the main app when no valid license is found.
 * Primary flow: Sign In via Auth0 (creates/activates license automatically).
 * Fallback: manual license key entry (for offline / enterprise users).
 */
export default function LicenseGate({ onLicensed }) {
  const [key, setKey] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [showManualKey, setShowManualKey] = useState(false);

  // On mount, check if the main process already validated a cached license.
  useEffect(() => {
    window.excel?.license?.getStatus?.().then((status) => {
      if (status?.valid) {
        onLicensed(status.licenseData);
      }
    }).catch(() => {});
  }, [onLicensed]);

  const handleSignIn = async () => {
    setLoading(true);
    setError('');

    try {
      // Open system browser for Auth0 login.
      const authResult = await window.excel.auth.login();

      if (!authResult?.success) {
        setError(authResult?.message || 'Sign in failed.');
        setLoading(false);
        return;
      }

      // Auth0 succeeded — activate the Keygen license tied to this user.
      const licenseResult = await window.excel.license.activateFromAuth(authResult);

      if (licenseResult?.success) {
        onLicensed(licenseResult.licenseData);
      } else {
        setError(licenseResult?.message || 'License activation failed.');
      }
    } catch {
      setError('Something went wrong. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  const handleManualActivate = async () => {
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
      handleManualActivate();
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
          {!showManualKey ? (
            <>
              <button
                className="license-gate-btn"
                onClick={handleSignIn}
                disabled={loading}
              >
                {loading ? 'Signing in...' : 'Sign In'}
              </button>

              {error && <div className="license-gate-error">{error}</div>}

              <button
                type="button"
                className="license-gate-link license-gate-toggle"
                onClick={() => { setShowManualKey(true); setError(''); }}
              >
                Use a license key instead
              </button>
            </>
          ) : (
            <>
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
                onClick={handleManualActivate}
                disabled={loading}
              >
                {loading ? 'Activating...' : 'Activate'}
              </button>

              <button
                type="button"
                className="license-gate-link license-gate-toggle"
                onClick={() => { setShowManualKey(false); setError(''); }}
              >
                Sign in with your account instead
              </button>
            </>
          )}
        </div>

        <div className="license-gate-footer">
          <span className="license-gate-footer-text">
            New to MacroFlow? Sign in to get started.
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
