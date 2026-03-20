import { useState, useEffect } from 'react';
import { CloseIcon, MacroFlowLogo } from './icons';

/* global __APP_VERSION__ */

/**
 * LicenseGate — shown instead of the main app when no valid license is found.
 * Primary flow: Sign In via Auth0 (creates/activates license automatically).
 * Fallback: manual license key entry (for offline / enterprise users).
 */
export default function LicenseGate({ onLicensed }) {
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


  return (
    <div className="app-container">
      <div className="license-gate">
        {/* Close button */}
        <div className="license-gate-controls">
          <button
            className="window-control-btn close"
            onClick={() => window.excel?.app?.close?.()}
            title="Close"
          >
            <CloseIcon size={18} />
          </button>
        </div>

        {/* Branding */}
        <div className="license-gate-branding">
          <MacroFlowLogo size={96} />
          <div className="license-gate-title">MacroFlow</div>
          <div className="license-gate-subtitle">Turn Hours into Seconds</div>
        </div>

        {/* Body */}
        <div className="license-gate-body">
          <button
            className="license-gate-btn"
            onClick={handleSignIn}
            disabled={loading}
          >
            {loading ? 'Signing in...' : 'Get Started'}
          </button>

          {error && <div className="license-gate-error">{error}</div>}

          <div className="license-gate-beta">
            Beta access. Free for early users.
          </div>
        </div>

        {/* Footer */}
        <div className="license-gate-footer">
          <span className="license-gate-version">v{__APP_VERSION__}</span>
          <button
            className="license-gate-footer-text license-gate-link"
            onClick={() => window.excel?.app?.openExternal?.('mailto:ronan@macroflow.ai?subject=MacroFlow%20Support%20Ticket')}
          >
            Need help? ronan@macroflow.ai
          </button>
        </div>
      </div>
    </div>
  );
}
