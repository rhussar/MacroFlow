import React, { useState, useEffect } from 'react';
import {
  MailIcon,
  SettingsIcon,
  ExitIcon,
  SunIcon,
  MoonIcon
} from './icons';
/* global __APP_VERSION__ */

function getTheme() {
  return document.documentElement.getAttribute('data-theme') || 'dark';
}

function setTheme(theme) {
  document.documentElement.setAttribute('data-theme', theme);
  localStorage.setItem('macroflow-theme', theme);
}

// ── Settings Panel (full overlay) ───────────────────────────────────────────

function SettingsPanel({ onClose }) {
  const [theme, setThemeState] = useState(getTheme);
  const [licenseInfo, setLicenseInfo] = useState(null);

  useEffect(() => {
    window.excel?.license?.getStatus?.().then((status) => {
      if (status?.licenseData) {
        setLicenseInfo(status.licenseData);
      }
    }).catch(() => {});
  }, []);

  const toggleTheme = () => {
    const next = theme === 'dark' ? 'light' : 'dark';
    setTheme(next);
    setThemeState(next);
  };

  const handleSignOut = async () => {
    await window.excel?.license?.deactivate?.();
    await window.excel?.auth?.logout?.();
    window.location.reload();
  };

  return (
    <>
      <div className="settings-panel-overlay" onClick={onClose} />
      <div className="settings-panel">
        <div className="settings-panel-header">
          <span className="settings-panel-title">Settings</span>
          <button className="settings-panel-close" onClick={onClose}>&times;</button>
        </div>

        <div className="settings-panel-content">
          {/* Account Section */}
          <div className="settings-section">
            <div className="settings-section-label">Account</div>
            <div className="settings-row">
              <span className="settings-row-label">Email</span>
              <span className="settings-row-value">{licenseInfo?.email || '—'}</span>
            </div>
            <div className="settings-row">
              <span className="settings-row-label">License</span>
              <span className="settings-row-value settings-row-badge">
                {licenseInfo?.status === 'ACTIVE' ? 'Active' : licenseInfo?.status || '—'}
              </span>
            </div>
            <button className="settings-text-btn danger" onClick={handleSignOut}>
              Sign Out
            </button>
          </div>

          {/* Appearance Section */}
          <div className="settings-section">
            <div className="settings-section-label">Appearance</div>
            <div className="settings-row clickable" onClick={toggleTheme}>
              <span className="settings-row-label">Theme</span>
              <span className="settings-row-value settings-row-toggle">
                {theme === 'dark' ? (
                  <><SunIcon size={14} /> Light</>
                ) : (
                  <><MoonIcon size={14} /> Dark</>
                )}
              </span>
            </div>
          </div>

          {/* About Section */}
          <div className="settings-section">
            <div className="settings-section-label">About</div>
            <div className="settings-row">
              <span className="settings-row-label">Version</span>
              <span className="settings-row-value">{__APP_VERSION__}</span>
            </div>
            <div className="settings-row">
              <span className="settings-row-label">Updates</span>
              <span className="settings-row-value">
                <button
                  className="settings-text-btn-inline"
                  onClick={() => window.excel?.updater?.checkNow?.()}
                >
                  Check now
                </button>
              </span>
            </div>
          </div>
        </div>
      </div>
    </>
  );
}

// ── Settings Menu (quick dropdown from gear icon) ───────────────────────────

const SettingsMenu = ({ isOpen, onClose, onQuit }) => {
  const [panelOpen, setPanelOpen] = useState(false);

  if (panelOpen) {
    return (
      <SettingsPanel
        onClose={() => {
          setPanelOpen(false);
          onClose();
        }}
      />
    );
  }

  if (!isOpen) {
    return null;
  }

  const menuItems = [
    {
      icon: <MailIcon />,
      label: 'Send Feedback',
      action: () => {
        window.excel?.app?.openExternal?.('mailto:ronan@macroflow.ai?subject=MacroFlow%20Feedback');
      },
    },
    {
      icon: <SettingsIcon size={16} />,
      label: 'Settings',
      action: () => setPanelOpen(true),
      keepOpen: true,
    },
    {
      icon: <ExitIcon />,
      label: 'Quit MacroFlow',
      action: onQuit,
      danger: true,
    },
  ];

  return (
    <>
      <div className="settings-menu-overlay" onClick={onClose} />
      <div className="settings-menu">
        <div className="settings-menu-header">MacroFlow v{__APP_VERSION__}</div>
        {menuItems.map((item, index) => (
          <button
            type="button"
            key={`${item.label}-${index}`}
            className={`settings-menu-item ${item.danger ? 'danger' : ''}`}
            onClick={() => {
              item.action?.();
              if (!item.keepOpen) {
                onClose();
              }
            }}
          >
            <span className="settings-menu-icon">{item.icon}</span>
            <span className="settings-menu-item-body">
              <span className="settings-menu-item-label">{item.label}</span>
            </span>
          </button>
        ))}
      </div>
    </>
  );
};

export default SettingsMenu;
