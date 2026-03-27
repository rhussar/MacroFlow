import React, { useEffect, useRef, useState } from 'react';
import {
  MailIcon,
  SettingsIcon,
  ExitIcon,
  SunIcon,
  MoonIcon
} from './icons';
import { getMacroStats, formatTimeSaved } from '../features/run/macroStats';
/* global __APP_VERSION__ */

function getTheme() {
  return document.documentElement.getAttribute('data-theme') || 'dark';
}

function setTheme(theme) {
  document.documentElement.setAttribute('data-theme', theme);
  localStorage.setItem('macroflow-theme', theme);
}

function SettingsPanel({ onClose }) {
  const [theme, setThemeState] = useState(getTheme);
  const [licenseInfo, setLicenseInfo] = useState(null);
  const [updateStatus, setUpdateStatus] = useState('Up to date');
  const [aiStatus, setAiStatus] = useState(null);
  const [aiNotice, setAiNotice] = useState('');
  const previousRemoveInProgressRef = useRef(false);
  const aiProgressPercent = typeof aiStatus?.progress === 'number'
    ? Math.max(0, Math.min(100, Math.round(aiStatus.progress * 100)))
    : null;
  const showAiDetail = Boolean(
    aiStatus?.setupInProgress ||
    aiStatus?.removeInProgress ||
    aiStatus?.stage === 'error' ||
    aiStatus?.stage === 'runtime_conflict'
  );

  useEffect(() => {
    window.excel?.license?.getStatus?.().then((status) => {
      if (status?.licenseData) {
        setLicenseInfo(status.licenseData);
      }
    }).catch(() => {});

    window.excel?.updater?.checkNow?.().then((result) => {
      if (result?.success) {
        setUpdateStatus('Checking...');
      } else if (result?.reason === 'development mode') {
        setUpdateStatus('Dev mode');
      } else {
        setUpdateStatus('Up to date');
      }
    }).catch(() => {});

    const unsubscribeAi = window.excel?.ai?.onStatus?.((status) => {
      setAiStatus(status || null);
    }) || (() => {});

    window.excel?.ai?.getStatus?.().then((status) => {
      setAiStatus(status || null);
    }).catch(() => {});

    return () => {
      unsubscribeAi();
    };
  }, []);

  useEffect(() => {
    const wasRemoving = previousRemoveInProgressRef.current;
    const isRemoving = Boolean(aiStatus?.removeInProgress);

    if (isRemoving) {
      setAiNotice('');
    } else if (
      wasRemoving &&
      aiStatus &&
      !aiStatus.modelInstalled &&
      aiStatus.runtimeInstalled &&
      aiStatus.stage !== 'error'
    ) {
      setAiNotice('Local model removed.');
      const timer = setTimeout(() => {
        setAiNotice('');
      }, 4000);
      previousRemoveInProgressRef.current = isRemoving;
      return () => clearTimeout(timer);
    }

    previousRemoveInProgressRef.current = isRemoving;
    return undefined;
  }, [aiStatus]);

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

  const handleAiSetup = async () => {
    await window.excel?.ai?.setup?.();
  };

  const handleAiRemoveModel = async () => {
    const confirmed = window.confirm(
      'Remove the local AI model from this device? You can install it again later from Settings or Create.'
    );

    if (!confirmed) {
      return;
    }

    await window.excel?.ai?.removeModel?.();
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
          <div className="settings-section">
            <div className="settings-section-label">Account</div>
            <div className="settings-row">
              <span className="settings-row-label">Email</span>
              <span className="settings-row-value">{licenseInfo?.email || '-'}</span>
            </div>
            <div className="settings-row">
              <span className="settings-row-label"> </span>
              <span className="settings-row-value">
                <button className="settings-sign-out-btn danger" onClick={handleSignOut}>
                  Sign Out
                </button>
              </span>
            </div>
          </div>

          <div className="settings-section">
            <div className="settings-section-label">Appearance</div>
            <div className="settings-row">
              <span className="settings-row-label">Theme</span>
              <span className="settings-row-value">
                <button className="settings-theme-btn" onClick={toggleTheme}>
                  {theme === 'dark' ? (
                    <><SunIcon size={14} /> Light</>
                  ) : (
                    <><MoonIcon size={14} /> Dark</>
                  )}
                </button>
              </span>
            </div>
          </div>

          <div className="settings-section">
            <div className="settings-section-label">About</div>
            <div className="settings-row">
              <span className="settings-row-label">Version</span>
              <span className="settings-row-value">{__APP_VERSION__}</span>
            </div>
            <div className="settings-row">
              <span className="settings-row-label">Updates</span>
              <span className="settings-row-value">{updateStatus}</span>
            </div>
          </div>

          <div className="settings-section">
            <div className="settings-section-label">Local AI</div>
            <div className="settings-row">
              <span className="settings-row-label">Status</span>
              <span className="settings-row-value">
                {aiStatus?.ready
                  ? 'Ready'
                  : aiStatus?.removeInProgress
                    ? 'Removing'
                    : aiStatus?.setupInProgress
                      ? 'Setting up'
                      : 'Needs setup'}
              </span>
            </div>
            <div className="settings-row">
              <span className="settings-row-label">Model</span>
              <span className="settings-row-value">{aiStatus?.model || 'qwen2.5-coder:3b'}</span>
            </div>
            <div className="settings-row">
              <span className="settings-row-label">Runtime</span>
              <span className="settings-row-value">
                {aiStatus?.runtimeInstalled ? 'Installed' : 'Not installed'}
              </span>
            </div>
            {showAiDetail && (
              <div className="settings-subcopy">
                {aiStatus?.removeInProgress
                  ? aiStatus?.statusText || 'Removing local model...'
                  : aiStatus?.statusText || 'Checking local AI...'}
              </div>
            )}
            {aiNotice && !showAiDetail && (
              <div className="settings-subcopy success">{aiNotice}</div>
            )}
            {aiProgressPercent !== null && (
              <div className="settings-progress">
                <div className="settings-progress-track">
                  <div className="settings-progress-fill" style={{ width: `${aiProgressPercent}%` }} />
                </div>
                <div className="settings-progress-label">{aiProgressPercent}%</div>
              </div>
            )}
            {!aiStatus?.ready && !aiStatus?.removeInProgress && (
              <div className="settings-row">
                <span className="settings-row-label">Setup</span>
                <span className="settings-row-value">
                  <button
                    type="button"
                    className="settings-text-btn-inline"
                    onClick={handleAiSetup}
                    disabled={Boolean(aiStatus?.setupInProgress || aiStatus?.removeInProgress)}
                  >
                    {aiStatus?.setupInProgress ? 'Setting up...' : 'Install Local AI'}
                  </button>
                </span>
              </div>
            )}
            {(aiStatus?.modelInstalled || aiStatus?.removeInProgress) && (
              <div className="settings-row">
                <span className="settings-row-label">Actions</span>
                <span className="settings-row-value">
                  <button
                    type="button"
                    className="settings-text-btn-inline"
                    onClick={handleAiRemoveModel}
                    disabled={Boolean(aiStatus?.setupInProgress || aiStatus?.removeInProgress)}
                  >
                    {aiStatus?.removeInProgress ? 'Removing...' : 'Remove Local AI Model'}
                  </button>
                </span>
              </div>
            )}
          </div>
        </div>
      </div>
    </>
  );
}

const SettingsMenu = ({ isOpen, onClose, onQuit }) => {
  const [panelOpen, setPanelOpen] = useState(false);
  const stats = isOpen ? getMacroStats() : null;

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
      icon: <MailIcon size={16} />,
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
      icon: <ExitIcon size={16} />,
      label: 'Quit MacroFlow',
      action: onQuit,
      danger: true,
    },
  ];

  return (
    <>
      <div className="settings-menu-overlay" onClick={onClose} />
      <div className="settings-menu">
        {stats && stats.totalRuns > 0 && (
          <div className="settings-menu-stats">
            <span className="settings-menu-stat">
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2" /></svg>
              {stats.totalRuns.toLocaleString()} runs
            </span>
            <span className="settings-menu-stat">
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="10" /><polyline points="12 6 12 12 16 14" /></svg>
              {formatTimeSaved(stats.totalTimeSavedMs)} saved
            </span>
          </div>
        )}
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
