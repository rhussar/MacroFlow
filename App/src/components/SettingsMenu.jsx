import React, { useEffect, useState } from 'react';
import {
  MailIcon,
  SettingsIcon,
  ExitIcon,
  SunIcon,
  MoonIcon,
  InfoIcon,
  SlidersIcon,
  SparkleIcon
} from './icons';
import { saveSessions } from '../features/build/sessionStore';
import { getMacroStats, formatTimeSaved } from '../features/run/macroStats';
/* global __APP_VERSION__ */

function getTheme() {
  return document.documentElement.getAttribute('data-theme') || 'dark';
}

function setTheme(theme) {
  document.documentElement.setAttribute('data-theme', theme);
  localStorage.setItem('macroflow-theme', theme);
}

const SHOW_MODULES_KEY = 'macroflow-show-modules';

export function getShowModules() {
  const val = localStorage.getItem(SHOW_MODULES_KEY);
  return val === null ? true : val === 'true';
}

function setShowModules(value) {
  localStorage.setItem(SHOW_MODULES_KEY, String(Boolean(value)));
}

const NAV_GROUPS = [
  {
    label: 'SETTINGS',
    items: [
      { id: 'general', label: 'General', icon: <SlidersIcon size={15} /> },
      { id: 'ai', label: 'AI', icon: <SparkleIcon size={15} /> },
      { id: 'about', label: 'About', icon: <InfoIcon size={15} /> },
    ],
  },
];

function GeneralContent({ theme, toggleTheme, showModules, onToggleShowModules, personalVisible, onTogglePersonalVisibility }) {
  const [chatCleared, setChatCleared] = useState(false);

  const handleClearChats = () => {
    const confirmed = window.confirm('Clear all chat history? This cannot be undone.');
    if (!confirmed) return;
    saveSessions([]);
    setChatCleared(true);
    setTimeout(() => setChatCleared(false), 3000);
  };

  return (
    <div className="settings-section">
      <div className="settings-item">
        <div className="settings-item-info">
          <div className="settings-item-title">Theme</div>
          <div className="settings-item-desc">Switch between light and dark mode</div>
        </div>
        <button className="settings-action-btn" onClick={toggleTheme}>
          {theme === 'dark' ? (
            <><SunIcon size={14} /> Light</>
          ) : (
            <><MoonIcon size={14} /> Dark</>
          )}
        </button>
      </div>
      <div className="settings-item">
        <div className="settings-item-info">
          <div className="settings-item-title">Chat History</div>
          <div className="settings-item-desc">
            {chatCleared ? 'Chat history cleared.' : 'Clear all saved chat sessions'}
          </div>
        </div>
        <button className="settings-action-btn danger" onClick={handleClearChats} disabled={chatCleared}>
          {chatCleared ? 'Cleared' : 'Clear All'}
        </button>
      </div>
      <div className="settings-item">
        <div className="settings-item-info">
          <div className="settings-item-title">Show Modules</div>
          <div className="settings-item-desc">Show VBA modules in the Files explorer</div>
        </div>
        <button
          type="button"
          className={`personal-visibility-switch ${showModules ? 'on' : ''}`}
          onClick={onToggleShowModules}
        >
          <span className="personal-visibility-thumb" />
        </button>
      </div>
      <div className="settings-item">
        <div className="settings-item-info">
          <div className="settings-item-title">Show PERSONAL.XLSB</div>
          <div className="settings-item-desc">Show or hide the PERSONAL.XLSB workbook in Excel</div>
        </div>
        <button
          type="button"
          className={`personal-visibility-switch ${personalVisible ? 'on' : ''}`}
          onClick={onTogglePersonalVisibility}
        >
          <span className="personal-visibility-thumb" />
        </button>
      </div>
    </div>
  );
}

function AboutContent({ updateStatus }) {
  return (
    <div className="settings-section">
      <div className="settings-item">
        <div className="settings-item-info">
          <div className="settings-item-title">Version</div>
          <div className="settings-item-desc">{__APP_VERSION__}</div>
        </div>
      </div>
      <div className="settings-item">
        <div className="settings-item-info">
          <div className="settings-item-title">Updates</div>
          <div className="settings-item-desc">{updateStatus}</div>
        </div>
      </div>
    </div>
  );
}

function AiContent({ aiStatus }) {
  const model = aiStatus?.model || 'claude-sonnet-5';
  const statusLabel = aiStatus?.ready === false ? 'Unavailable' : 'Ready';

  return (
    <div className="settings-section">
      <div className="settings-item">
        <div className="settings-item-info">
          <div className="settings-item-title">Status</div>
          <div className="settings-item-desc">{statusLabel}</div>
        </div>
      </div>
      <div className="settings-item">
        <div className="settings-item-info">
          <div className="settings-item-title">Model</div>
          <div className="settings-item-desc">Claude Sonnet 5</div>
        </div>
      </div>
      <div className="settings-item">
        <div className="settings-item-info">
          <div className="settings-item-title">Provider</div>
          <div className="settings-item-desc">Anthropic via MacroFlow cloud proxy</div>
        </div>
      </div>
      <div className="settings-subcopy">
        Create uses `{model}`. No local model download is required.
      </div>
    </div>
  );
}

function SettingsPanel({ onClose, onShowModulesChange }) {
  const [activeTab, setActiveTab] = useState('general');
  const [theme, setThemeState] = useState(getTheme);
  const [showModulesState, setShowModulesState] = useState(getShowModules);
  const [personalVisible, setPersonalVisible] = useState(false);
  const [updateStatus, setUpdateStatus] = useState('Up to date');
  const [aiStatus, setAiStatus] = useState(null);

  useEffect(() => {
    window.excel?.updater?.checkNow?.().then((result) => {
      if (result?.success) {
        setUpdateStatus('Checking...');
      } else if (result?.reason === 'development mode') {
        setUpdateStatus('Dev mode');
      } else {
        setUpdateStatus('Up to date');
      }
    }).catch(() => {});

    window.excel?.ai?.getStatus?.().then((status) => {
      setAiStatus(status || null);
    }).catch(() => {});

    window.excel?.personal?.getVisibility?.().then((result) => {
      if (result?.success) {
        setPersonalVisible(result.visible === true);
      }
    }).catch(() => {});
  }, []);

  const toggleShowModules = () => {
    const next = !showModulesState;
    setShowModules(next);
    setShowModulesState(next);
    onShowModulesChange?.(next);
  };

  const togglePersonalVisibility = async () => {
    const next = !personalVisible;
    try {
      const result = await window.excel?.personal?.setVisibility?.({ visible: next });
      if (result?.success) {
        setPersonalVisible(next);
      }
    } catch (_) { /* ignore */ }
  };

  const toggleTheme = () => {
    const next = theme === 'dark' ? 'light' : 'dark';
    setTheme(next);
    setThemeState(next);
  };

  const activeLabel = NAV_GROUPS.flatMap((g) => g.items).find((i) => i.id === activeTab)?.label;

  return (
    <>
      <div className="settings-panel-overlay" onClick={onClose} />
      <div className="settings-panel">
        <nav className="settings-panel-sidebar">
          {NAV_GROUPS.map((group) => (
            <div className="settings-nav-group" key={group.label}>
              <div className="settings-nav-group-label">{group.label}</div>
              {group.items.map((item) => (
                <button
                  key={item.id}
                  className={`settings-nav-item${activeTab === item.id ? ' active' : ''}`}
                  onClick={() => setActiveTab(item.id)}
                >
                  <span className="settings-nav-icon">{item.icon}</span>
                  {item.label}
                </button>
              ))}
            </div>
          ))}
          <div className="settings-sidebar-version">MacroFlow v{__APP_VERSION__}</div>
        </nav>

        <div className="settings-panel-main">
          <div className="settings-panel-header">
            <span className="settings-panel-title">{activeLabel}</span>
          </div>
          <div className="settings-panel-content">
            {activeTab === 'general' && (
              <GeneralContent theme={theme} toggleTheme={toggleTheme} showModules={showModulesState} onToggleShowModules={toggleShowModules} personalVisible={personalVisible} onTogglePersonalVisibility={togglePersonalVisibility} />
            )}
            {activeTab === 'about' && (
              <AboutContent updateStatus={updateStatus} />
            )}
            {activeTab === 'ai' && (
              <AiContent aiStatus={aiStatus} />
            )}
          </div>
        </div>
      </div>
    </>
  );
}

const SettingsMenu = ({ isOpen, onClose, onQuit, onShowModulesChange }) => {
  const [panelOpen, setPanelOpen] = useState(false);
  const stats = isOpen ? getMacroStats() : null;

  if (panelOpen) {
    return (
      <SettingsPanel
        onClose={() => {
          setPanelOpen(false);
          onClose();
        }}
        onShowModulesChange={onShowModulesChange}
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
