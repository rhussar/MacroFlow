import React, { useState } from 'react';
import {
  MailIcon,
  DocumentIcon,
  MacroFlowLogo,
  SettingsIcon,
  ExitIcon,
  SunIcon,
  MoonIcon,
} from './icons';

function getTheme() {
  return document.documentElement.getAttribute('data-theme') || 'dark';
}

function setTheme(theme) {
  document.documentElement.setAttribute('data-theme', theme);
  localStorage.setItem('macroflow-theme', theme);
}

const SettingsMenu = ({ isOpen, onClose, onQuit }) => {
  const [theme, setThemeState] = useState(getTheme);

  if (!isOpen) return null;
  const noAction = () => {};

  const toggleTheme = () => {
    const next = theme === 'dark' ? 'light' : 'dark';
    setTheme(next);
    setThemeState(next);
  };

  const menuItems = [
    { icon: <MailIcon />, label: 'Send Feedback', action: noAction },
    { icon: <DocumentIcon />, label: 'Documentation', action: noAction },
    { icon: <MacroFlowLogo size={16} />, label: 'About MacroFlow', action: noAction },
    { icon: <SettingsIcon />, label: 'Settings', action: noAction },
    {
      icon: theme === 'dark' ? <SunIcon /> : <MoonIcon />,
      label: theme === 'dark' ? 'Light Mode' : 'Dark Mode',
      action: toggleTheme,
      keepOpen: true,
    },
    { icon: <ExitIcon />, label: 'Quit MacroFlow', action: onQuit, danger: true },
  ];

  return (
    <>
      {/* Overlay to catch clicks outside */}
      <div className="settings-menu-overlay" onClick={onClose} />

      {/* Menu */}
      <div className="settings-menu">
        <div className="settings-menu-header">MacroFlow v0.01.2.0</div>
        {menuItems.map((item, index) => (
          <div
            key={index}
            className={`settings-menu-item ${item.danger ? 'danger' : ''}`}
            onClick={() => {
              item.action?.();
              if (!item.keepOpen) onClose();
            }}
          >
            <span className="settings-menu-icon">{item.icon}</span>
            <span>{item.label}</span>
          </div>
        ))}
      </div>
    </>
  );
};

export default SettingsMenu;
