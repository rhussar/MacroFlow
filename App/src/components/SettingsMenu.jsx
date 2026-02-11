import {
  MailIcon,
  DocumentIcon,
  MacroFlowLogo,
  SettingsIcon,
  ExitIcon,
} from './icons';

const SettingsMenu = ({ isOpen, onClose, onQuit }) => {
  if (!isOpen) return null;

  const menuItems = [
    { icon: <MailIcon />, label: 'Send Feedback', action: () => console.log('Feedback') },
    { icon: <DocumentIcon />, label: 'Documentation', action: () => console.log('Docs') },
    { icon: <MacroFlowLogo size={16} />, label: 'About MacroFlow', action: () => console.log('About') },
    { icon: <SettingsIcon />, label: 'Settings', action: () => console.log('Settings') },
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
              onClose();
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
