import React from 'react';
import appIcon from '../../assets/app-icon.png';
const excelLogo = 'data:image/png;base64,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';
const folderLogo = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAACXBIWXMAAAsTAAALEwEAmpwYAAACTUlEQVR4nO2XwW4SQRjHF9ST9WSMMaVvUGBml1o1FeIuEpOefQVfg3O9mPgKJh5gp9AS02CjtKbKUKvGg00wYXd1sdhUQxOBplDmM7PbNFCg0ABtD0zy29M3M7/9725mP0G4KANAcEBYcPYECI7BCwinW5SLDFSgkPRczcUkdy6KPSeRjXjR5vzUdUsCwoORAAg7dYKX83ERDIL3DRXVOlA1VMx0FeV1It7hc1OpwOW+Bf4kvOOFBVw2bQH2Yx5DJ3SC2fYrCX4tiNvZqDTN529sSFe4SDcg8vhSW4HvxO3SVVz6GRP5BnW+yUloBB8UErZELiLd7vudMLkAQRVLQMXMIBi6oam4bknExd85FT3XCZrTCXragTlNRc8M4p21ZY89NvNIAIMWRUxXMfQCrzVjmBWTPth93Zli0sdKb6Zga1Esbr6QblkSjZ+y2SDwN3mX7a36obLSG+WUn5VX7td6YW/VXwUqrwFVnjRJmMTt0lRU2Vr0QZ0qDDJBGAo0yODbI4CvoV14N3OjKQEuUEj4ANIyA6rAUEgr9toZ5R+shybOR2CdJyGXWSrgOjcBRuUKoyOBwEjANRIYJTBK4KIlgEuHAnWrcDjUrWM5LZeaBPIx30Q+hu0EqMwgYx2Zg4f/a3wMAnwOluHL7PiRwE783jUtirI7S9O8qMqofNCRdCtAlVpvyPuWAJWXG3sKB7/kXk7erL6dkeC9gtuSeYhaWDuEhrw2D5r5JHsgfYwPQTekAmMtbRmcsjXrZ7TtLYFLQNh5NpzdzXYd/wFj1V/Xq1V+gQAAAABJRU5ErkJggg==';

// MacroFlow Logo (app icon)
export const MacroFlowLogo = ({ className = '', size = 20 }) => (
  <img
    src={appIcon}
    alt="MacroFlow"
    width={size}
    height={size}
    className={className}
  />
);

// Folder Icon (uses folder-logo asset)
export const FolderIcon = ({ className = '', size = 24 }) => (
  <img
    src={folderLogo}
    alt="Folder"
    width={size}
    height={size}
    className={className}
  />
);

// Large Folder Icon (same asset, larger default)
export const FolderIconLarge = ({ className = '', size = 80 }) => (
  <img
    src={folderLogo}
    alt="Folder"
    width={size}
    height={size}
    className={className}
  />
);

// Arrow Left (back button / macro icon)
export const ArrowLeftIcon = ({ className = '', size = 20 }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <path d="M19 12H5M12 19l-7-7 7-7" />
  </svg>
);

// Close (X) Icon
export const CloseIcon = ({ className = '', size = 18 }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <path d="M18 6L6 18M6 6l12 12" />
  </svg>
);

// Chevron Down (dropdown)
export const ChevronDownIcon = ({ className = '', size = 16 }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <path d="M6 9l6 6 6-6" />
  </svg>
);

// Check Icon
export const CheckIcon = ({ className = '', size = 14 }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="3"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <path d="M20 6L9 17l-5-5" />
  </svg>
);

// Edit (Pencil) Icon
export const EditIcon = ({ className = '', size = 14 }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7" />
    <path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z" />
  </svg>
);

// Mail/Feedback Icon
export const MailIcon = ({ className = '', size = 16 }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <rect x="2" y="4" width="20" height="16" rx="2" />
    <path d="M22 7l-10 5L2 7" />
  </svg>
);

// Document Icon
export const DocumentIcon = ({ className = '', size = 16 }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z" />
    <path d="M14 2v6h6M16 13H8M16 17H8M10 9H8" />
  </svg>
);

// Settings/Gear Icon
export const SettingsIcon = ({ className = '', size = 16 }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <circle cx="12" cy="12" r="3" />
    <path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-2 2 2 2 0 01-2-2v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83 0 2 2 0 010-2.83l.06-.06a1.65 1.65 0 00.33-1.82 1.65 1.65 0 00-1.51-1H3a2 2 0 01-2-2 2 2 0 012-2h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 010-2.83 2 2 0 012.83 0l.06.06a1.65 1.65 0 001.82.33H9a1.65 1.65 0 001-1.51V3a2 2 0 012-2 2 2 0 012 2v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 0 2 2 0 010 2.83l-.06.06a1.65 1.65 0 00-.33 1.82V9a1.65 1.65 0 001.51 1H21a2 2 0 012 2 2 2 0 01-2 2h-.09a1.65 1.65 0 00-1.51 1z" />
  </svg>
);

// Exit/Logout Icon
export const ExitIcon = ({ className = '', size = 16 }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4" />
    <polyline points="16,17 21,12 16,7" />
    <line x1="21" y1="12" x2="9" y2="12" />
  </svg>
);

// List Icon (for filter)
export const ListIcon = ({ className = '', size = 16 }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <line x1="8" y1="6" x2="21" y2="6" />
    <line x1="8" y1="12" x2="21" y2="12" />
    <line x1="8" y1="18" x2="21" y2="18" />
    <line x1="3" y1="6" x2="3.01" y2="6" />
    <line x1="3" y1="12" x2="3.01" y2="12" />
    <line x1="3" y1="18" x2="3.01" y2="18" />
  </svg>
);

// Refresh/Agent Icon
export const RefreshIcon = ({ className = '', size = 16 }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <path d="M23 4v6h-6" />
    <path d="M1 20v-6h6" />
    <path d="M3.51 9a9 9 0 0114.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0020.49 15" />
  </svg>
);

// Message/Chat Icon
export const MessageIcon = ({ className = '', size = 16 }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <path d="M21 15a2 2 0 01-2 2H7l-4 4V5a2 2 0 012-2h14a2 2 0 012 2z" />
  </svg>
);

// Workbook/Excel Icon (uses Excel logo asset)
export const WorkbookIcon = ({ className = '', size = 16 }) => (
  <img
    src={excelLogo}
    alt="Workbook"
    width={size}
    height={size}
    className={className}
  />
);

// Return/Enter Arrow (for macros)
export const ReturnIcon = ({ className = '', size = 16 }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <polyline points="9,10 4,15 9,20" />
    <path d="M20 4v7a4 4 0 01-4 4H4" />
  </svg>
);

export const SidebarIcon = ({ className = '', size = 16 }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <rect x="3" y="3" width="18" height="18" rx="2" />
    <line x1="9" y1="3" x2="9" y2="21" />
  </svg>
);