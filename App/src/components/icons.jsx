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