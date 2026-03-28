import React from 'react';
import appIcon from '../../assets/app-icon.png';
import lightningIcon from '../../assets/lightning.png';


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

// Folder Icon (inline SVG)
export const FolderIcon = ({ className = '', size = 24 }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="1.5"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z" />
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
    viewBox="0 0 16 16"
    fill="none"
    xmlns="http://www.w3.org/2000/svg"
    className={className}
  >
    <g clipPath="url(#settings-clip)">
      <path d="M8 10C9.10457 10 10 9.10459 10 8.00002C10 6.89545 9.10457 6.00002 8 6.00002C6.89543 6.00002 6 6.89545 6 8.00002C6 9.10459 6.89543 10 8 10Z" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
      <path d="M12.9333 10C12.8446 10.2011 12.8181 10.4241 12.8573 10.6404C12.8965 10.8567 12.9996 11.0562 13.1533 11.2134L13.1933 11.2534C13.3173 11.3772 13.4156 11.5242 13.4827 11.6861C13.5498 11.848 13.5844 12.0215 13.5844 12.1967C13.5844 12.3719 13.5498 12.5454 13.4827 12.7073C13.4156 12.8691 13.3173 13.0162 13.1933 13.14C13.0695 13.264 12.9225 13.3623 12.7606 13.4294C12.5987 13.4965 12.4252 13.5311 12.25 13.5311C12.0748 13.5311 11.9013 13.4965 11.7394 13.4294C11.5775 13.3623 11.4305 13.264 11.3067 13.14L11.2667 13.1C11.1095 12.9463 10.91 12.8432 10.6937 12.804C10.4775 12.7648 10.2544 12.7913 10.0533 12.88C9.85615 12.9645 9.68799 13.1048 9.56954 13.2837C9.45109 13.4626 9.38752 13.6722 9.38667 13.8867V14C9.38667 14.3536 9.24619 14.6928 8.99614 14.9428C8.74609 15.1929 8.40696 15.3334 8.05333 15.3334C7.69971 15.3334 7.36057 15.1929 7.11052 14.9428C6.86048 14.6928 6.72 14.3536 6.72 14V13.94C6.71484 13.7194 6.64341 13.5053 6.51501 13.3258C6.3866 13.1463 6.20716 13.0095 6 12.9334C5.79892 12.8446 5.57587 12.8181 5.35961 12.8573C5.14334 12.8966 4.94379 12.9997 4.78667 13.1534L4.74667 13.1934C4.62284 13.3173 4.47578 13.4157 4.31392 13.4828C4.15206 13.5499 3.97855 13.5844 3.80333 13.5844C3.62811 13.5844 3.45461 13.5499 3.29275 13.4828C3.13088 13.4157 2.98383 13.3173 2.86 13.1934C2.73603 13.0695 2.63769 12.9225 2.57059 12.7606C2.50349 12.5987 2.46895 12.4252 2.46895 12.25C2.46895 12.0748 2.50349 11.9013 2.57059 11.7394C2.63769 11.5776 2.73603 11.4305 2.86 11.3067L2.9 11.2667C3.05369 11.1096 3.15679 10.91 3.196 10.6937C3.23522 10.4775 3.20874 10.2544 3.12 10.0534C3.03549 9.85617 2.89517 9.68801 2.71631 9.56956C2.53745 9.45111 2.32786 9.38754 2.11333 9.38669H2C1.64638 9.38669 1.30724 9.24621 1.05719 8.99616C0.807143 8.74611 0.666667 8.40698 0.666667 8.05335C0.666667 7.69973 0.807143 7.36059 1.05719 7.11054C1.30724 6.8605 1.64638 6.72002 2 6.72002H2.06C2.28066 6.71486 2.49467 6.64343 2.6742 6.51503C2.85373 6.38662 2.99048 6.20718 3.06667 6.00002C3.15541 5.79894 3.18188 5.57589 3.14267 5.35963C3.10346 5.14336 3.00036 4.94381 2.84667 4.78669L2.80667 4.74669C2.6827 4.62286 2.58435 4.4758 2.51725 4.31394C2.45015 4.15208 2.41562 3.97857 2.41562 3.80335C2.41562 3.62813 2.45015 3.45463 2.51725 3.29277C2.58435 3.1309 2.6827 2.98385 2.80667 2.86002C2.9305 2.73605 3.07755 2.63771 3.23941 2.57061C3.40128 2.50351 3.57478 2.46897 3.75 2.46897C3.92522 2.46897 4.09872 2.50351 4.26059 2.57061C4.42245 2.63771 4.5695 2.73605 4.69333 2.86002L4.73333 2.90002C4.89045 3.05371 5.09001 3.15681 5.30627 3.19602C5.52254 3.23524 5.74559 3.20876 5.94667 3.12002H6C6.19718 3.03551 6.36534 2.89519 6.4838 2.71633C6.60225 2.53747 6.66581 2.32788 6.66667 2.11335V2.00002C6.66667 1.6464 6.80714 1.30726 7.05719 1.05721C7.30724 0.807163 7.64638 0.666687 8 0.666687C8.35362 0.666687 8.69276 0.807163 8.94281 1.05721C9.19286 1.30726 9.33333 1.6464 9.33333 2.00002V2.06002C9.33419 2.27454 9.39775 2.48414 9.5162 2.663C9.63466 2.84186 9.80282 2.98218 10 3.06669C10.2011 3.15543 10.4241 3.1819 10.6404 3.14269C10.8567 3.10348 11.0562 3.00038 11.2133 2.84669L11.2533 2.80669C11.3772 2.68272 11.5242 2.58437 11.6861 2.51727C11.8479 2.45017 12.0214 2.41564 12.1967 2.41564C12.3719 2.41564 12.5454 2.45017 12.7073 2.51727C12.8691 2.58437 13.0162 2.68272 13.14 2.80669C13.264 2.93052 13.3623 3.07757 13.4294 3.23943C13.4965 3.4013 13.531 3.5748 13.531 3.75002C13.531 3.92524 13.4965 4.09874 13.4294 4.26061C13.3623 4.42247 13.264 4.56952 13.14 4.69335L13.1 4.73335C12.9463 4.89047 12.8432 5.09003 12.804 5.30629C12.7648 5.52256 12.7913 5.74561 12.88 5.94669V6.00002C12.9645 6.1972 13.1048 6.36536 13.2837 6.48382C13.4625 6.60227 13.6721 6.66583 13.8867 6.66669H14C14.3536 6.66669 14.6928 6.80716 14.9428 7.05721C15.1929 7.30726 15.3333 7.6464 15.3333 8.00002C15.3333 8.35364 15.1929 8.69278 14.9428 8.94283C14.6928 9.19288 14.3536 9.33335 14 9.33335H13.94C13.7255 9.33421 13.5159 9.39777 13.337 9.51622C13.1582 9.63468 13.0178 9.80284 12.9333 10Z" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
    </g>
    <defs>
      <clipPath id="settings-clip">
        <rect width="16" height="16" fill="white"/>
      </clipPath>
    </defs>
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

// Workbook/Excel Icon (inline SVG)
export const WorkbookIcon = ({ className = '', size = 16 }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="1.5"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
    <polyline points="14 2 14 8 20 8" />
    <line x1="16" y1="13" x2="8" y2="13" />
    <line x1="16" y1="17" x2="8" y2="17" />
    <polyline points="10 9 9 9 8 9" />
  </svg>
);

// Lightning bolt icon (for macros)
export const ReturnIcon = ({ className = '', size = 16 }) => (
  <img
    src={lightningIcon}
    alt="macro"
    width={size}
    height={size}
    loading="eager"
    fetchPriority="high"
    className={className}
  />
);

// Info Icon (circled "i")
export const InfoIcon = ({ className = '', size = 16 }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 12 12"
    fill="none"
    xmlns="http://www.w3.org/2000/svg"
    className={className}
  >
    <path
      d="M6 8V6M6 4H6.005M11 6C11 8.76142 8.76142 11 6 11C3.23858 11 1 8.76142 1 6C1 3.23858 3.23858 1 6 1C8.76142 1 11 3.23858 11 6Z"
      stroke="currentColor"
      strokeWidth="1.2"
      strokeLinecap="round"
      strokeLinejoin="round"
    />
  </svg>
);

export const SunIcon = ({ className = '', size = 16 }) => (
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
    <circle cx="12" cy="12" r="5" />
    <line x1="12" y1="1" x2="12" y2="3" />
    <line x1="12" y1="21" x2="12" y2="23" />
    <line x1="4.22" y1="4.22" x2="5.64" y2="5.64" />
    <line x1="18.36" y1="18.36" x2="19.78" y2="19.78" />
    <line x1="1" y1="12" x2="3" y2="12" />
    <line x1="21" y1="12" x2="23" y2="12" />
    <line x1="4.22" y1="19.78" x2="5.64" y2="18.36" />
    <line x1="18.36" y1="5.64" x2="19.78" y2="4.22" />
  </svg>
);

export const MoonIcon = ({ className = '', size = 16 }) => (
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
    <path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z" />
  </svg>
);

// Minimize Icon (window control)
export const MinimizeIcon = ({ className = '', size = 16 }) => (
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
    <line x1="5" y1="12" x2="19" y2="12" />
  </svg>
);

export const SidebarIcon = ({ className = '', size = 16 }) => (
  <svg
    width={size}
    height={size * 16 / 21}
    viewBox="0 0 21 16"
    fill="none"
    stroke="currentColor"
    strokeWidth="1.3"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <path d="M7.16667 0.5V15.5M2.72222 0.5H18.2778C19.5051 0.5 20.5 1.24619 20.5 2.16667V13.8333C20.5 14.7538 19.5051 15.5 18.2778 15.5H2.72222C1.49492 15.5 0.5 14.7538 0.5 13.8333V2.16667C0.5 1.24619 1.49492 0.5 2.72222 0.5Z" />
  </svg>
);

// User Icon (person silhouette)
export const UserIcon = ({ className = '', size = 16 }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="1.5"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2" />
    <circle cx="12" cy="7" r="4" />
  </svg>
);

// CPU / Chip Icon (for Local AI)
export const CpuIcon = ({ className = '', size = 16 }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="1.5"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <rect x="4" y="4" width="16" height="16" rx="2" />
    <rect x="9" y="9" width="6" height="6" />
    <line x1="9" y1="1" x2="9" y2="4" />
    <line x1="15" y1="1" x2="15" y2="4" />
    <line x1="9" y1="20" x2="9" y2="23" />
    <line x1="15" y1="20" x2="15" y2="23" />
    <line x1="20" y1="9" x2="23" y2="9" />
    <line x1="20" y1="15" x2="23" y2="15" />
    <line x1="1" y1="9" x2="4" y2="9" />
    <line x1="1" y1="15" x2="4" y2="15" />
  </svg>
);

// Sliders Icon (horizontal controls)
export const SlidersIcon = ({ className = '', size = 16 }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="1.5"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <line x1="4" y1="21" x2="4" y2="14" />
    <line x1="4" y1="10" x2="4" y2="3" />
    <line x1="12" y1="21" x2="12" y2="12" />
    <line x1="12" y1="8" x2="12" y2="3" />
    <line x1="20" y1="21" x2="20" y2="16" />
    <line x1="20" y1="12" x2="20" y2="3" />
    <line x1="1" y1="14" x2="7" y2="14" />
    <line x1="9" y1="8" x2="15" y2="8" />
    <line x1="17" y1="16" x2="23" y2="16" />
  </svg>
);

// Sparkle Icon (for AI)
export const SparkleIcon = ({ className = '', size = 16 }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="1.5"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <path d="M12 2L14.5 9.5L22 12L14.5 14.5L12 22L9.5 14.5L2 12L9.5 9.5L12 2Z" />
  </svg>
);
