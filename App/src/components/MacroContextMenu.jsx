import { useEffect, useRef } from 'react';

/**
 * Right-click context menu for macro rows.
 * Appears at cursor position and provides icon management options.
 */
export default function MacroContextMenu({ x, y, hasIcon, onAssignIcon, onRemoveIcon, onClose }) {
  const menuRef = useRef(null);

  useEffect(() => {
    const handleClickOutside = (e) => {
      if (menuRef.current && !menuRef.current.contains(e.target)) {
        onClose();
      }
    };
    const handleEsc = (e) => {
      if (e.key === 'Escape') onClose();
    };
    document.addEventListener('mousedown', handleClickOutside);
    document.addEventListener('keydown', handleEsc);
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
      document.removeEventListener('keydown', handleEsc);
    };
  }, [onClose]);

  // Keep menu within viewport
  const style = {
    position: 'fixed',
    left: x,
    top: y,
    zIndex: 9999,
  };

  return (
    <div className="macro-context-menu" style={style} ref={menuRef}>
      <button
        type="button"
        className="macro-context-menu-item"
        onClick={() => { onAssignIcon(); onClose(); }}
      >
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
          <rect x="3" y="3" width="18" height="18" rx="2" ry="2" />
          <circle cx="8.5" cy="8.5" r="1.5" />
          <polyline points="21,15 16,10 5,21" />
        </svg>
        <span>Assign Icon...</span>
      </button>
      {hasIcon && (
        <button
          type="button"
          className="macro-context-menu-item macro-context-menu-item--danger"
          onClick={() => { onRemoveIcon(); onClose(); }}
        >
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <line x1="18" y1="6" x2="6" y2="18" />
            <line x1="6" y1="6" x2="18" y2="18" />
          </svg>
          <span>Remove Icon</span>
        </button>
      )}
    </div>
  );
}
