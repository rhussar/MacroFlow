import { useEffect, useRef } from 'react';

/**
 * Right-click context menu for macro rows.
 * Appears at cursor position and provides icon management options.
 */
export default function MacroContextMenu({ x, y, hasIcon, onAssignIcon, onRemoveIcon, onRename, onEdit, onClose }) {
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
        <span>Change Icon...</span>
      </button>
      {onEdit && (
        <button
          type="button"
          className="macro-context-menu-item"
          onClick={() => { onEdit(); onClose(); }}
        >
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
            <path d="M12 2L14.5 9.5L22 12L14.5 14.5L12 22L9.5 14.5L2 12L9.5 9.5L12 2Z" />
          </svg>
          <span>Edit</span>
        </button>
      )}
      {onRename && (
        <button
          type="button"
          className="macro-context-menu-item"
          onClick={() => { onRename(); onClose(); }}
        >
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M17 3a2.85 2.85 0 0 1 4 4L7.5 20.5 2 22l1.5-5.5L17 3z" />
          </svg>
          <span>Rename</span>
        </button>
      )}
      {hasIcon && !onRename && (
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
