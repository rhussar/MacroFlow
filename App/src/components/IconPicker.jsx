import { useState, useMemo, useCallback, useRef, useEffect } from 'react';
import ImageMsoIcon, { getAllIconNames } from './ImageMsoIcon';
import { CloseIcon } from './icons';
import { ICON_CATEGORIES, POPULAR_ICONS } from '../features/icons/imagemso-categories';

const GRID_COLS = 10;
const ICON_CELL_SIZE = 36;
const VISIBLE_HEIGHT = 300;

/**
 * Icon picker modal for assigning ImageMSO icons to macros.
 * Shows a "Popular" tab with curated categories and an "All" tab with search.
 */
export default function IconPicker({ onSelect, onClose, currentIcon }) {
  const [tab, setTab] = useState('popular');
  const [search, setSearch] = useState('');
  const [hoveredIcon, setHoveredIcon] = useState(null);
  const scrollRef = useRef(null);
  const searchRef = useRef(null);

  const allNames = useMemo(() => getAllIconNames(), []);

  // Focus search when switching to "All" tab
  useEffect(() => {
    if (tab === 'all' && searchRef.current) {
      searchRef.current.focus();
    }
  }, [tab]);

  // Close on Escape
  useEffect(() => {
    const handleKeyDown = (e) => {
      if (e.key === 'Escape') onClose();
    };
    document.addEventListener('keydown', handleKeyDown);
    return () => document.removeEventListener('keydown', handleKeyDown);
  }, [onClose]);

  // Filtered list for "All" tab
  const filteredAll = useMemo(() => {
    if (!search.trim()) return allNames;
    const q = search.toLowerCase();
    return allNames.filter(name => name.toLowerCase().includes(q));
  }, [allNames, search]);

  const handleSelect = useCallback((name) => {
    onSelect(name);
    onClose();
  }, [onSelect, onClose]);

  // Virtualized grid for "All" tab
  const VirtualGrid = useCallback(({ items }) => {
    const [scrollTop, setScrollTop] = useState(0);
    const totalRows = Math.ceil(items.length / GRID_COLS);
    const totalHeight = totalRows * ICON_CELL_SIZE;
    const startRow = Math.floor(scrollTop / ICON_CELL_SIZE);
    const visibleRows = Math.ceil(VISIBLE_HEIGHT / ICON_CELL_SIZE) + 2;
    const endRow = Math.min(startRow + visibleRows, totalRows);

    const visibleItems = [];
    for (let row = startRow; row < endRow; row++) {
      for (let col = 0; col < GRID_COLS; col++) {
        const idx = row * GRID_COLS + col;
        if (idx < items.length) {
          visibleItems.push({ name: items[idx], row, col, idx });
        }
      }
    }

    return (
      <div
        className="icon-picker-virtual-scroll"
        style={{ height: VISIBLE_HEIGHT, overflow: 'auto' }}
        onScroll={(e) => setScrollTop(e.currentTarget.scrollTop)}
        ref={scrollRef}
      >
        <div style={{ height: totalHeight, position: 'relative' }}>
          {visibleItems.map(({ name, row, col }) => (
            <button
              key={name}
              type="button"
              className={`icon-picker-cell ${name === currentIcon ? 'icon-picker-cell--selected' : ''}`}
              style={{
                position: 'absolute',
                left: col * ICON_CELL_SIZE,
                top: row * ICON_CELL_SIZE,
                width: ICON_CELL_SIZE,
                height: ICON_CELL_SIZE,
              }}
              onClick={() => handleSelect(name)}
              onMouseEnter={() => setHoveredIcon(name)}
              onMouseLeave={() => setHoveredIcon(null)}
              title={name}
            >
              <ImageMsoIcon name={name} size={20} />
            </button>
          ))}
        </div>
      </div>
    );
  }, [currentIcon, handleSelect]);

  return (
    <div className="icon-picker-backdrop" onClick={onClose}>
      <div className="icon-picker-modal" onClick={(e) => e.stopPropagation()}>
        {/* Header */}
        <div className="icon-picker-header">
          <h3 className="icon-picker-title">Choose an Icon</h3>
          <button type="button" className="icon-picker-close" onClick={onClose} title="Close">
            <CloseIcon size={16} />
          </button>
        </div>

        {/* Tabs */}
        <div className="icon-picker-tabs">
          <button
            type="button"
            className={`icon-picker-tab ${tab === 'popular' ? 'icon-picker-tab--active' : ''}`}
            onClick={() => setTab('popular')}
          >
            Popular
          </button>
          <button
            type="button"
            className={`icon-picker-tab ${tab === 'all' ? 'icon-picker-tab--active' : ''}`}
            onClick={() => setTab('all')}
          >
            All ({allNames.length})
          </button>
        </div>

        {/* Search (All tab only) */}
        {tab === 'all' && (
          <div className="icon-picker-search-wrap">
            <input
              ref={searchRef}
              type="text"
              className="icon-picker-search"
              placeholder="Search icons..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
            />
          </div>
        )}

        {/* Content */}
        <div className="icon-picker-content">
          {tab === 'popular' ? (
            <div className="icon-picker-categories">
              {ICON_CATEGORIES.map((cat) => (
                <div key={cat.label} className="icon-picker-category">
                  <div className="icon-picker-category-label">{cat.label}</div>
                  <div className="icon-picker-grid">
                    {cat.icons.map((name) => (
                      <button
                        key={name}
                        type="button"
                        className={`icon-picker-cell ${name === currentIcon ? 'icon-picker-cell--selected' : ''}`}
                        onClick={() => handleSelect(name)}
                        onMouseEnter={() => setHoveredIcon(name)}
                        onMouseLeave={() => setHoveredIcon(null)}
                        title={name}
                      >
                        <ImageMsoIcon name={name} size={20} />
                      </button>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          ) : (
            <VirtualGrid items={filteredAll} />
          )}
        </div>

        {/* Footer — shows hovered icon name */}
        <div className="icon-picker-footer">
          {hoveredIcon ? (
            <span className="icon-picker-footer-name">{hoveredIcon}</span>
          ) : currentIcon ? (
            <span className="icon-picker-footer-name">Current: {currentIcon}</span>
          ) : (
            <span className="icon-picker-footer-hint">Hover over an icon to see its name</span>
          )}
        </div>
      </div>
    </div>
  );
}
