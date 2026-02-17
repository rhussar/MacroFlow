import React, { useState, useMemo, useEffect } from 'react';
import {
  FolderIcon,
  FolderIconLarge,
  ReturnIcon,
  ArrowLeftIcon,
  CloseIcon,
  ListIcon,
  ChevronDownIcon,
} from './icons';
import {
  filterExplorerItems,
  buildItemMetadata,
} from '../features/search/explorer-selectors';

const defaultSearchData = {
  status: 'idle',
  workbook: null,
  modules: [],
  macros: [],
  error: null,
};

const statusConfig = {
  idle: {
    title: 'Loading workbook data',
    message: 'Connecting to the active Excel workbook.',
  },
  loading: {
    title: 'Loading workbook data',
    message: 'Refreshing modules and macros from Excel.',
  },
  no_excel: {
    title: 'Excel is not running',
    message: 'Open Excel. MacroFlow will retry automatically.',
  },
  no_workbook: {
    title: 'No active workbook',
    message: 'Open or create a workbook. MacroFlow will retry automatically.',
  },
  error: {
    title: 'Could not load workbook data',
    message: 'Something went wrong while reading workbook data. Retrying automatically.',
  },
};

const FileExplorer = ({
  onBack,
  onClose,
  searchData = defaultSearchData,
  shortcutByMacroId = {},
}) => {
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedItem, setSelectedItem] = useState(null);

  const status = searchData.status || 'idle';

  const filteredItems = useMemo(
    () => filterExplorerItems(searchData.modules, searchData.macros, searchQuery, shortcutByMacroId),
    [searchData.modules, searchData.macros, searchQuery, shortcutByMacroId]
  );

  const metadata = useMemo(
    () => buildItemMetadata(selectedItem),
    [selectedItem]
  );

  // Clear selection when the selected item no longer exists in the data
  useEffect(() => {
    if (!selectedItem) return;
    const allItems = [...(searchData.modules || []), ...(searchData.macros || [])];
    const stillExists = allItems.some((item) => item.id === selectedItem.id);
    if (!stillExists) setSelectedItem(null);
  }, [searchData.modules, searchData.macros, selectedItem]);

  const isModule = selectedItem?.itemType === 'module';

  const renderNonReadyState = () => {
    const config = statusConfig[status] || statusConfig.error;
    const message = status === 'error' && searchData.error?.message
      ? searchData.error.message
      : config.message;
    const isLoadingState = status === 'idle' || status === 'loading';

    return (
      <div className={`search-status-panel search-status-${status}`}>
        <div className="search-status-header">
          {isLoadingState && <span className="status-spinner" />}
          <span className="search-status-title">{config.title}</span>
        </div>
        <p className="search-status-message">{message}</p>
      </div>
    );
  };

  return (
    <>
      {/* Header */}
      <header className="header">
        <button className="header-back-btn" onClick={onBack}>
          <ArrowLeftIcon size={20} />
        </button>

        <div className="search-input-wrapper">
          <input
            type="text"
            className="search-input"
            placeholder="Search files..."
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
          />
        </div>

        <div className="header-actions">
          <button className="filter-dropdown">
            <ListIcon size={14} />
            All files
            <ChevronDownIcon size={14} />
          </button>
          <button className="close-btn" onClick={onClose}>
            <CloseIcon />
          </button>
        </div>
      </header>

      {/* Main area */}
      {status !== 'ready' ? (
        <main className="main-content">
          {renderNonReadyState()}
        </main>
      ) : (
        <div className="split-view">
          {/* Left Panel - File List */}
          <div className="split-left">
            <div className="section-header">
              <span className="section-title">All Files</span>
              <span className="section-count">{filteredItems.length} items</span>
            </div>

            <div className="file-list">
              {filteredItems.length === 0 && (
                <div className="search-empty-state">No items match this search.</div>
              )}
              {filteredItems.map((item) => (
                <div
                  key={item.id}
                  className={`file-item ${selectedItem?.id === item.id ? 'selected' : ''}`}
                  onClick={() => setSelectedItem(item)}
                >
                  <div className={`file-icon ${item.itemType === 'macro' ? 'macro' : ''}`}>
                    {item.itemType === 'module' ? (
                      <FolderIcon size={20} />
                    ) : (
                      <ReturnIcon size={16} />
                    )}
                  </div>
                  <span className="file-name">{item.name}</span>
                </div>
              ))}
            </div>
          </div>

          {/* Right Panel - Details */}
          <div className="split-right">
            {selectedItem ? (
              <div className="details-panel">
                {isModule && (
                  <div className="folder-icon-large">
                    <FolderIconLarge size={80} />
                  </div>
                )}
                <div className="metadata-section">
                  <div className="metadata-title">Metadata</div>
                  {metadata.rows.map((row) => (
                    <div className="metadata-row" key={row.label}>
                      <span className="metadata-label">{row.label}</span>
                      <span className="metadata-value">{row.value}</span>
                    </div>
                  ))}
                </div>
              </div>
            ) : (
              <div className="details-panel">
                <p className="search-status-message">Select an item to view details.</p>
              </div>
            )}
          </div>
        </div>
      )}
    </>
  );
};

export default FileExplorer;
