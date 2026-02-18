import React, { useState, useMemo, useEffect, useCallback, useRef } from 'react';
import {
  FolderIcon,
  FolderIconLarge,
  ReturnIcon,
  ArrowLeftIcon,
  CloseIcon,
  ListIcon,
  ChevronDownIcon,
  WorkbookIcon,
} from './icons';
import {
  buildExplorerTree,
  filterExplorerTree,
  buildNodeMetadata,
  countTreeItems,
  getDefaultExpandedIds,
} from '../features/search/explorer-selectors';
import { usePersonalMacros } from '../features/search/usePersonalMacros';

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
  const [selectedNode, setSelectedNode] = useState(null);
  const [expandedIds, setExpandedIds] = useState(new Set());
  const hasInitializedRef = useRef(false);

  const status = searchData.status || 'idle';
  const personalState = usePersonalMacros(searchData);

  // Build the full tree
  const tree = useMemo(
    () => buildExplorerTree(searchData, personalState),
    [searchData, personalState]
  );

  // Initialize expand state when tree first becomes available
  useEffect(() => {
    if (tree.length > 0 && !hasInitializedRef.current) {
      hasInitializedRef.current = true;
      setExpandedIds(getDefaultExpandedIds(tree));
    }
  }, [tree]);

  // Reset init flag when workbook changes
  useEffect(() => {
    hasInitializedRef.current = false;
  }, [searchData.workbook?.path, searchData.workbook?.name]);

  // Filter tree by search query
  const { filteredTree, matchedIds } = useMemo(
    () => filterExplorerTree(tree, searchQuery, shortcutByMacroId),
    [tree, searchQuery, shortcutByMacroId]
  );

  // During search, auto-expand all matched branches; otherwise use manual state
  const effectiveExpandedIds = useMemo(() => {
    if (searchQuery.trim()) return matchedIds;
    return expandedIds;
  }, [searchQuery, matchedIds, expandedIds]);

  const itemCount = useMemo(() => countTreeItems(filteredTree), [filteredTree]);

  const metadata = useMemo(
    () => buildNodeMetadata(selectedNode),
    [selectedNode]
  );

  // Clear stale selection when tree changes
  useEffect(() => {
    if (!selectedNode) return;
    function findNode(nodes, targetId) {
      for (const node of nodes) {
        if (node.id === targetId) return true;
        if (node.children.length > 0 && findNode(node.children, targetId)) return true;
      }
      return false;
    }
    if (!findNode(tree, selectedNode.id)) setSelectedNode(null);
  }, [tree, selectedNode]);

  const toggleExpand = useCallback((nodeId) => {
    setExpandedIds((prev) => {
      const next = new Set(prev);
      if (next.has(nodeId)) next.delete(nodeId);
      else next.add(nodeId);
      return next;
    });
  }, []);

  // --- Render helpers ---

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

  const renderTreeNode = (node, depth = 0) => {
    const isExpanded = effectiveExpandedIds.has(node.id);
    const isSelected = selectedNode?.id === node.id;
    const hasChildren = node.children.length > 0;
    const paddingLeft = 8 + depth * 16;

    return (
      <React.Fragment key={node.id}>
        <div
          className={`tree-node tree-node--${node.nodeType} ${isSelected ? 'tree-node--selected' : ''}`}
          style={{ paddingLeft }}
          onClick={() => setSelectedNode(node)}
        >
          <span
            className={`tree-node__chevron ${hasChildren ? '' : 'tree-node__chevron--hidden'}`}
            onClick={(e) => {
              e.stopPropagation();
              if (hasChildren) toggleExpand(node.id);
            }}
          >
            {hasChildren && (
              <ChevronDownIcon
                size={14}
                className={`tree-chevron-icon ${isExpanded ? '' : 'tree-chevron-icon--collapsed'}`}
              />
            )}
          </span>

          <span className={`tree-node__icon tree-node__icon--${node.nodeType}`}>
            {node.nodeType === 'workbook' && <WorkbookIcon size={16} />}
            {node.nodeType === 'module' && <FolderIcon size={16} />}
            {node.nodeType === 'macro' && <ReturnIcon size={14} />}
          </span>

          <span className="tree-node__label">{node.label}</span>
        </div>

        {hasChildren && isExpanded && node.children.map((child) => renderTreeNode(child, depth + 1))}
      </React.Fragment>
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
          {/* Left Panel - Tree */}
          <div className="split-left">
            <div className="section-header">
              <span className="section-title">Explorer</span>
              <span className="section-count">{itemCount} items</span>
            </div>

            <div className="tree-list">
              {filteredTree.length === 0 && (
                <div className="search-empty-state">No items match this search.</div>
              )}
              {filteredTree.map((rootNode) => renderTreeNode(rootNode, 0))}
            </div>
          </div>

          {/* Right Panel - Details */}
          <div className="split-right">
            {selectedNode ? (
              <div className="details-panel">
                <div className="folder-icon-large">
                  {selectedNode.nodeType === 'workbook' && <WorkbookIcon size={80} />}
                  {selectedNode.nodeType === 'module' && <FolderIconLarge size={80} />}
                  {selectedNode.nodeType === 'macro' && <ReturnIcon size={60} />}
                </div>
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
