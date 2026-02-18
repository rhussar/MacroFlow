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
import { getSearchStatusView } from '../features/search/search-selectors';
import { usePersonalMacros } from '../features/search/usePersonalMacros';
import { resolveInitialModuleNode, useExplorerWorkbookData } from '../features/search/useExplorerWorkbookData';

const defaultSearchData = {
  status: 'idle',
  workbook: null,
  modules: [],
  macros: [],
  error: null,
};

const FileExplorer = ({
  onBack,
  onClose,
  searchData = defaultSearchData,
  shortcutByMacroId = {},
  explorerContext = null,
  onExplorerContextConsumed
}) => {
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedNode, setSelectedNode] = useState(null);
  const [expandedIds, setExpandedIds] = useState(new Set());
  const hasInitializedRef = useRef(false);
  const consumedContextRef = useRef('');
  const resolvedSearchData = useExplorerWorkbookData(searchData, explorerContext);
  const status = resolvedSearchData?.status || 'idle';
  const personalState = usePersonalMacros(resolvedSearchData);

  // Build the full tree
  const tree = useMemo(
    () => buildExplorerTree(resolvedSearchData, personalState),
    [resolvedSearchData, personalState]
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
  }, [resolvedSearchData?.workbook?.path, resolvedSearchData?.workbook?.name]);

  useEffect(() => {
    const workbookKey = String(explorerContext?.workbook?.key || explorerContext?.workbook?.path || explorerContext?.workbook?.name || '').trim();
    const moduleId = String(explorerContext?.initialModuleId || '').trim();
    const moduleName = String(explorerContext?.initialModuleName || '').trim();
    const contextSignature = `${workbookKey}::${moduleId}::${moduleName}`;
    if (!workbookKey || consumedContextRef.current === contextSignature || tree.length === 0 || status !== 'ready') {
      return;
    }

    const targetNode = resolveInitialModuleNode(tree, {
      moduleId,
      moduleName
    });

    const findFirstNodeByType = (nodes, nodeType) => {
      for (const node of nodes) {
        if (node?.nodeType === nodeType) {
          return node;
        }
        if (Array.isArray(node?.children) && node.children.length > 0) {
          const childMatch = findFirstNodeByType(node.children, nodeType);
          if (childMatch) {
            return childMatch;
          }
        }
      }
      return null;
    };

    const fallbackNode = findFirstNodeByType(tree, 'module') || findFirstNodeByType(tree, 'macro') || null;
    setSelectedNode(targetNode || fallbackNode || null);
    consumedContextRef.current = contextSignature;
    onExplorerContextConsumed?.();
  }, [
    explorerContext?.initialModuleId,
    explorerContext?.initialModuleName,
    explorerContext?.workbook?.key,
    explorerContext?.workbook?.name,
    explorerContext?.workbook?.path,
    onExplorerContextConsumed,
    status,
    tree
  ]);

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
    const statusView = getSearchStatusView(resolvedSearchData);

    return (
      <div className={`search-status-panel search-status-${statusView.status}`}>
        <div className="search-status-header">
          {statusView.isLoading && <span className="status-spinner" />}
          <span className="search-status-title">{statusView.title}</span>
        </div>
        <p className="search-status-message">{statusView.message}</p>
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
