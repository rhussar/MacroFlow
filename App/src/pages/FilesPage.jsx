import React, { useState, useMemo, useEffect, useCallback, useRef } from 'react';
import {
  FolderIcon,
  ReturnIcon,
  ChevronDownIcon,
  WorkbookIcon,
} from '../components/icons';
import ImageMsoIcon, { isSpriteReady } from '../components/ImageMsoIcon';
import { getMacroIcon } from '../features/icons/macroIconStore';
import CodePreview from '../components/CodePreview';
import SplitDivider from '../components/SplitDivider';
import {
  buildExplorerTree,
  buildNodeMetadata,
  countTreeItems,
  getDefaultExpandedIds,
} from '../features/search/explorer-selectors';
import { getSearchStatusView } from '../features/search/search-selectors';
import {
  buildWorkbookInvalidationDescriptors,
  invalidateSearchBuckets
} from '../features/search/search-invalidation';
import { PERSONAL_WORKBOOK_NAME, usePersonalMacros } from '../features/search/usePersonalMacros';
import { useExplorerAllFilesData } from '../features/search/useExplorerAllFilesData';
import { resolveInitialModuleNode } from '../features/search/explorer-selection';
import {
  buildWorkbookModuleRequest,
  buildMacroRenameRequest,
  canShowModuleContextActions,
  canShowMacroContextActions,
  isValidVbaModuleName,
  shouldCommitModuleRename,
  displayMacroName,
  encodeMacroName
} from '../features/search/module-actions';

const defaultSearchData = {
  status: 'idle',
  workbook: null,
  modules: [],
  macros: [],
  error: null,
};

function getNodeWorkbookTarget(node, fallbackWorkbook = null) {
  return {
    name: String(node?.data?.workbookName || fallbackWorkbook?.name || '').trim(),
    path: String(node?.data?.workbookPath || fallbackWorkbook?.path || '').trim()
  };
}

function isSameWorkbookTarget(target, workbook) {
  const targetName = String(target?.workbookName || target?.name || '').trim().toLowerCase();
  const targetPath = String(target?.workbookPath || target?.path || '').trim().toLowerCase();
  const workbookName = String(workbook?.name || '').trim().toLowerCase();
  const workbookPath = String(workbook?.path || '').trim().toLowerCase();

  if (targetPath && workbookPath) {
    return targetPath === workbookPath;
  }

  if (targetName && workbookName) {
    return targetName === workbookName;
  }

  return false;
}

function isPersonalWorkbook(node) {
  if (node?.nodeType !== 'workbook') return false;
  const name = String(node?.data?.name || node?.label || '').trim().toUpperCase();
  return name === 'PERSONAL.XLSB';
}

const FilesPage = ({
  searchData = defaultSearchData,
  explorerContext = null,
  onExplorerContextConsumed,
  onActionStatus,
  sidebarOpen = true,
  onEditModule
}) => {
  const [selectedNode, setSelectedNode] = useState(null);
  const [expandedIds, setExpandedIds] = useState(new Set());
  const [moduleContextMenu, setModuleContextMenu] = useState(null);
  const [moduleRenameState, setModuleRenameState] = useState(null);
  const [moduleDeleteTarget, setModuleDeleteTarget] = useState(null);
  const [moduleActionInFlight, setModuleActionInFlight] = useState(false);
  const [previewCode, setPreviewCode] = useState(null);
  const [metadataRename, setMetadataRename] = useState(null);
  const [splitPct, setSplitPct] = useState(45);
  const hasInitializedRef = useRef(false);
  const consumedContextRef = useRef('');
  const moduleContextMenuRef = useRef(null);
  const renameCommitInFlightRef = useRef(false);
  const clickTimerRef = useRef(null);
  const status = searchData?.status || 'idle';
  const {
    workbooks: explorerWorkbooks,
    modules: explorerModules,
    workbookListSignature
  } = useExplorerAllFilesData(searchData);
  const personalState = usePersonalMacros(searchData, workbookListSignature, {
    includeShortcutAudit: false,
    focusRefreshPolicy: 'stale',
    visibilityRefreshPolicy: 'stale'
  });

  // Build the full tree
  const tree = useMemo(
    () => buildExplorerTree({
      searchData,
      personalState,
      workbooks: explorerWorkbooks,
      allFilesModules: explorerModules
    }),
    [explorerModules, explorerWorkbooks, personalState, searchData]
  );

  // Initialize expand state when tree first becomes available
  useEffect(() => {
    if (tree.length > 0 && !hasInitializedRef.current) {
      hasInitializedRef.current = true;
      setExpandedIds(getDefaultExpandedIds());
    }
  }, [tree]);

  // Reset init flag when workbook changes
  useEffect(() => {
    hasInitializedRef.current = false;
  }, [searchData?.workbook?.path, searchData?.workbook?.name]);

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

  const itemCount = useMemo(() => countTreeItems(tree), [tree]);

  const metadata = useMemo(
    () => buildNodeMetadata(selectedNode),
    [selectedNode]
  );

  const previewTarget = useMemo(() => {
    if (!selectedNode || selectedNode.nodeType === 'workbook') {
      return null;
    }

    const moduleName = selectedNode.nodeType === 'module'
      ? String(selectedNode.data?.name || selectedNode.label || '').trim()
      : String(selectedNode.data?.module || '').trim();
    if (!moduleName) {
      return null;
    }

    const targetWorkbook = getNodeWorkbookTarget(selectedNode, searchData?.workbook);
    return {
      moduleName,
      workbookName: targetWorkbook.name,
      workbookPath: targetWorkbook.path
    };
  }, [searchData?.workbook?.name, searchData?.workbook?.path, selectedNode]);

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

  // Fetch module code when a module or macro is selected
  useEffect(() => {
    if (!previewTarget) {
      setPreviewCode(null);
      return;
    }

    const moduleCodeApi = window.excel?.vba?.moduleCodeByWorkbook;
    if (typeof moduleCodeApi !== 'function') {
      setPreviewCode(null);
      return;
    }

    let cancelled = false;
    moduleCodeApi(previewTarget).then((result) => {
      if (cancelled) return;
      if (result?.success && result?.code) {
        setPreviewCode(String(result.code));
      } else {
        setPreviewCode(null);
      }
    }).catch(() => {
      if (!cancelled) setPreviewCode(null);
    });

    return () => { cancelled = true; };
  }, [previewTarget]);

  const toggleExpand = useCallback((nodeId) => {
    setExpandedIds((prev) => {
      const next = new Set(prev);
      if (next.has(nodeId)) next.delete(nodeId);
      else next.add(nodeId);
      return next;
    });
  }, []);

  const closeModuleContextMenu = useCallback(() => {
    setModuleContextMenu(null);
  }, []);

  const invalidateWorkbookMutation = useCallback((targetWorkbook = null, options = {}) => {
    const workbookTarget = {
      name: String(targetWorkbook?.workbookName || targetWorkbook?.name || '').trim(),
      path: String(targetWorkbook?.workbookPath || targetWorkbook?.path || '').trim()
    };
    const affectsActiveWorkbook = isSameWorkbookTarget(workbookTarget, searchData?.workbook);
    const affectsPersonalWorkbook =
      String(workbookTarget.name || '').trim().toUpperCase() === PERSONAL_WORKBOOK_NAME;

    invalidateSearchBuckets(buildWorkbookInvalidationDescriptors({
      workbook: workbookTarget,
      includeActiveWorkbook: options.includeActiveWorkbook === true && affectsActiveWorkbook,
      includeWorkbookList: options.includeWorkbookList === true,
      includeExplorerAllFiles: options.includeExplorerAllFiles === true,
      includePersonalMacros: options.includePersonalMacros === true && affectsPersonalWorkbook,
      includeShortcutAudit: options.includeShortcutAudit === true,
      includeWorkbookScopedData: options.includeWorkbookScopedData === true
    }));
  }, [searchData?.workbook]);

  const handleOpenContextMenu = useCallback((event, node) => {
    event.preventDefault();
    event.stopPropagation();

    const item = node?.data || null;
    if (node.nodeType === 'module' && canShowModuleContextActions(item)) {
      setSelectedNode(node);
      setModuleContextMenu({
        x: Number(event.clientX) || 0,
        y: Number(event.clientY) || 0,
        nodeId: String(node?.id || ''),
        nodeType: 'module',
        module: item
      });
    } else if (node.nodeType === 'macro' && canShowMacroContextActions(item)) {
      setSelectedNode(node);
      setModuleContextMenu({
        x: Number(event.clientX) || 0,
        y: Number(event.clientY) || 0,
        nodeId: String(node?.id || ''),
        nodeType: 'macro',
        macro: item
      });
    } else {
      setModuleContextMenu(null);
    }
  }, []);

  const handleStartRenameModule = useCallback((target) => {
    const moduleItem = target?.module || null;
    if (!moduleItem) {
      return;
    }
    setModuleRenameState({
      nodeId: String(target?.nodeId || ''),
      type: 'module',
      module: moduleItem,
      draft: displayMacroName(String(moduleItem?.name || ''))
    });
    setModuleContextMenu(null);
  }, []);

  const handleStartRenameMacro = useCallback((target) => {
    const macroItem = target?.macro || null;
    if (!macroItem) {
      return;
    }
    setModuleRenameState({
      nodeId: String(target?.nodeId || ''),
      type: 'macro',
      macro: macroItem,
      draft: displayMacroName(String(macroItem?.name || ''))
    });
    setModuleContextMenu(null);
  }, []);

  const handleEditModule = useCallback((target) => {
    const moduleItem = target?.module || null;
    if (!moduleItem || typeof onEditModule !== 'function') return;
    setModuleContextMenu(null);
    const workbook = {
      name: String(moduleItem?.workbookName || '').trim(),
      path: String(moduleItem?.workbookPath || '').trim(),
      key: String(moduleItem?.workbookPath || moduleItem?.workbookName || '').trim()
    };
    onEditModule(workbook, {
      mode: 'existing_module',
      moduleName: moduleItem.name,
      source: 'all-files',
      originMode: 'files'
    });
  }, [onEditModule]);

  const handleRenameDraftChange = useCallback((nextDraft) => {
    setModuleRenameState((previous) => {
      if (!previous) {
        return previous;
      }
      return {
        ...previous,
        draft: String(nextDraft || '')
      };
    });
  }, []);

  const handleCancelRenameModule = useCallback(() => {
    setModuleRenameState(null);
  }, []);

  const handleCommitRenameModule = useCallback(async () => {
    if (renameCommitInFlightRef.current) {
      return;
    }

    const currentRename = moduleRenameState;
    if (!currentRename?.module) {
      return;
    }

    const currentName = String(currentRename.module.name || '').trim();
    const nextName = encodeMacroName(String(currentRename.draft || '').trim());
    if (!shouldCommitModuleRename({ currentName, nextName })) {
      setModuleRenameState(null);
      return;
    }

    if (!isValidVbaModuleName(nextName)) {
      onActionStatus?.('error', 'Invalid name.');
      return;
    }

    const renameApi = window.excel?.vba?.renameModuleByWorkbook;
    if (typeof renameApi !== 'function') {
      onActionStatus?.('error', 'Rename unavailable.');
      return;
    }

    const request = buildWorkbookModuleRequest(currentRename.module, searchData?.workbook);
    if (!request.workbookName && !request.workbookPath) {
      onActionStatus?.('error', 'Workbook not found.');
      return;
    }

    renameCommitInFlightRef.current = true;
    setModuleActionInFlight(true);
    onActionStatus?.('running', 'Renaming...');
    try {
      const result = await renameApi({
        ...request,
        nextModuleName: nextName
      });
      if (!result?.success) {
        const message = String(result?.message || 'Unable to rename module.');
        onActionStatus?.('error', message);
        return;
      }

      setModuleRenameState(null);
      invalidateWorkbookMutation(request, {
        includeActiveWorkbook: true,
        includeExplorerAllFiles: true,
        includePersonalMacros: true,
        includeShortcutAudit: true,
        includeWorkbookScopedData: true
      });
      onActionStatus?.('success', `Renamed to "${nextName}".`);
    } catch (error) {
      const message = error?.message ? String(error.message) : 'Unable to rename module.';
      onActionStatus?.('error', message);
    } finally {
      setModuleActionInFlight(false);
      renameCommitInFlightRef.current = false;
    }
  }, [invalidateWorkbookMutation, moduleRenameState, onActionStatus, searchData?.workbook]);

  const handleCommitRenameMacro = useCallback(async () => {
    if (renameCommitInFlightRef.current) {
      return;
    }

    const currentRename = moduleRenameState;
    if (!currentRename?.macro || currentRename?.type !== 'macro') {
      return;
    }

    const currentName = String(currentRename.macro.name || '').trim();
    const nextName = encodeMacroName(String(currentRename.draft || '').trim());
    if (!shouldCommitModuleRename({ currentName, nextName })) {
      setModuleRenameState(null);
      return;
    }

    if (!isValidVbaModuleName(nextName)) {
      onActionStatus?.('error', 'Invalid name.');
      return;
    }

    const renameApi = window.excel?.vba?.renameMacroByWorkbook;
    if (typeof renameApi !== 'function') {
      onActionStatus?.('error', 'Rename unavailable.');
      return;
    }

    const request = buildMacroRenameRequest(currentRename.macro, searchData?.workbook);
    if (!request.workbookName && !request.workbookPath) {
      onActionStatus?.('error', 'Workbook not found.');
      return;
    }

    renameCommitInFlightRef.current = true;
    setModuleActionInFlight(true);
    onActionStatus?.('running', 'Renaming...');
    try {
      const result = await renameApi({
        ...request,
        nextMacroName: nextName
      });
      if (!result?.success) {
        const message = String(result?.message || 'Unable to rename macro.');
        onActionStatus?.('error', message);
        return;
      }

      setModuleRenameState(null);
      invalidateWorkbookMutation(request, {
        includeActiveWorkbook: true,
        includeExplorerAllFiles: true,
        includePersonalMacros: true,
        includeShortcutAudit: true,
        includeWorkbookScopedData: true
      });
      onActionStatus?.('success', `Renamed to "${nextName}".`);
    } catch (error) {
      const message = error?.message ? String(error.message) : 'Unable to rename macro.';
      onActionStatus?.('error', message);
    } finally {
      setModuleActionInFlight(false);
      renameCommitInFlightRef.current = false;
    }
  }, [invalidateWorkbookMutation, moduleRenameState, onActionStatus, searchData?.workbook]);

  const commitMetadataRename = useCallback(async () => {
    const draft = encodeMacroName((metadataRename || '').trim());
    setMetadataRename(null);
    if (!selectedNode || !draft) return;

    const currentName = String(selectedNode.data?.name || selectedNode.label || '').trim();
    if (!shouldCommitModuleRename({ currentName, nextName: draft })) return;
    if (!isValidVbaModuleName(draft)) {
      onActionStatus?.('error', 'Invalid name.');
      return;
    }

    if (selectedNode.nodeType === 'module') {
      const renameApi = window.excel?.vba?.renameModuleByWorkbook;
      if (typeof renameApi !== 'function') {
        onActionStatus?.('error', 'Rename unavailable.');
        return;
      }
      const request = buildWorkbookModuleRequest(selectedNode.data, searchData?.workbook);
      if (!request.workbookName && !request.workbookPath) {
        onActionStatus?.('error', 'Workbook not found.');
        return;
      }
      setModuleActionInFlight(true);
      onActionStatus?.('running', 'Renaming...');
      try {
        const result = await renameApi({ ...request, nextModuleName: draft });
        if (!result?.success) {
          onActionStatus?.('error', String(result?.message || 'Unable to rename module.'));
          return;
        }
        invalidateWorkbookMutation(request, {
          includeActiveWorkbook: true,
          includeExplorerAllFiles: true,
          includePersonalMacros: true,
          includeShortcutAudit: true,
          includeWorkbookScopedData: true
        });
        onActionStatus?.('success', `Renamed to "${draft}".`);
      } catch (error) {
        onActionStatus?.('error', error?.message ? String(error.message) : 'Unable to rename module.');
      } finally {
        setModuleActionInFlight(false);
      }
    } else if (selectedNode.nodeType === 'macro') {
      const renameApi = window.excel?.vba?.renameMacroByWorkbook;
      if (typeof renameApi !== 'function') {
        onActionStatus?.('error', 'Rename unavailable.');
        return;
      }
      const request = buildMacroRenameRequest(selectedNode.data, searchData?.workbook);
      if (!request.workbookName && !request.workbookPath) {
        onActionStatus?.('error', 'Workbook not found.');
        return;
      }
      setModuleActionInFlight(true);
      onActionStatus?.('running', 'Renaming...');
      try {
        const result = await renameApi({ ...request, nextMacroName: draft });
        if (!result?.success) {
          onActionStatus?.('error', String(result?.message || 'Unable to rename macro.'));
          return;
        }
        invalidateWorkbookMutation(request, {
          includeActiveWorkbook: true,
          includeExplorerAllFiles: true,
          includePersonalMacros: true,
          includeShortcutAudit: true,
          includeWorkbookScopedData: true
        });
        onActionStatus?.('success', `Renamed to "${draft}".`);
      } catch (error) {
        onActionStatus?.('error', error?.message ? String(error.message) : 'Unable to rename macro.');
      } finally {
        setModuleActionInFlight(false);
      }
    }
  }, [invalidateWorkbookMutation, metadataRename, onActionStatus, searchData?.workbook, selectedNode]);

  const handleRequestDeleteModule = useCallback((target) => {
    setModuleDeleteTarget(target?.module || null);
    setModuleContextMenu(null);
  }, []);

  const handleCancelDeleteModule = useCallback(() => {
    setModuleDeleteTarget(null);
  }, []);

  const handleConfirmDeleteModule = useCallback(async () => {
    if (!moduleDeleteTarget || moduleActionInFlight) {
      return;
    }

    const deleteApi = window.excel?.vba?.deleteModuleByWorkbook;
    if (typeof deleteApi !== 'function') {
      onActionStatus?.('error', 'Delete unavailable.');
      return;
    }

    const request = buildWorkbookModuleRequest(moduleDeleteTarget, searchData?.workbook);
    if (!request.workbookName && !request.workbookPath) {
      onActionStatus?.('error', 'Workbook not found.');
      return;
    }

    const moduleName = String(moduleDeleteTarget?.name || '').trim();
    setModuleActionInFlight(true);
    onActionStatus?.('running', 'Deleting...');
    try {
      const result = await deleteApi(request);
      if (!result?.success) {
        const message = String(result?.message || 'Unable to delete module.');
        onActionStatus?.('error', message);
        return;
      }

      setModuleDeleteTarget(null);
      setModuleRenameState((previous) => (
        previous?.module?.name === moduleName ? null : previous
      ));
      invalidateWorkbookMutation(request, {
        includeActiveWorkbook: true,
        includeExplorerAllFiles: true,
        includePersonalMacros: true,
        includeShortcutAudit: true,
        includeWorkbookScopedData: true
      });
      onActionStatus?.('success', 'Module deleted.');
    } catch (error) {
      const message = error?.message ? String(error.message) : 'Unable to delete module.';
      onActionStatus?.('error', message);
    } finally {
      setModuleActionInFlight(false);
    }
  }, [invalidateWorkbookMutation, moduleActionInFlight, moduleDeleteTarget, onActionStatus, searchData?.workbook]);

  useEffect(() => {
    if (!moduleContextMenu) {
      return;
    }

    const handlePointerDown = (event) => {
      if (moduleContextMenuRef.current && !moduleContextMenuRef.current.contains(event.target)) {
        setModuleContextMenu(null);
      }
    };

    const handleEscape = (event) => {
      if (event.key === 'Escape') {
        setModuleContextMenu(null);
      }
    };

    document.addEventListener('mousedown', handlePointerDown);
    window.addEventListener('keydown', handleEscape);
    window.addEventListener('scroll', closeModuleContextMenu, true);
    return () => {
      document.removeEventListener('mousedown', handlePointerDown);
      window.removeEventListener('keydown', handleEscape);
      window.removeEventListener('scroll', closeModuleContextMenu, true);
    };
  }, [closeModuleContextMenu, moduleContextMenu]);

  useEffect(() => {
    if (status === 'ready') {
      return;
    }
    setModuleContextMenu(null);
    setModuleRenameState(null);
    setModuleDeleteTarget(null);
  }, [status]);

  // --- Render helpers ---

  const renderNonReadyState = () => {
    const statusView = getSearchStatusView(searchData);

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
    const isExpanded = expandedIds.has(node.id);
    const isSelected = selectedNode?.id === node.id;
    const hasChildren = node.children.length > 0;
    const isExpandable = hasChildren || node.nodeType === 'workbook';
    const paddingLeft = 8 + depth * 16;
    const isRenaming = (node.nodeType === 'module' || node.nodeType === 'macro') && moduleRenameState?.nodeId === node.id;

    return (
      <React.Fragment key={node.id}>
        <div
          className={`tree-node tree-node--${node.nodeType} ${isSelected ? 'tree-node--selected' : ''}`}
          style={{ paddingLeft }}
          onClick={() => {
            if (isRenaming) return;
            clearTimeout(clickTimerRef.current);
            clickTimerRef.current = setTimeout(() => {
              setSelectedNode(node);
            }, 200);
          }}
          onDoubleClick={() => {
            clearTimeout(clickTimerRef.current);
            setSelectedNode(node);
            if (node.nodeType === 'module' && canShowModuleContextActions(node.data)) {
              handleStartRenameModule({ nodeId: node.id, module: node.data });
            } else if (node.nodeType === 'macro' && canShowMacroContextActions(node.data)) {
              handleStartRenameMacro({ nodeId: node.id, macro: node.data });
            }
          }}
          onContextMenu={(event) => {
            if (node.nodeType === 'module' || node.nodeType === 'macro') {
              handleOpenContextMenu(event, node);
            }
          }}
        >
          <span
            className={`tree-node__chevron ${isExpandable ? '' : 'tree-node__chevron--hidden'}`}
            onClick={(e) => {
              e.stopPropagation();
              if (isExpandable) toggleExpand(node.id);
            }}
          >
            {isExpandable && (
              <ChevronDownIcon
                size={14}
                className={`tree-chevron-icon ${isExpanded ? '' : 'tree-chevron-icon--collapsed'}`}
              />
            )}
          </span>

          <span className={`tree-node__icon tree-node__icon--${node.nodeType}`}>
            {node.nodeType === 'workbook' && <WorkbookIcon size={16} />}
            {node.nodeType === 'module' && <FolderIcon size={16} />}
            {node.nodeType === 'macro' && (() => {
              const macroIcon = node.data?.id ? getMacroIcon(node.data.id) : null;
              return isSpriteReady()
                ? <ImageMsoIcon name={macroIcon || 'MacroRecord'} size={14} />
                : <ReturnIcon size={14} />;
            })()}
          </span>

          <span className="tree-node__label">
            {isRenaming ? (
              <input
                type="text"
                className="module-inline-rename-input"
                value={moduleRenameState?.draft || ''}
                autoFocus
                spellCheck={false}
                maxLength={80}
                onChange={(event) => handleRenameDraftChange(event.target.value)}
                onClick={(event) => event.stopPropagation()}
                onKeyDown={(event) => {
                  if (event.key === 'Enter') {
                    event.preventDefault();
                    if (moduleRenameState?.type === 'macro') {
                      void handleCommitRenameMacro();
                    } else {
                      void handleCommitRenameModule();
                    }
                  } else if (event.key === 'Escape') {
                    event.preventDefault();
                    handleCancelRenameModule();
                  }
                }}
                onBlur={() => {
                  if (moduleRenameState?.type === 'macro') {
                    void handleCommitRenameMacro();
                  } else {
                    void handleCommitRenameModule();
                  }
                }}
              />
            ) : (
              (node.nodeType === 'module' || node.nodeType === 'macro') ? displayMacroName(node.label) : node.label
            )}
          </span>
        </div>

        {isExpanded && hasChildren && node.children.map((child) => renderTreeNode(child, depth + 1))}
        {isExpanded && !hasChildren && node.nodeType === 'workbook' && (
          <div className="tree-node-empty" style={{ paddingLeft: (depth + 1) * 16 + 28 }}>
            {String(node.data?.name || '').toUpperCase() === 'PERSONAL.XLSB' && personalState?.status !== 'ready' ? (
              <div className="ai-loading-indicator">
                <span className="ai-loading-dot" />
                <span className="ai-loading-dot" />
                <span className="ai-loading-dot" />
              </div>
            ) : (
              'This workbook has no macros'
            )}
          </div>
        )}
      </React.Fragment>
    );
  };

  return (
    <>
      {/* Main area */}
      {status !== 'ready' ? (
        <main className="main-content">
          {renderNonReadyState()}
        </main>
      ) : (
        <div className="split-view">
          {/* Left Panel - Tree */}
          {sidebarOpen && <div className="split-left" style={{ width: `${splitPct}%` }}>
            <div className="section-header">
              <span className="section-title">File Explorer</span>
            </div>

            <div className="tree-list">
              {tree.length === 0 && (
                <div className="search-empty-state">
                  {searchData?.vbaLocked ? 'This workbook\'s VBA project is locked.' : 'No files are available.'}
                </div>
              )}
              {tree.map((rootNode) => renderTreeNode(rootNode, 0))}
            </div>
          </div>}

          {sidebarOpen && <SplitDivider onResize={setSplitPct} />}

          {/* Right Panel - Details */}
          <div className="split-right">
            {selectedNode ? (
              <div className="details-panel">
                {selectedNode.nodeType === 'workbook' && (
                  <div className="folder-icon-large">
                    <WorkbookIcon size={80} />
                  </div>
                )}
                {selectedNode.nodeType === 'workbook' && isPersonalWorkbook(selectedNode) && (
                  <div className="details-workbook-actions">
                    <button
                      type="button"
                      className="details-edit-btn"
                      onClick={async () => {
                        const openFolderApi = window.excel?.personal?.openFolder;
                        if (typeof openFolderApi !== 'function') {
                          onActionStatus?.('error', 'Action unavailable.');
                          return;
                        }
                        try {
                          const result = await openFolderApi();
                          if (!result?.success) {
                            onActionStatus?.('error', String(result?.message || 'Unable to open folder.'));
                          }
                        } catch (error) {
                          onActionStatus?.('error', error?.message ? String(error.message) : 'Unable to open folder.');
                        }
                      }}
                    >
                      Open in File Explorer
                    </button>
                    <button
                      type="button"
                      className="details-edit-btn"
                      onClick={async () => {
                        const personalApi = window.excel?.personal;
                        if (!personalApi) return;
                        const isVisible = personalState.workbookFound
                          && personalState.windowVisible === true
                          && personalState.windowHidden !== true;
                        try {
                          if (isVisible) {
                            if (typeof personalApi.setVisibility !== 'function') return;
                            onActionStatus?.('running', 'Hiding...');
                            const result = await personalApi.setVisibility({ visible: false });
                            onActionStatus?.(result?.success ? 'success' : 'error', String(result?.message || 'Updated.'));
                          } else if (personalState.workbookFound) {
                            if (typeof personalApi.setVisibility !== 'function') return;
                            onActionStatus?.('running', 'Showing...');
                            const result = await personalApi.setVisibility({ visible: true });
                            onActionStatus?.(result?.success ? 'success' : 'error', String(result?.message || 'Updated.'));
                          } else {
                            if (typeof personalApi.open !== 'function') return;
                            onActionStatus?.('running', 'Opening...');
                            const result = await personalApi.open({ visible: true });
                            onActionStatus?.(result?.success ? 'success' : 'error', String(result?.message || 'Updated.'));
                          }
                          invalidateWorkbookMutation(
                            {
                              name: PERSONAL_WORKBOOK_NAME,
                              path: personalState.workbookPath || personalState.workbook?.path || ''
                            },
                            {
                              includeWorkbookList: personalState.workbookFound !== true,
                              includeExplorerAllFiles: personalState.workbookFound !== true,
                              includePersonalMacros: true,
                              includeWorkbookScopedData: true
                            }
                          );
                        } catch (error) {
                          onActionStatus?.('error', error?.message ? String(error.message) : 'Unable to update visibility.');
                        }
                      }}
                    >
                      {personalState.workbookFound && personalState.windowVisible === true && personalState.windowHidden !== true
                        ? 'Hide in Excel'
                        : 'Show in Excel'}
                    </button>
                  </div>
                )}
                {previewCode && (selectedNode.nodeType === 'module' || selectedNode.nodeType === 'macro') && (
                  <>
                    <CodePreview
                      code={previewCode}
                      showHeader={false}
                      editable={false}
                      status="normal"
                    />
                    <button
                      type="button"
                      className="details-edit-btn"
                      onClick={() => {
                        if (typeof onEditModule !== 'function') return;
                        const moduleName = selectedNode.nodeType === 'module'
                          ? (selectedNode.data?.name || selectedNode.label)
                          : (selectedNode.data?.module || '');
                        if (!moduleName) return;
                        const targetWorkbook = getNodeWorkbookTarget(selectedNode, searchData?.workbook);
                        onEditModule({
                          name: targetWorkbook.name,
                          path: targetWorkbook.path,
                          key: targetWorkbook.path || targetWorkbook.name
                        }, {
                          mode: 'existing_module',
                          moduleName,
                          source: 'all-files',
                          originMode: 'files'
                        });
                      }}
                    >
                      Edit
                    </button>
                  </>
                )}
                <div className="metadata-section">
                  <div className="metadata-title">Metadata</div>
                  {metadata.rows.map((row) => {
                    const isNameRow = row.label === 'Name'
                      && (selectedNode.nodeType === 'module' || selectedNode.nodeType === 'macro');
                    const isEditingName = isNameRow && metadataRename != null;

                    return (
                      <div className="metadata-row" key={row.label}>
                        <span className="metadata-label">{row.label}</span>
                        {isEditingName ? (
                          <input
                            type="text"
                            className="module-inline-rename-input metadata-rename-input"
                            value={metadataRename}
                            autoFocus
                            spellCheck={false}
                            maxLength={80}
                            onChange={(e) => setMetadataRename(e.target.value)}
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') {
                                e.preventDefault();
                                commitMetadataRename();
                              } else if (e.key === 'Escape') {
                                e.preventDefault();
                                setMetadataRename(null);
                              }
                            }}
                            onBlur={() => commitMetadataRename()}
                          />
                        ) : (
                          <span
                            className={`metadata-value${isNameRow ? ' metadata-value--editable' : ''}`}
                            onDoubleClick={() => {
                              if (isNameRow) {
                                setMetadataRename(displayMacroName(row.value));
                              }
                            }}
                          >
                            {isNameRow ? displayMacroName(row.value) : row.value}
                          </span>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
            ) : (
              <div className="details-panel details-panel--empty">
                <p className="details-empty-title">File Explorer</p>
                <p className="details-empty-blurb">Browse your workbooks, modules, and macros. Select an item from the tree to view its code and metadata.</p>
              </div>
            )}
          </div>
        </div>
      )}

      {moduleContextMenu && (
        <div
          ref={moduleContextMenuRef}
          className="module-context-menu"
          style={{ left: `${moduleContextMenu.x}px`, top: `${moduleContextMenu.y}px` }}
        >
          {moduleContextMenu.nodeType === 'module' && (
            <>
              <button
                type="button"
                className="module-context-menu-item"
                onClick={() => handleEditModule(moduleContextMenu)}
                disabled={moduleActionInFlight}
              >
                Edit
              </button>
              <button
                type="button"
                className="module-context-menu-item"
                onClick={() => handleStartRenameModule(moduleContextMenu)}
                disabled={moduleActionInFlight}
              >
                Rename
              </button>
              <button
                type="button"
                className="module-context-menu-item danger"
                onClick={() => handleRequestDeleteModule(moduleContextMenu)}
                disabled={moduleActionInFlight}
              >
                Delete
              </button>
            </>
          )}
          {moduleContextMenu.nodeType === 'macro' && (
            <button
              type="button"
              className="module-context-menu-item"
              onClick={() => handleStartRenameMacro(moduleContextMenu)}
              disabled={moduleActionInFlight}
            >
              Rename
            </button>
          )}
        </div>
      )}

      {moduleDeleteTarget && (
        <div className="module-action-overlay" role="dialog" aria-modal="true" aria-label="Delete module confirmation">
          <div className="module-action-dialog">
            <h3 className="module-action-title">Delete module?</h3>
            <p className="module-action-message">
              {`Delete "${displayMacroName(moduleDeleteTarget.name)}" from "${moduleDeleteTarget.workbookName || searchData?.workbook?.name || 'workbook'}"?`}
            </p>
            <div className="module-action-buttons">
              <button
                type="button"
                className="module-action-btn"
                onClick={handleCancelDeleteModule}
                disabled={moduleActionInFlight}
              >
                Cancel
              </button>
              <button
                type="button"
                className="module-action-btn danger"
                onClick={() => { void handleConfirmDeleteModule(); }}
                disabled={moduleActionInFlight}
              >
                Delete
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
};

export default FilesPage;
