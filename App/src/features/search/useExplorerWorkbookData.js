import { useEffect, useMemo, useRef, useState } from 'react';
import { normalizeMacros, normalizeModules } from '../../lib/search-data.js';

const INITIAL_EXPLORER_DATA = {
  status: 'idle',
  workbook: null,
  modules: [],
  macros: [],
  error: null
};

function toSafeString(value) {
  return String(value || '').trim();
}

function toExplorerWorkbookRequest(workbook) {
  return {
    workbookName: toSafeString(workbook?.name),
    workbookPath: toSafeString(workbook?.path)
  };
}

function toExplorerWorkbookModel(workbook, fallback = null) {
  const name = toSafeString(workbook?.name || fallback?.name);
  const path = toSafeString(workbook?.path || fallback?.path);
  if (!name && !path) {
    return null;
  }
  return {
    name: name || 'Workbook',
    path,
    key: path || name
  };
}

export function resolveInitialModuleNode(tree, { moduleId = '', moduleName = '' } = {}) {
  const source = Array.isArray(tree) ? tree : [];
  const normalizedId = toSafeString(moduleId);
  const normalizedName = toSafeString(moduleName).toLowerCase();

  const walk = (nodes) => {
    for (const node of nodes) {
      if (node?.nodeType === 'module') {
        if (normalizedId && String(node.id || '').trim() === normalizedId) {
          return node;
        }
        const nodeName = toSafeString(node?.data?.name || node?.label).toLowerCase();
        if (!normalizedId && normalizedName && nodeName === normalizedName) {
          return node;
        }
      }
      if (Array.isArray(node?.children) && node.children.length > 0) {
        const childMatch = walk(node.children);
        if (childMatch) {
          return childMatch;
        }
      }
    }
    return null;
  };

  return walk(source);
}

export function useExplorerWorkbookData(searchData, explorerContext) {
  const [explorerData, setExplorerData] = useState(INITIAL_EXPLORER_DATA);
  const requestSequence = useRef(0);

  const contextWorkbook = useMemo(
    () => toExplorerWorkbookModel(explorerContext?.workbook),
    [explorerContext?.workbook?.name, explorerContext?.workbook?.path, explorerContext?.workbook?.key]
  );
  const hasContext = Boolean(contextWorkbook?.key);

  useEffect(() => {
    if (!hasContext) {
      requestSequence.current += 1;
      setExplorerData(INITIAL_EXPLORER_DATA);
      return;
    }

    const modulesByWorkbookApi = window.excel?.vba?.modulesByWorkbook;
    const proceduresByWorkbookApi = window.excel?.vba?.proceduresByWorkbook;
    if (!modulesByWorkbookApi || !proceduresByWorkbookApi) {
      setExplorerData({
        status: 'error',
        workbook: contextWorkbook,
        modules: [],
        macros: [],
        error: { message: 'Workbook-scoped VBA APIs are unavailable.' }
      });
      return;
    }

    let cancelled = false;
    const requestId = ++requestSequence.current;
    setExplorerData({
      status: 'loading',
      workbook: contextWorkbook,
      modules: [],
      macros: [],
      error: null
    });

    (async () => {
      try {
        const request = toExplorerWorkbookRequest(contextWorkbook);
        const [modulesResult, proceduresResult] = await Promise.all([
          modulesByWorkbookApi(request),
          proceduresByWorkbookApi(request)
        ]);

        if (cancelled || requestId !== requestSequence.current) {
          return;
        }

        if (!modulesResult?.success || !proceduresResult?.success) {
          const message = [modulesResult?.message, proceduresResult?.message]
            .filter(Boolean)
            .join(' | ') || 'Unable to load workbook data.';
          setExplorerData({
            status: 'error',
            workbook: contextWorkbook,
            modules: [],
            macros: [],
            error: { message }
          });
          return;
        }

        const workbookFound = modulesResult?.workbookFound !== false && proceduresResult?.workbookFound !== false;
        if (!workbookFound) {
          setExplorerData({
            status: 'error',
            workbook: contextWorkbook,
            modules: [],
            macros: [],
            error: { message: `The workbook "${contextWorkbook.name}" is no longer open.` }
          });
          return;
        }

        const workbook = toExplorerWorkbookModel(
          modulesResult?.workbook || proceduresResult?.workbook,
          contextWorkbook
        ) || contextWorkbook;
        const normalizedWorkbook = { name: workbook.name, path: workbook.path };
        const modules = normalizeModules(modulesResult?.modules, normalizedWorkbook);
        const macros = normalizeMacros(proceduresResult?.procedures);
        setExplorerData({
          status: 'ready',
          workbook,
          modules,
          macros,
          error: null
        });
      } catch (error) {
        if (cancelled || requestId !== requestSequence.current) {
          return;
        }
        const message = error?.message ? String(error.message) : 'Unable to load workbook data.';
        setExplorerData({
          status: 'error',
          workbook: contextWorkbook,
          modules: [],
          macros: [],
          error: { message }
        });
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [contextWorkbook, hasContext]);

  return hasContext ? explorerData : searchData;
}

