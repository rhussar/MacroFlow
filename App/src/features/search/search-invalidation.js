import { useSyncExternalStore } from 'react';

export const SEARCH_INVALIDATION_BUCKETS = Object.freeze({
  ACTIVE_WORKBOOK: 'active-workbook',
  WORKBOOK_LIST: 'workbook-list',
  EXPLORER_ALL_FILES: 'explorer-all-files',
  PERSONAL_MACROS: 'personal-macros',
  SHORTCUT_AUDIT: 'shortcut-audit',
  WORKBOOK_SCOPED_DATA: 'workbook-scoped-data'
});

const bucketRevisionByKey = new Map();
const bucketListenersByKey = new Map();

export function normalizeSearchInvalidationScope(scope, fallback = 'global') {
  const normalizedScope = String(scope || '').trim();
  if (normalizedScope) {
    return normalizedScope;
  }

  const normalizedFallback = String(fallback || '').trim();
  return normalizedFallback || 'global';
}

export function getWorkbookInvalidationScope(workbook, fallback = 'workbook') {
  return normalizeSearchInvalidationScope(
    workbook?.key || workbook?.path || workbook?.name,
    fallback
  );
}

export function getShortcutAuditInvalidationScope({ scope = 'active', workbook = null } = {}) {
  return scope === 'workbook'
    ? getWorkbookInvalidationScope(workbook, 'workbook')
    : 'active-workbook';
}

export function getWorkbookScopedDataInvalidationScope(workbook, namespaceMacros = false) {
  const workbookScope = getWorkbookInvalidationScope(workbook, 'workbook');
  return `${workbookScope}::${namespaceMacros ? 'namespaced' : 'plain'}`;
}

export function buildWorkbookInvalidationDescriptors({
  workbook = null,
  includeActiveWorkbook = false,
  includeWorkbookList = false,
  includeExplorerAllFiles = false,
  includePersonalMacros = false,
  includeShortcutAudit = false,
  includeWorkbookScopedData = false
} = {}) {
  const descriptors = [];

  if (includeActiveWorkbook) {
    descriptors.push({ bucket: SEARCH_INVALIDATION_BUCKETS.ACTIVE_WORKBOOK });
  }
  if (includeWorkbookList) {
    descriptors.push({ bucket: SEARCH_INVALIDATION_BUCKETS.WORKBOOK_LIST });
  }
  if (includeExplorerAllFiles) {
    descriptors.push({ bucket: SEARCH_INVALIDATION_BUCKETS.EXPLORER_ALL_FILES });
  }
  if (includePersonalMacros) {
    descriptors.push({ bucket: SEARCH_INVALIDATION_BUCKETS.PERSONAL_MACROS });
  }
  if (includeShortcutAudit) {
    descriptors.push({
      bucket: SEARCH_INVALIDATION_BUCKETS.SHORTCUT_AUDIT,
      scope: getShortcutAuditInvalidationScope({ scope: 'workbook', workbook })
    });
  }
  if (includeWorkbookScopedData) {
    descriptors.push({
      bucket: SEARCH_INVALIDATION_BUCKETS.WORKBOOK_SCOPED_DATA,
      scope: getWorkbookScopedDataInvalidationScope(workbook, false)
    });
    descriptors.push({
      bucket: SEARCH_INVALIDATION_BUCKETS.WORKBOOK_SCOPED_DATA,
      scope: getWorkbookScopedDataInvalidationScope(workbook, true)
    });
  }

  return descriptors;
}

export function buildSearchInvalidationKey(bucket, scope = 'global') {
  const normalizedBucket = String(bucket || '').trim();
  const normalizedScope = normalizeSearchInvalidationScope(scope);
  return `${normalizedBucket}::${normalizedScope}`;
}

function getSearchInvalidationRevision(bucket, scope = 'global') {
  return Number(bucketRevisionByKey.get(buildSearchInvalidationKey(bucket, scope)) || 0);
}

function emitSearchInvalidation(bucket, scope = 'global') {
  const key = buildSearchInvalidationKey(bucket, scope);
  const nextRevision = getSearchInvalidationRevision(bucket, scope) + 1;
  bucketRevisionByKey.set(key, nextRevision);

  const listeners = bucketListenersByKey.get(key);
  if (listeners) {
    listeners.forEach((listener) => listener());
  }

  return nextRevision;
}

export function invalidateSearchBucket(bucket, scope = 'global') {
  return emitSearchInvalidation(bucket, scope);
}

export function invalidateSearchBuckets(descriptors = []) {
  const seenKeys = new Set();
  (Array.isArray(descriptors) ? descriptors : []).forEach((descriptor) => {
    const bucket = String(descriptor?.bucket || '').trim();
    if (!bucket) {
      return;
    }

    const scope = normalizeSearchInvalidationScope(descriptor?.scope);
    const key = buildSearchInvalidationKey(bucket, scope);
    if (seenKeys.has(key)) {
      return;
    }

    seenKeys.add(key);
    emitSearchInvalidation(bucket, scope);
  });
}

function subscribeToSearchInvalidation(bucket, scope = 'global', listener) {
  const key = buildSearchInvalidationKey(bucket, scope);
  const listeners = bucketListenersByKey.get(key) || new Set();
  listeners.add(listener);
  bucketListenersByKey.set(key, listeners);

  return () => {
    const currentListeners = bucketListenersByKey.get(key);
    if (!currentListeners) {
      return;
    }

    currentListeners.delete(listener);
    if (currentListeners.size === 0) {
      bucketListenersByKey.delete(key);
    }
  };
}

export function useSearchInvalidationRevision(bucket, scope = 'global') {
  const normalizedScope = normalizeSearchInvalidationScope(scope);
  return useSyncExternalStore(
    (listener) => subscribeToSearchInvalidation(bucket, normalizedScope, listener),
    () => getSearchInvalidationRevision(bucket, normalizedScope),
    () => 0
  );
}
