function defaultIsValidValue(value) {
  return value !== null && value !== undefined;
}

export function normalizeAsyncResourceKey(key, fallback = '__default__') {
  const normalizedKey = String(key || '').trim();
  if (normalizedKey) {
    return normalizedKey;
  }

  const normalizedFallback = String(fallback || '').trim();
  return normalizedFallback || '__default__';
}

export function createAsyncResourceStore(options = {}) {
  const defaultKey = normalizeAsyncResourceKey(options?.defaultKey, '__default__');
  const cache = new Map();
  const pendingRequests = new Map();

  const resolveKey = (key) => normalizeAsyncResourceKey(key, defaultKey);

  const getEntry = (key) => cache.get(resolveKey(key)) || null;

  const getFreshEntry = ({ key, ttlMs, now = Date.now(), isValid = defaultIsValidValue } = {}) => {
    const entry = getEntry(key);
    if (!entry) {
      return null;
    }

    const ageMs = Number(now) - Number(entry?.fetchedAt || 0);
    if (!Number.isFinite(ageMs) || ageMs < 0 || ageMs >= Number(ttlMs)) {
      return null;
    }

    if (typeof isValid === 'function' && !isValid(entry.value)) {
      return null;
    }

    return entry;
  };

  const setValue = (key, value, options = {}) => {
    const fetchedAt = Number.isFinite(Number(options?.fetchedAt))
      ? Number(options.fetchedAt)
      : Date.now();
    cache.set(resolveKey(key), {
      value,
      fetchedAt
    });
    return value;
  };

  const clear = (key) => {
    const resolvedKey = resolveKey(key);
    cache.delete(resolvedKey);
    pendingRequests.delete(resolvedKey);
  };

  const clearAll = () => {
    cache.clear();
    pendingRequests.clear();
  };

  const run = (key, requestFactory) => {
    const resolvedKey = resolveKey(key);
    const pendingRequest = pendingRequests.get(resolvedKey);
    if (pendingRequest) {
      return pendingRequest;
    }

    const requestPromise = Promise.resolve()
      .then(requestFactory)
      .finally(() => {
        if (pendingRequests.get(resolvedKey) === requestPromise) {
          pendingRequests.delete(resolvedKey);
        }
      });

    pendingRequests.set(resolvedKey, requestPromise);
    return requestPromise;
  };

  return {
    clear,
    clearAll,
    getEntry,
    getFreshEntry,
    run,
    setValue
  };
}
