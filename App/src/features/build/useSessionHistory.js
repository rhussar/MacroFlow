import { useState, useCallback, useRef } from 'react';

function generateId() {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return crypto.randomUUID();
  }
  return Date.now() + '-' + Math.random().toString(36).slice(2);
}

export function useSessionHistory() {
  const sessionsRef = useRef([]);
  const [generation, setGeneration] = useState(0);

  const bump = () => setGeneration((g) => g + 1);

  const saveSession = useCallback((payload) => {
    if (!payload?.sessionContext || !Array.isArray(payload.messages) || payload.messages.length === 0) {
      return null;
    }

    const entry = {
      id: generateId(),
      sessionContext: payload.sessionContext,
      messages: payload.messages,
      editedCode: payload.editedCode || '',
      savedAt: Date.now()
    };

    sessionsRef.current = [...sessionsRef.current, entry];
    bump();
    return entry.id;
  }, []);

  const restoreSession = useCallback((id) => {
    return sessionsRef.current.find((s) => s.id === id) || null;
  }, []);

  const removeSession = useCallback((id) => {
    sessionsRef.current = sessionsRef.current.filter((s) => s.id !== id);
    bump();
  }, []);

  const clearAll = useCallback(() => {
    sessionsRef.current = [];
    bump();
  }, []);

  // Read snapshot — generation dependency ensures consumers re-render
  void generation;
  const sessions = sessionsRef.current;

  return { sessions, saveSession, restoreSession, removeSession, clearAll };
}
