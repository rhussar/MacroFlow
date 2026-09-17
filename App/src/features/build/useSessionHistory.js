import { useState, useCallback, useRef } from 'react';
import { loadSessions, saveSessions } from './sessionStore';

function generateId() {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return crypto.randomUUID();
  }
  return Date.now() + '-' + Math.random().toString(36).slice(2);
}

export function useSessionHistory() {
  const sessionsRef = useRef(loadSessions());
  const [generation, setGeneration] = useState(0);

  const bump = () => setGeneration((g) => g + 1);

  const saveSession = useCallback((payload) => {
    if (!payload?.sessionContext || !Array.isArray(payload.messages) || payload.messages.length === 0) {
      return null;
    }

    // Prevent duplicate saves for the same module with identical message count
    const ctx = payload.sessionContext;
    const duplicate = sessionsRef.current.find((s) =>
      s.sessionContext?.moduleName === ctx?.moduleName &&
      s.sessionContext?.workbook?.key === ctx?.workbook?.key &&
      s.messages.length === payload.messages.length
    );
    if (duplicate) {
      return duplicate.id;
    }

    const entry = {
      id: generateId(),
      sessionContext: payload.sessionContext,
      messages: payload.messages,
      editedCode: payload.editedCode || '',
      savedAt: Date.now()
    };

    sessionsRef.current = [...sessionsRef.current, entry];
    saveSessions(sessionsRef.current);
    bump();
    return entry.id;
  }, []);

  const updateSession = useCallback((id, payload) => {
    const idx = sessionsRef.current.findIndex((s) => s.id === id);
    if (idx === -1) return false;
    sessionsRef.current = sessionsRef.current.map((s) =>
      s.id === id
        ? { ...s, messages: payload.messages, editedCode: payload.editedCode || '', savedAt: Date.now() }
        : s
    );
    saveSessions(sessionsRef.current);
    bump();
    return true;
  }, []);

  const restoreSession = useCallback((id) => {
    return sessionsRef.current.find((s) => s.id === id) || null;
  }, []);

  const removeSession = useCallback((id) => {
    sessionsRef.current = sessionsRef.current.filter((s) => s.id !== id);
    saveSessions(sessionsRef.current);
    bump();
  }, []);

  const clearAll = useCallback(() => {
    sessionsRef.current = [];
    saveSessions([]);
    bump();
  }, []);

  // Read snapshot — generation dependency ensures consumers re-render
  void generation;
  const sessions = sessionsRef.current;

  return { sessions, saveSession, updateSession, restoreSession, removeSession, clearAll };
}
