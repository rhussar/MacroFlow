import { useState, useRef, useCallback, useEffect } from 'react';
import { BUILD_MODE_SEED_CODE } from './build-target';

export function useCodeEditing() {
  const [editedCode, setEditedCode] = useState(BUILD_MODE_SEED_CODE);
  const [hasPendingChanges, setHasPendingChanges] = useState(false);
  const editedCodeRef = useRef(editedCode);
  const dirtyLocalRef = useRef(false);

  useEffect(() => {
    editedCodeRef.current = editedCode;
  }, [editedCode]);

  const setDirtyState = useCallback((value) => {
    const nextValue = Boolean(value);
    dirtyLocalRef.current = nextValue;
    setHasPendingChanges(nextValue);
  }, []);

  const markLocalDirty = useCallback(() => {
    setDirtyState(true);
  }, [setDirtyState]);

  return {
    editedCode, setEditedCode,
    editedCodeRef,
    hasPendingChanges,
    dirtyLocalRef,
    setDirtyState,
    markLocalDirty
  };
}
