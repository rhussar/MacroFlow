import { useState, useRef, useCallback } from 'react';

export function useMacroRun({ loadSearchDataRef, setActionStatus }) {
  const [selectedMacro, setSelectedMacro] = useState(null);
  const [runState, setRunState] = useState('idle');
  const macroRunInFlightRef = useRef(false);

  const handleRunMacro = useCallback(async (macro) => {
    if (!macro || macroRunInFlightRef.current || runState === 'running') {
      return;
    }

    setSelectedMacro(macro);
    const macroName = macro?.fullName || macro?.runTarget || macro?.name || '';
    if (!macroName) {
      setRunState('error');
      setActionStatus('error', 'Run failed: Macro identity is missing.');
      return;
    }

    const runApi = window.excel?.vba?.run;
    if (!runApi) {
      setRunState('error');
      setActionStatus('error', 'Run failed: Excel VBA run API is unavailable.');
      return;
    }

    macroRunInFlightRef.current = true;
    setRunState('running');
    setActionStatus('running', `Running ${macro.name || macroName}...`);

    try {
      const result = await runApi({ macroName });
      if (result?.success) {
        const backendMessage = result?.message || `Executed "${macroName}"`;
        setRunState('success');
        setActionStatus('success', `Run succeeded: ${backendMessage}`);
      } else {
        const backendMessage = result?.message || 'Unknown error.';
        setRunState('error');
        setActionStatus('error', `Run failed: ${backendMessage}`);
      }
    } catch (error) {
      const backendMessage = error?.message ? String(error.message) : 'Unexpected error.';
      setRunState('error');
      setActionStatus('error', `Run failed: ${backendMessage}`);
    } finally {
      macroRunInFlightRef.current = false;
      if (typeof loadSearchDataRef?.current === 'function') {
        await loadSearchDataRef.current({ silent: true });
      }
      setSelectedMacro(null);
    }
  }, [loadSearchDataRef, runState, setActionStatus]);

  return {
    selectedMacro,
    setSelectedMacro,
    runState,
    handleRunMacro,
    macroRunInFlightRef
  };
}
