import { useState, useEffect, useCallback, useRef } from 'react';
import { MacroFlowLogo } from './icons';

/* global __APP_VERSION__ */

const STAGE_LABELS = {
  checking: 'Checking system',
  downloading_runtime: 'Downloading runtime',
  verifying_runtime: 'Verifying download',
  downloading_runtime_gpu: 'Downloading GPU support',
  verifying_runtime_gpu: 'Verifying GPU support',
  extracting_runtime: 'Installing runtime',
  starting_runtime: 'Starting AI engine',
  downloading_model: 'Downloading model',
  ready: 'Finishing up',
};

// Map Ollama's raw status strings to friendly labels
function friendlyModelStatus(statusText) {
  if (!statusText) return null;
  const s = statusText.toLowerCase();
  if (s.includes('pulling manifest')) return 'Preparing download';
  if (s.includes('downloading')) return 'Downloading model';
  if (s.includes('verifying')) return 'Verifying model';
  if (s.includes('writing manifest')) return 'Finalizing';
  if (s.includes('success')) return 'Finishing up';
  return null;
}

/**
 * AiGate — overlay shown on the Create page when local AI model is not downloaded.
 * Blurs the background content and shows a clean download prompt.
 * Only checks once per session.
 */
export default function AiGate({ onReady, onNotReady, initialReady, children }) {
  const [aiStatus, setAiStatus] = useState(null);
  const [setupInProgress, setSetupInProgress] = useState(false);
  const [error, setError] = useState('');
  const [ready, setReady] = useState(initialReady ? true : null); // skip check if parent already knows

  const onReadyRef = useRef(onReady);
  onReadyRef.current = onReady;
  const onNotReadyRef = useRef(onNotReady);
  onNotReadyRef.current = onNotReady;

  const handleStatusUpdate = useCallback((status) => {
    setAiStatus(status);
    if (status?.ready) {
      setReady(true);
      setSetupInProgress(false);
      onReadyRef.current?.();
    } else if (status?.lastError) {
      setError(status.lastError);
      setSetupInProgress(false);
    }
  }, []);

  useEffect(() => {
    let cancelled = false;
    // Check once on mount
    window.excel?.ai?.getStatus?.().then((status) => {
      if (cancelled) return;
      if (status?.ready) {
        setReady(true);
        onReadyRef.current?.();
      } else {
        setReady(false);
        setAiStatus(status);
        onNotReadyRef.current?.();
      }
    }).catch(() => {
      if (!cancelled) { setReady(false); onNotReadyRef.current?.(); }
    });

    const unsubscribe = window.excel?.ai?.onStatus?.(handleStatusUpdate) || (() => {});
    return () => { cancelled = true; unsubscribe(); };
  }, [handleStatusUpdate]);

  const handleDownload = async () => {
    setSetupInProgress(true);
    setError('');
    try {
      const result = await window.excel?.ai?.setup?.();
      if (result) handleStatusUpdate(result);
    } catch {
      setError('Setup failed. Please try again.');
      setSetupInProgress(false);
    }
  };

  // Progress info — track if we've ever hit 100% to avoid flickering back
  const hitFullRef = useRef(false);
  const rawProgress = aiStatus?.progress;
  const rawPct = rawProgress != null ? Math.round(rawProgress * 100) : null;
  const stage = aiStatus?.stage;
  const statusText = aiStatus?.statusText || '';

  // Once progress hits 100% during model download, lock to indeterminate
  if (stage === 'downloading_model' && rawPct === 100) hitFullRef.current = true;
  // Reset if stage changes away from downloading_model
  if (stage && stage !== 'downloading_model') hitFullRef.current = false;

  const pastFull = hitFullRef.current;
  const progress = pastFull ? null : rawPct;

  // Use friendly Ollama status when downloading model, otherwise use stage label
  const stageLabel = (stage === 'downloading_model' && friendlyModelStatus(statusText))
    || STAGE_LABELS[stage]
    || '';
  // Indeterminate when no progress value
  const isIndeterminate = setupInProgress && progress === null;

  // Model size from status or default
  const modelSize = aiStatus?.modelSize || '5.1 GB';

  // Ready — render children directly, no wrapper
  if (ready === true) return children;

  return (
    <div style={{ position: 'relative', flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
      <div className="ai-gate-bg-blur" />
      <div className="ai-gate-overlay">
        <div className="ai-gate-card">
          <MacroFlowLogo size={64} />
          <p className="ai-gate-description">
            To start building macros with AI, you need to download the local model first. This can easily be undone.
          </p>
          {!setupInProgress && (
            <button className="ai-gate-download-btn" onClick={handleDownload}>
              {`Download Now (${modelSize})`}
            </button>
          )}
          {setupInProgress && (
            <div className="ai-gate-progress-section">
              <div className="ai-gate-progress-label">
                <span>{stageLabel || 'Setting up'}</span>
                {progress !== null && !isIndeterminate && <span>{progress}%</span>}
              </div>
              <div className="ai-gate-progress">
                <div
                  className={`ai-gate-progress-bar${isIndeterminate ? ' ai-gate-progress-indeterminate' : ''}`}
                  style={progress !== null ? { width: `${progress}%` } : undefined}
                />
              </div>
            </div>
          )}
          {error && <div className="ai-gate-error">{error}</div>}
          <div className="ai-gate-badges">
            <span className="ai-gate-badge">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <rect x="2" y="2" width="20" height="20" rx="2" />
                <path d="M8 12h8M12 8v8" />
              </svg>
              100% Offline
            </span>
            <span className="ai-gate-badge">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <rect x="3" y="11" width="18" height="11" rx="2" ry="2" />
                <path d="M7 11V7a5 5 0 0 1 10 0v4" />
              </svg>
              Fully Secure
            </span>
          </div>
        </div>
      </div>
    </div>
  );
}
