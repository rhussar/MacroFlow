const { spawn } = require('node:child_process');
const fs = require('node:fs');
const path = require('node:path');
const readline = require('node:readline');

const DEFAULT_MAX_RESTART_ATTEMPTS = 3;
const DEFAULT_RESTART_BASE_DELAY_MS = 300;
const HEARTBEAT_INTERVAL_MS = 30_000;
const HEARTBEAT_TIMEOUT_MS = 5_000;
const STOP_KILL_TIMEOUT_MS = 700;

function resolveHelperPath() {
  const candidates = [];

  if (process.resourcesPath) {
    candidates.push(path.join(process.resourcesPath, 'bin-helper', 'WindowFocusHelper.exe'));
  }

  // Dev and local build output locations.
  candidates.push(path.join(__dirname, '../native/window-focus-helper/bin-helper/WindowFocusHelper.exe'));
  candidates.push(
    path.join(__dirname, '../native/window-focus-helper/bin/Release/net8.0-windows/win-x64/publish/WindowFocusHelper.exe')
  );

  for (const candidate of candidates) {
    if (fs.existsSync(candidate)) {
      return candidate;
    }
  }

  return null;
}

class WindowFocusHelperClient {
  constructor(options = {}) {
    this.onStateChange = options.onStateChange || (() => {});
    this.onError = options.onError || (() => {});
    this.onFatal = options.onFatal || (() => {});
    this.logger = options.logger || null;

    this.maxRestartAttempts = Number.isInteger(options.maxRestartAttempts)
      ? Math.max(options.maxRestartAttempts, 0)
      : DEFAULT_MAX_RESTART_ATTEMPTS;
    this.restartBaseDelayMs = typeof options.restartBaseDelayMs === 'number'
      ? Math.max(100, options.restartBaseDelayMs)
      : DEFAULT_RESTART_BASE_DELAY_MS;

    this.targetHwnd = null;
    this.process = null;
    this.stdoutInterface = null;
    this.stopping = false;
    this.restartAttempts = 0;
    this.restartTimer = null;

    this.heartbeatTimer = null;
    this.pongTimeout = null;
    this.awaitingPong = false;
    this.lastPongAt = 0;
    this.unexpectedExitHandled = false;
  }

  start(targetHwnd) {
    if (targetHwnd === null || targetHwnd === undefined) {
      this.reportFatal('Cannot start window helper: target window handle is missing.');
      return;
    }

    this.targetHwnd = String(targetHwnd);
    this.stopping = false;

    if (this.process) {
      this.updateTarget(this.targetHwnd);
      return;
    }

    this.restartAttempts = 0;
    this.unexpectedExitHandled = false;
    this.launchHelper('initial-start');
  }

  updateTarget(targetHwnd) {
    if (targetHwnd === null || targetHwnd === undefined) {
      return;
    }

    this.targetHwnd = String(targetHwnd);
    this.send({ type: 'setTarget', hwnd: this.targetHwnd });
  }

  stop() {
    this.stopping = true;
    this.unexpectedExitHandled = true;
    this.clearRestartTimer();
    this.stopHeartbeat();

    if (!this.process) {
      return;
    }

    this.send({ type: 'shutdown' });

    const child = this.process;
    setTimeout(() => {
      if (child && !child.killed) {
        try {
          child.kill();
        } catch {
          // Ignore kill errors when process has already exited.
        }
      }
    }, STOP_KILL_TIMEOUT_MS);
  }

  launchHelper(reason) {
    const helperPath = resolveHelperPath();
    if (!helperPath) {
      this.reportFatal(
        'WindowFocusHelper.exe not found. Run "npm run build:window-helper" from App before starting Electron.'
      );
      return;
    }

    let child = null;
    try {
      child = spawn(helperPath, [], {
        windowsHide: true,
        stdio: ['pipe', 'pipe', 'pipe']
      });
    } catch (error) {
      this.handleUnexpectedTermination(`Window helper failed to launch: ${error.message}`);
      return;
    }

    this.process = child;
    this.unexpectedExitHandled = false;
    this.log('info', 'Window helper started', {
      pid: child.pid,
      reason,
      helperPath
    });

    child.stderr.setEncoding('utf8');
    child.stderr.on('data', (chunk) => {
      const message = String(chunk || '').trim();
      if (message) {
        this.onError(`[WindowHelper STDERR] ${message}`, { type: 'stderr', message });
      }
    });

    child.on('error', (error) => {
      if (this.stopping) {
        return;
      }
      this.handleUnexpectedTermination(`Window helper process error: ${error.message}`);
    });

    child.on('exit', (code, signal) => {
      const wasStopping = this.stopping;
      this.cleanupHandles();

      if (wasStopping) {
        this.log('info', 'Window helper stopped', { code, signal: signal || 'none' });
        return;
      }

      this.handleUnexpectedTermination(
        `Window helper exited unexpectedly (code=${code}, signal=${signal || 'none'}).`
      );
    });

    this.stdoutInterface = readline.createInterface({ input: child.stdout });
    this.stdoutInterface.on('line', (line) => this.handleMessage(line));

    this.startHeartbeat();
    this.send({ type: 'setTarget', hwnd: this.targetHwnd });
    this.sendPing();
  }

  handleUnexpectedTermination(message) {
    if (this.unexpectedExitHandled) {
      return;
    }
    this.unexpectedExitHandled = true;

    this.onError(message, { type: 'process-exit', message });
    this.log('warn', 'Window helper unexpected termination', {
      message,
      attempt: this.restartAttempts,
      maxAttempts: this.maxRestartAttempts
    });

    if (this.stopping) {
      return;
    }

    this.stopHeartbeat();
    this.cleanupHandles();

    if (this.restartAttempts >= this.maxRestartAttempts) {
      this.reportFatal(
        `Window helper failed after ${this.maxRestartAttempts} restart attempt(s). Falling back to always-on-top mode.`
      );
      return;
    }

    this.restartAttempts += 1;
    const delayMs = Math.min(3_000, this.restartBaseDelayMs * (2 ** (this.restartAttempts - 1)));
    this.log('warn', 'Scheduling helper restart', {
      attempt: this.restartAttempts,
      maxAttempts: this.maxRestartAttempts,
      delayMs
    });

    this.clearRestartTimer();
    this.restartTimer = setTimeout(() => {
      this.restartTimer = null;
      if (this.stopping) {
        return;
      }
      this.launchHelper('auto-restart');
    }, delayMs);
  }

  reportFatal(message) {
    this.unexpectedExitHandled = true;
    this.log('error', 'Window helper fatal', { message });
    this.onError(message, { type: 'fatal', message });
    this.onFatal(message);
  }

  startHeartbeat() {
    this.stopHeartbeat();

    this.heartbeatTimer = setInterval(() => {
      if (!this.process) {
        return;
      }
      this.sendPing();
    }, HEARTBEAT_INTERVAL_MS);
  }

  stopHeartbeat() {
    if (this.heartbeatTimer) {
      clearInterval(this.heartbeatTimer);
      this.heartbeatTimer = null;
    }

    if (this.pongTimeout) {
      clearTimeout(this.pongTimeout);
      this.pongTimeout = null;
    }

    this.awaitingPong = false;
  }

  sendPing() {
    if (!this.process) {
      return;
    }

    if (this.awaitingPong) {
      this.handleUnexpectedTermination('Window helper heartbeat timed out before pong was received.');
      if (this.process && !this.process.killed) {
        try {
          this.process.kill();
        } catch {
          // Ignore kill race.
        }
      }
      return;
    }

    this.awaitingPong = true;
    this.send({ type: 'ping' });

    if (this.pongTimeout) {
      clearTimeout(this.pongTimeout);
    }

    this.pongTimeout = setTimeout(() => {
      this.pongTimeout = null;
      if (!this.awaitingPong) {
        return;
      }
      this.handleUnexpectedTermination('Window helper heartbeat timeout waiting for pong.');
      if (this.process && !this.process.killed) {
        try {
          this.process.kill();
        } catch {
          // Ignore kill race.
        }
      }
    }, HEARTBEAT_TIMEOUT_MS);
  }

  send(message) {
    if (!this.process || !this.process.stdin || this.process.stdin.destroyed) {
      return;
    }

    try {
      this.process.stdin.write(`${JSON.stringify(message)}\n`);
    } catch (error) {
      this.onError(`Failed to send helper message: ${error.message}`, { type: 'send-failed', error: error.message });
    }
  }

  handleMessage(line) {
    const trimmed = String(line || '').trim();
    if (!trimmed) {
      return;
    }

    let payload = null;
    try {
      payload = JSON.parse(trimmed);
    } catch {
      this.onError(`Malformed helper output: ${trimmed}`, { type: 'malformed-output', raw: trimmed });
      return;
    }

    if (payload.type === 'state') {
      this.onStateChange(payload);
      return;
    }

    if (payload.type === 'pong') {
      this.awaitingPong = false;
      this.lastPongAt = Date.now();
      if (this.pongTimeout) {
        clearTimeout(this.pongTimeout);
        this.pongTimeout = null;
      }
      return;
    }

    if (payload.type === 'error') {
      this.onError(payload.message || 'Unknown helper error', payload);
    }
  }

  clearRestartTimer() {
    if (this.restartTimer) {
      clearTimeout(this.restartTimer);
      this.restartTimer = null;
    }
  }

  cleanupHandles() {
    if (this.stdoutInterface) {
      this.stdoutInterface.removeAllListeners();
      this.stdoutInterface.close();
      this.stdoutInterface = null;
    }

    if (this.process) {
      this.process.removeAllListeners();
      this.process = null;
    }

    if (this.pongTimeout) {
      clearTimeout(this.pongTimeout);
      this.pongTimeout = null;
    }
    this.awaitingPong = false;
  }

  log(level, message, meta) {
    if (!this.logger || typeof this.logger[level] !== 'function') {
      return;
    }
    this.logger[level](message, meta);
  }
}

module.exports = WindowFocusHelperClient;





