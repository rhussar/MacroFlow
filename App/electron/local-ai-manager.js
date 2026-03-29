const fs = require('node:fs');
const path = require('node:path');
const http = require('node:http');
const https = require('node:https');
const { spawn } = require('node:child_process');

const logger = require('./logger');
const { requestJson } = require('./http-client');
const {
  LOCAL_AI_PROVIDER,
  LOCAL_AI_MODEL,
  LOCAL_AI_BASE_URL,
  OLLAMA_RUNTIME_VERSION,
  OLLAMA_WINDOWS_ZIP_URL,
  HEALTHCHECK_TIMEOUT_MS,
  SETUP_SERVER_TIMEOUT_MS
} = require('./llm-config');
const { resolveLlmPerformanceProfile } = require('./llm-performance');

let electronApp = null;

try {
  const electron = require('electron');
  electronApp = electron?.app || null;
} catch {
  // Tests may load this module outside Electron.
}

function getUserDataRoot() {
  if (electronApp && typeof electronApp.getPath === 'function') {
    return electronApp.getPath('userData');
  }
  return path.join(process.cwd(), '.macroflow-local-ai');
}

const managedRoot = path.join(getUserDataRoot(), 'local-ai');
const storePath = path.join(managedRoot, 'local-ai.json');
const downloadsDir = path.join(managedRoot, 'downloads');
const runtimeDir = path.join(managedRoot, 'runtime', 'ollama');
const modelsDir = path.join(managedRoot, 'models');
const runtimeZipPath = path.join(downloadsDir, 'ollama-windows-amd64.zip');

const state = {
  snapshot: null,
  setupPromise: null,
  removePromise: null,
  listeners: new Set(),
  ollamaProcess: null
};

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function ensureManagedDirectories() {
  fs.mkdirSync(managedRoot, { recursive: true });
  fs.mkdirSync(downloadsDir, { recursive: true });
  fs.mkdirSync(modelsDir, { recursive: true });
}

function readStore() {
  try {
    if (fs.existsSync(storePath)) {
      return JSON.parse(fs.readFileSync(storePath, 'utf8'));
    }
  } catch (error) {
    logger.warn('[LocalAI] failed to read store', { error: error.message });
  }
  return {};
}

function writeStore(partial) {
  try {
    ensureManagedDirectories();
    const nextValue = {
      ...readStore(),
      ...partial,
      updatedAt: new Date().toISOString()
    };
    fs.writeFileSync(storePath, JSON.stringify(nextValue, null, 2), 'utf8');
  } catch (error) {
    logger.warn('[LocalAI] failed to write store', { error: error.message });
  }
}

function getSelectedModel() {
  const store = readStore();
  const configured = String(store.selectedModel || LOCAL_AI_MODEL).trim();
  return configured || LOCAL_AI_MODEL;
}

function getBaseUrl() {
  return LOCAL_AI_BASE_URL;
}

function getOllamaHost() {
  try {
    const parsed = new URL(LOCAL_AI_BASE_URL);
    return `${parsed.hostname}:${parsed.port || 11544}`;
  } catch {
    return '127.0.0.1:11544';
  }
}

// --- Runtime discovery ---

function findManagedRuntimePath() {
  const directPath = path.join(runtimeDir, 'ollama.exe');
  if (fs.existsSync(directPath)) {
    return directPath;
  }
  // Check one level deep (zip may have a subfolder)
  try {
    for (const entry of fs.readdirSync(runtimeDir, { withFileTypes: true })) {
      if (entry.isDirectory()) {
        const nested = path.join(runtimeDir, entry.name, 'ollama.exe');
        if (fs.existsSync(nested)) {
          return nested;
        }
      }
    }
  } catch {
    // Directory may not exist yet.
  }
  return '';
}

function findOllamaRuntime() {
  const envPath = String(process.env.MACROFLOW_OLLAMA_PATH || '').trim();
  if (envPath && fs.existsSync(envPath)) {
    return { installed: true, commandPath: path.normalize(envPath), source: 'env' };
  }

  const managedPath = findManagedRuntimePath();
  if (managedPath) {
    return { installed: true, commandPath: managedPath, source: 'managed' };
  }

  return { installed: false, commandPath: '', source: '' };
}

// --- Server health ---

async function pingOllamaServer() {
  try {
    const response = await requestJson({
      url: `${getBaseUrl()}/api/tags`,
      timeoutMs: HEALTHCHECK_TIMEOUT_MS
    });
    return response.statusCode >= 200 && response.statusCode < 300;
  } catch {
    return false;
  }
}

async function listInstalledModels() {
  try {
    const response = await requestJson({
      url: `${getBaseUrl()}/api/tags`,
      timeoutMs: HEALTHCHECK_TIMEOUT_MS
    });
    if (response.statusCode >= 200 && response.statusCode < 300 && Array.isArray(response.parsedBody?.models)) {
      return response.parsedBody.models;
    }
  } catch {
    // Server not reachable.
  }
  return [];
}

function isModelInstalled(models, targetModel) {
  const target = String(targetModel || '').trim().toLowerCase();
  if (!target) return false;
  return models.some((entry) => {
    const name = String(entry?.name || entry?.model || '').trim().toLowerCase();
    return name === target;
  });
}

// --- Snapshot / status ---

function buildSnapshot(overrides = {}) {
  const base = state.snapshot || {};
  return {
    success: true,
    provider: LOCAL_AI_PROVIDER,
    baseUrl: getBaseUrl(),
    model: getSelectedModel(),
    ready: false,
    needsSetup: true,
    setupInProgress: Boolean(state.setupPromise),
    removeInProgress: Boolean(state.removePromise),
    runtimeInstalled: false,
    serverReachable: false,
    modelInstalled: false,
    stage: 'runtime_missing',
    progress: null,
    statusText: 'Local AI setup is required.',
    lastError: '',
    ...base,
    ...overrides
  };
}

function setSnapshot(overrides = {}) {
  state.snapshot = buildSnapshot(overrides);
  for (const listener of state.listeners) {
    try {
      listener(state.snapshot);
    } catch {
      // Ignore listener failures.
    }
  }
  return state.snapshot;
}

function subscribe(listener) {
  if (typeof listener !== 'function') return () => {};
  state.listeners.add(listener);
  if (state.snapshot) listener(state.snapshot);
  return () => { state.listeners.delete(listener); };
}

function resolveStatusFields(runtime, serverReachable, modelInstalled) {
  const ready = Boolean(runtime.installed && serverReachable && modelInstalled);
  const stage = ready
    ? 'ready'
    : !runtime.installed
      ? 'runtime_missing'
      : !serverReachable
        ? 'runtime_not_running'
        : 'model_missing';
  const statusText = ready
    ? 'Local AI is ready.'
    : !runtime.installed
      ? 'The local AI runtime is not installed yet.'
      : !serverReachable
        ? 'The local AI runtime is installed but not running yet.'
        : `The local model ${getSelectedModel()} is not installed yet.`;
  return { ready, stage, statusText };
}

async function getStatus() {
  if (state.setupPromise || state.removePromise) {
    return buildSnapshot();
  }

  const runtime = findOllamaRuntime();
  const serverReachable = runtime.installed ? await pingOllamaServer() : false;
  const models = serverReachable ? await listInstalledModels() : [];
  const modelInstalled = serverReachable && isModelInstalled(models, getSelectedModel());
  const { ready, stage, statusText } = resolveStatusFields(runtime, serverReachable, modelInstalled);

  return setSnapshot({
    runtimeInstalled: runtime.installed,
    serverReachable,
    modelInstalled,
    ready,
    needsSetup: !ready,
    setupInProgress: false,
    removeInProgress: false,
    stage,
    progress: null,
    statusText,
    lastError: ''
  });
}

async function ensureReady() {
  if (state.setupPromise || state.removePromise) {
    return buildSnapshot();
  }

  const runtime = findOllamaRuntime();
  if (!runtime.installed) {
    return getStatus();
  }

  let serverReachable = await pingOllamaServer();

  if (!serverReachable) {
    try {
      await ensureServer(runtime);
      serverReachable = await pingOllamaServer();
    } catch (error) {
      logger.warn('[LocalAI] ensureReady auto-start failed', { error: error.message });
    }
  }

  const models = serverReachable ? await listInstalledModels() : [];
  const modelInstalled = serverReachable && isModelInstalled(models, getSelectedModel());
  const { ready, stage, statusText } = resolveStatusFields(runtime, serverReachable, modelInstalled);

  return setSnapshot({
    runtimeInstalled: true,
    serverReachable,
    modelInstalled,
    ready,
    needsSetup: !ready,
    setupInProgress: false,
    removeInProgress: false,
    stage,
    progress: null,
    statusText,
    lastError: ''
  });
}

// --- Download & extract ---

function downloadFile(url, destinationPath, onProgress, redirectCount = 0) {
  return new Promise((resolve, reject) => {
    if (redirectCount > 5) {
      reject(new Error('Too many download redirects.'));
      return;
    }

    ensureManagedDirectories();
    fs.rmSync(destinationPath, { force: true });
    const client = String(url || '').trim().toLowerCase().startsWith('https:') ? https : http;
    const request = client.get(url, (response) => {
      const statusCode = Number(response.statusCode) || 0;
      const location = String(response.headers.location || '').trim();

      if (statusCode >= 300 && statusCode < 400 && location) {
        response.resume();
        downloadFile(location, destinationPath, onProgress, redirectCount + 1).then(resolve, reject);
        return;
      }

      if (statusCode < 200 || statusCode >= 300) {
        response.resume();
        reject(new Error(`Download failed (HTTP ${statusCode}).`));
        return;
      }

      const totalBytes = Number(response.headers['content-length']) || 0;
      let downloadedBytes = 0;
      const file = fs.createWriteStream(destinationPath);

      response.on('data', (chunk) => {
        downloadedBytes += chunk.length;
        if (typeof onProgress === 'function') {
          onProgress({
            downloadedBytes,
            totalBytes,
            progress: totalBytes > 0 ? downloadedBytes / totalBytes : null
          });
        }
      });

      file.on('finish', () => {
        file.close(() => resolve(destinationPath));
      });

      file.on('error', (error) => {
        fs.rmSync(destinationPath, { force: true });
        file.close(() => reject(error));
      });

      response.pipe(file);
    });

    request.on('error', (error) => {
      fs.rmSync(destinationPath, { force: true });
      reject(error);
    });
  });
}

async function extractZipArchive(zipPath, destinationDir) {
  ensureManagedDirectories();
  fs.mkdirSync(destinationDir, { recursive: true });
  const escapedZip = String(zipPath).replace(/'/g, "''");
  const escapedDest = String(destinationDir).replace(/'/g, "''");

  return new Promise((resolve, reject) => {
    const child = spawn(
      'powershell.exe',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command',
        `Expand-Archive -LiteralPath '${escapedZip}' -DestinationPath '${escapedDest}' -Force`],
      { windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] }
    );
    let stderr = '';
    child.stderr.on('data', (chunk) => { stderr += String(chunk); });
    child.on('error', reject);
    child.on('close', (code) => {
      if (code === 0) resolve();
      else reject(new Error(stderr.trim() || `Extraction failed (exit ${code}).`));
    });
  });
}

async function downloadAndExtractRuntime() {
  ensureManagedDirectories();

  setSnapshot({
    stage: 'downloading_runtime',
    statusText: 'Downloading the local AI runtime...',
    progress: null
  });

  await downloadFile(OLLAMA_WINDOWS_ZIP_URL, runtimeZipPath, ({ progress }) => {
    setSnapshot({
      stage: 'downloading_runtime',
      statusText: 'Downloading the local AI runtime...',
      progress: typeof progress === 'number' ? progress : null
    });
  });

  setSnapshot({
    stage: 'extracting_runtime',
    statusText: 'Extracting the local AI runtime...',
    progress: null
  });

  fs.rmSync(runtimeDir, { recursive: true, force: true });
  fs.mkdirSync(runtimeDir, { recursive: true });
  await extractZipArchive(runtimeZipPath, runtimeDir);

  const runtimePath = findManagedRuntimePath();
  if (!runtimePath) {
    throw new Error('Extraction completed but ollama.exe was not found.');
  }

  return { installed: true, commandPath: runtimePath, source: 'managed' };
}

// --- Server lifecycle ---

function getManagedOllamaEnv() {
  ensureManagedDirectories();
  const performanceProfile = resolveLlmPerformanceProfile();
  const env = {
    ...process.env,
    OLLAMA_HOST: getOllamaHost(),
    OLLAMA_MODELS: modelsDir,
    OLLAMA_CONTEXT_LENGTH: String(performanceProfile.contextLength),
    OLLAMA_MAX_LOADED_MODELS: String(performanceProfile.ollamaMaxLoadedModels),
    OLLAMA_NUM_PARALLEL: String(performanceProfile.ollamaNumParallel),
    OLLAMA_KEEP_ALIVE: String(performanceProfile.ollamaKeepAlive),
    OLLAMA_NO_CLOUD: '1',
    OLLAMA_FLASH_ATTENTION: '1'
  };

  // Cap CPU threads so Ollama doesn't saturate every core.
  if (performanceProfile.ollamaNumThreads > 0) {
    env.OLLAMA_NUM_THREADS = String(performanceProfile.ollamaNumThreads);
  }

  // Force CPU-only on low/balanced tiers to avoid GPU memory exhaustion.
  if (performanceProfile.ollamaGpuLayers === 0) {
    env.OLLAMA_GPU_LAYERS = '0';
  }

  return env;
}

async function ensureServer(runtime) {
  if (await pingOllamaServer()) {
    return;
  }

  if (!runtime?.commandPath) {
    throw new Error('The local AI runtime is not installed.');
  }

  try {
    const child = spawn(runtime.commandPath, ['serve'], {
      detached: true,
      windowsHide: true,
      stdio: 'ignore',
      env: getManagedOllamaEnv()
    });
    child.unref();
    state.ollamaProcess = child;
    child.on('exit', () => { state.ollamaProcess = null; });
  } catch (error) {
    logger.warn('[LocalAI] failed to launch ollama serve', { error: error.message });
  }

  const startedAt = Date.now();
  while (Date.now() - startedAt < SETUP_SERVER_TIMEOUT_MS) {
    if (await pingOllamaServer()) return;
    await delay(1000);
  }

  throw new Error('The local AI runtime did not start in time.');
}

// --- Model pull ---

function friendlyPullStatus(rawStatus) {
  const status = String(rawStatus || '').trim().toLowerCase();
  if (!status || status.startsWith('pulling')) return 'Downloading...';
  if (status.includes('verifying')) return 'Verifying...';
  if (status.includes('writing')) return 'Installing...';
  if (status.includes('success')) return 'Installed.';
  return 'Setting up...';
}

function pullModel(modelName) {
  return new Promise((resolve, reject) => {
    const endpoint = `${getBaseUrl()}/api/pull`;
    const transport = endpoint.startsWith('https:') ? https : http;
    const request = transport.request(endpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    }, (response) => {
      if ((Number(response.statusCode) || 0) >= 400) {
        let errorText = '';
        response.setEncoding('utf8');
        response.on('data', (chunk) => { errorText += chunk; });
        response.on('end', () => {
          reject(new Error(errorText || `Model pull failed (HTTP ${response.statusCode}).`));
        });
        return;
      }

      response.setEncoding('utf8');
      let buffer = '';

      response.on('data', (chunk) => {
        buffer += chunk;
        const lines = buffer.split(/\r?\n/);
        buffer = lines.pop() || '';

        for (const line of lines) {
          if (!line.trim()) continue;
          try {
            const parsed = JSON.parse(line.trim());
            const total = Number(parsed?.total) || 0;
            const completed = Number(parsed?.completed) || 0;
            setSnapshot({
              stage: 'downloading_model',
              statusText: friendlyPullStatus(parsed?.status),
              progress: total > 0 ? completed / total : null
            });
          } catch {
            // Ignore malformed progress lines.
          }
        }
      });

      response.on('end', () => {
        if (buffer.trim()) {
          try {
            const parsed = JSON.parse(buffer.trim());
            if (parsed?.error) {
              reject(new Error(String(parsed.error)));
              return;
            }
          } catch {
            // Ignore trailing non-JSON.
          }
        }
        resolve();
      });
    });

    request.setTimeout(SETUP_SERVER_TIMEOUT_MS + 5 * 60 * 1000, () => {
      request.destroy(new Error('Model download timed out.'));
    });

    request.on('error', reject);
    request.write(JSON.stringify({ model: modelName, stream: true }));
    request.end();
  });
}

// --- Setup flow ---

async function runSetupFlow() {
  const model = getSelectedModel();
  writeStore({ selectedModel: model, lastError: '' });
  setSnapshot({
    setupInProgress: true,
    stage: 'checking',
    progress: null,
    statusText: 'Checking local AI runtime...',
    lastError: ''
  });

  try {
    let runtime = findOllamaRuntime();

    if (!runtime.installed) {
      runtime = await downloadAndExtractRuntime();
    }

    setSnapshot({
      setupInProgress: true,
      runtimeInstalled: true,
      stage: 'starting_runtime',
      progress: null,
      statusText: 'Starting the local AI runtime...'
    });

    await ensureServer(runtime);

    const models = await listInstalledModels();
    if (!isModelInstalled(models, model)) {
      setSnapshot({
        setupInProgress: true,
        runtimeInstalled: true,
        serverReachable: true,
        modelInstalled: false,
        stage: 'downloading_model',
        progress: null,
        statusText: 'Downloading...'
      });
      await pullModel(model);
    }

    setSnapshot({
      ready: true,
      needsSetup: false,
      setupInProgress: false,
      runtimeInstalled: true,
      serverReachable: true,
      modelInstalled: true,
      stage: 'ready',
      progress: null,
      statusText: 'Local AI is ready.',
      lastError: ''
    });
  } catch (error) {
    const message = String(error?.message || 'Local AI setup failed.').trim();
    writeStore({ lastError: message });
    setSnapshot({
      setupInProgress: false,
      ready: false,
      needsSetup: true,
      progress: null,
      stage: 'error',
      statusText: message,
      lastError: message
    });
    logger.error('[LocalAI] setup failed', { error: message });
  } finally {
    state.setupPromise = null;
  }
}

// --- Remove model flow ---

async function runRemoveModelFlow() {
  const model = getSelectedModel();
  writeStore({ selectedModel: model, lastError: '' });
  setSnapshot({
    removeInProgress: true,
    progress: 0.1,
    stage: 'removing_model',
    statusText: 'Removing local model...',
    lastError: ''
  });

  try {
    const runtime = findOllamaRuntime();
    if (runtime.installed) {
      await ensureServer(runtime);

      setSnapshot({
        removeInProgress: true,
        progress: 0.5,
        stage: 'removing_model',
        statusText: 'Deleting model files...'
      });

      const models = await listInstalledModels();
      if (isModelInstalled(models, model)) {
        const response = await requestJson({
          url: `${getBaseUrl()}/api/delete`,
          method: 'DELETE',
          bodyObject: { model },
          timeoutMs: SETUP_SERVER_TIMEOUT_MS
        });
        if (response.statusCode < 200 || response.statusCode >= 300) {
          const apiError = String(response.parsedBody?.error || response.bodyText || '').trim();
          throw new Error(apiError || `Model delete failed (HTTP ${response.statusCode}).`);
        }
      }
    }
  } catch (error) {
    const message = String(error?.message || 'Unable to remove the local AI model.').trim();
    writeStore({ lastError: message });
    setSnapshot({
      removeInProgress: false,
      ready: false,
      needsSetup: true,
      progress: null,
      stage: 'error',
      statusText: message,
      lastError: message
    });
    logger.error('[LocalAI] remove model failed', { error: message });
    return;
  } finally {
    state.removePromise = null;
  }

  writeStore({ lastError: '' });
  setSnapshot({
    removeInProgress: false,
    ready: false,
    needsSetup: true,
    modelInstalled: false,
    progress: null,
    stage: 'model_missing',
    statusText: 'Local AI model removed.',
    lastError: ''
  });
}

// --- Public API ---

async function setup() {
  if (state.setupPromise || state.removePromise) {
    return { success: true, started: false, status: buildSnapshot() };
  }

  const status = await getStatus();
  if (status.ready) {
    return { success: true, started: false, status };
  }

  state.setupPromise = runSetupFlow();
  return {
    success: true,
    started: true,
    status: buildSnapshot({
      setupInProgress: true,
      stage: 'checking',
      statusText: 'Checking local AI runtime...'
    })
  };
}

async function removeModel() {
  if (state.removePromise || state.setupPromise) {
    return { success: true, started: false, status: buildSnapshot() };
  }

  state.removePromise = runRemoveModelFlow();
  return {
    success: true,
    started: true,
    status: buildSnapshot({
      removeInProgress: true,
      progress: null,
      stage: 'removing_model',
      statusText: `Removing ${getSelectedModel()}...`
    })
  };
}

function shutdown() {
  if (state.ollamaProcess) {
    try {
      state.ollamaProcess.kill();
      logger.info('[LocalAI] Ollama process killed on shutdown');
    } catch (error) {
      logger.warn('[LocalAI] failed to kill Ollama process', { error: error.message });
    }
    state.ollamaProcess = null;
  }
}

module.exports = {
  getStatus,
  ensureReady,
  setup,
  removeModel,
  subscribe,
  shutdown
};
