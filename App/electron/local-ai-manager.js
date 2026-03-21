const fs = require('node:fs');
const path = require('node:path');
const http = require('node:http');
const https = require('node:https');
const crypto = require('node:crypto');
const { spawn, spawnSync } = require('node:child_process');

const logger = require('./logger');
const { requestJson } = require('./http-client');
const {
  LOCAL_AI_PROVIDER,
  LOCAL_AI_MODEL,
  LOCAL_AI_BASE_URL,
  LOCAL_AI_CONTEXT_LENGTH,
  ALLOW_EXTERNAL_OLLAMA,
  OLLAMA_RUNTIME_VERSION,
  OLLAMA_WINDOWS_ZIP_URL,
  OLLAMA_WINDOWS_ZIP_SHA256,
  OLLAMA_WINDOWS_ROCM_ZIP_URL,
  OLLAMA_WINDOWS_ROCM_ZIP_SHA256,
  HEALTHCHECK_TIMEOUT_MS,
  SETUP_SERVER_TIMEOUT_MS
} = require('./llm-config');

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
const runtimeRocmZipPath = path.join(downloadsDir, 'ollama-windows-amd64-rocm.zip');

const state = {
  snapshot: null,
  setupPromise: null,
  removePromise: null,
  listeners: new Set()
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

function normalizeRuntimePath(value) {
  const normalized = String(value || '').trim();
  return normalized ? path.normalize(normalized) : '';
}

function findFileRecursive(rootDir, fileName, maxDepth = 4) {
  if (!rootDir || !fs.existsSync(rootDir) || maxDepth < 0) {
    return '';
  }

  try {
    const entries = fs.readdirSync(rootDir, { withFileTypes: true });
    for (const entry of entries) {
      const nextPath = path.join(rootDir, entry.name);
      if (entry.isFile() && entry.name.toLowerCase() === fileName.toLowerCase()) {
        return nextPath;
      }
    }

    for (const entry of entries) {
      if (!entry.isDirectory()) {
        continue;
      }
      const nested = findFileRecursive(path.join(rootDir, entry.name), fileName, maxDepth - 1);
      if (nested) {
        return nested;
      }
    }
  } catch {
    return '';
  }

  return '';
}

function findManagedRuntimePath() {
  const directPath = path.join(runtimeDir, 'ollama.exe');
  if (fs.existsSync(directPath)) {
    return directPath;
  }
  return findFileRecursive(runtimeDir, 'ollama.exe', 3);
}

function findOllamaRuntime() {
  const candidates = [];
  const envPath = normalizeRuntimePath(process.env.MACROFLOW_OLLAMA_PATH);
  const managedPath = normalizeRuntimePath(findManagedRuntimePath());

  if (envPath) {
    candidates.push({ commandPath: envPath, source: 'env' });
  }

  if (managedPath) {
    candidates.push({ commandPath: managedPath, source: 'managed' });
  }

  if (ALLOW_EXTERNAL_OLLAMA) {
    const localAppData = String(process.env.LOCALAPPDATA || '').trim();
    if (localAppData) {
      candidates.push({
        commandPath: path.join(localAppData, 'Programs', 'Ollama', 'ollama.exe'),
        source: 'default-install'
      });
    }
  }

  for (const candidate of candidates) {
    if (candidate.commandPath && fs.existsSync(candidate.commandPath)) {
      return {
        installed: true,
        commandPath: candidate.commandPath,
        source: candidate.source
      };
    }
  }

  if (ALLOW_EXTERNAL_OLLAMA) {
    try {
      const result = spawnSync('where.exe', ['ollama'], {
        encoding: 'utf8',
        windowsHide: true
      });
      if (result.status === 0) {
        const commandPath = normalizeRuntimePath(String(result.stdout || '').split(/\r?\n/)[0]);
        if (commandPath && fs.existsSync(commandPath)) {
          return {
            installed: true,
            commandPath,
            source: 'path'
          };
        }
      }
    } catch {
      // Best-effort lookup only.
    }
  }

  return {
    installed: false,
    commandPath: '',
    source: ''
  };
}

function getOllamaHostFromBaseUrl() {
  try {
    return new URL(LOCAL_AI_BASE_URL).host;
  } catch {
    return '127.0.0.1:11434';
  }
}

function getOllamaPortFromBaseUrl() {
  try {
    const url = new URL(LOCAL_AI_BASE_URL);
    if (url.port) {
      return Number(url.port);
    }
    return url.protocol === 'https:' ? 443 : 80;
  } catch {
    return 11434;
  }
}

function getManagedOllamaEnv() {
  ensureManagedDirectories();
  return {
    ...process.env,
    OLLAMA_HOST: getOllamaHostFromBaseUrl(),
    OLLAMA_MODELS: modelsDir,
    OLLAMA_CONTEXT_LENGTH: String(LOCAL_AI_CONTEXT_LENGTH),
    OLLAMA_NO_CLOUD: '1'
  };
}

function getListeningProcessPath(port) {
  if (process.platform !== 'win32' || !Number.isFinite(port) || port <= 0) {
    return '';
  }

  try {
    const result = spawnSync(
      'powershell.exe',
      [
        '-NoProfile',
        '-NonInteractive',
        '-ExecutionPolicy',
        'Bypass',
        '-Command',
        [
          `$connection = Get-NetTCPConnection -State Listen -LocalPort ${Math.trunc(port)} -ErrorAction SilentlyContinue | Select-Object -First 1;`,
          'if (-not $connection) { return }',
          '$process = Get-Process -Id $connection.OwningProcess -ErrorAction SilentlyContinue;',
          'if ($process -and $process.Path) { Write-Output $process.Path }'
        ].join(' ')
      ],
      {
        encoding: 'utf8',
        windowsHide: true
      }
    );

    if (result.status !== 0) {
      return '';
    }

    return normalizeRuntimePath(String(result.stdout || '').trim());
  } catch {
    return '';
  }
}

async function getServerState(runtime) {
  const serverReachable = await pingOllamaServer();
  if (!serverReachable) {
    return {
      serverReachable: false,
      serverOwnedByRuntime: false,
      listenerProcessPath: ''
    };
  }

  const listenerProcessPath = getListeningProcessPath(getOllamaPortFromBaseUrl());
  if (ALLOW_EXTERNAL_OLLAMA || process.platform !== 'win32') {
    return {
      serverReachable: true,
      serverOwnedByRuntime: true,
      listenerProcessPath
    };
  }

  const expectedRuntimePath = normalizeRuntimePath(runtime?.commandPath).toLowerCase();
  const normalizedListenerPath = normalizeRuntimePath(listenerProcessPath).toLowerCase();
  return {
    serverReachable: true,
    serverOwnedByRuntime: Boolean(expectedRuntimePath && normalizedListenerPath && expectedRuntimePath === normalizedListenerPath),
    listenerProcessPath
  };
}

async function pingOllamaServer(timeoutMs = HEALTHCHECK_TIMEOUT_MS) {
  try {
    const response = await requestJson({
      url: `${LOCAL_AI_BASE_URL.replace(/\/+$/, '')}/api/tags`,
      timeoutMs
    });
    return response.statusCode >= 200 && response.statusCode < 300;
  } catch (error) {
    if (error?.code && error.code !== 'ECONNREFUSED') {
      logger.debug('[LocalAI] ping failed', { code: error.code, message: error.message });
    }
    return false;
  }
}

async function listInstalledModels() {
  const response = await requestJson({
    url: `${LOCAL_AI_BASE_URL.replace(/\/+$/, '')}/api/tags`,
    timeoutMs: HEALTHCHECK_TIMEOUT_MS
  });

  if (response.statusCode < 200 || response.statusCode >= 300 || !Array.isArray(response.parsedBody?.models)) {
    return [];
  }

  return response.parsedBody.models;
}

function isModelInstalled(models, targetModel) {
  const normalizedTarget = String(targetModel || '').trim().toLowerCase();
  if (!normalizedTarget) {
    return false;
  }

  return models.some((entry) => {
    const name = String(entry?.name || entry?.model || '').trim().toLowerCase();
    return name === normalizedTarget;
  });
}

function getStatusText({ runtimeInstalled, serverReachable, serverOwnedByRuntime, modelInstalled, ready, lastError }) {
  if (ready) {
    return 'Local AI is ready.';
  }
  if (lastError) {
    return lastError;
  }
  if (!runtimeInstalled) {
    return 'The local AI runtime is not installed yet.';
  }
  if (!serverReachable) {
    return 'The local AI runtime is installed but not running yet.';
  }
  if (!serverOwnedByRuntime) {
    return 'Another Ollama instance is already using MacroFlow\'s local AI port.';
  }
  if (!modelInstalled) {
    return `The local model ${getSelectedModel()} is not installed yet.`;
  }
  return 'Local AI setup is required.';
}

function buildSnapshot(overrides = {}) {
  const base = state.snapshot || {};
  return {
    success: true,
    provider: LOCAL_AI_PROVIDER,
    baseUrl: LOCAL_AI_BASE_URL,
    model: getSelectedModel(),
    ready: false,
    needsSetup: true,
    setupInProgress: Boolean(state.setupPromise),
    removeInProgress: Boolean(state.removePromise),
    runtimeInstalled: false,
    runtimeCommand: '',
    runtimeSource: '',
    serverReachable: false,
    serverOwnedByRuntime: false,
    listenerProcessPath: '',
    modelInstalled: false,
    localOnly: true,
    cloudFeaturesDisabled: true,
    contextLength: LOCAL_AI_CONTEXT_LENGTH,
    runtimeVersion: OLLAMA_RUNTIME_VERSION,
    externalRuntimeAllowed: ALLOW_EXTERNAL_OLLAMA,
    runtimeDirectory: runtimeDir,
    modelsDirectory: modelsDir,
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

async function getStatus() {
  if (state.setupPromise || state.removePromise) {
    return buildSnapshot();
  }

  // If we already know the model is ready, return cached snapshot immediately.
  // The snapshot is kept current by setup/remove flows and subscriptions.
  if (state.snapshot?.ready) {
    return state.snapshot;
  }

  const runtime = findOllamaRuntime();
  const serverState = runtime.installed
    ? await getServerState(runtime)
    : { serverReachable: false, serverOwnedByRuntime: false, listenerProcessPath: '' };
  const models = serverState.serverReachable && serverState.serverOwnedByRuntime ? await listInstalledModels() : [];
  const modelInstalled = serverState.serverReachable && serverState.serverOwnedByRuntime
    ? isModelInstalled(models, getSelectedModel())
    : false;
  const ready = Boolean(runtime.installed && serverState.serverReachable && serverState.serverOwnedByRuntime && modelInstalled);

  return setSnapshot({
    runtimeInstalled: runtime.installed,
    runtimeCommand: runtime.commandPath,
    runtimeSource: runtime.source,
    serverReachable: serverState.serverReachable,
    serverOwnedByRuntime: serverState.serverOwnedByRuntime,
    listenerProcessPath: serverState.listenerProcessPath,
    modelInstalled,
    ready,
    needsSetup: !ready,
    setupInProgress: false,
    removeInProgress: false,
    stage: ready
      ? 'ready'
      : !runtime.installed
        ? 'runtime_missing'
        : !serverState.serverReachable
          ? 'runtime_not_running'
          : !serverState.serverOwnedByRuntime
            ? 'runtime_conflict'
          : 'model_missing',
    progress: null,
    statusText: getStatusText({
      runtimeInstalled: runtime.installed,
      serverReachable: serverState.serverReachable,
      serverOwnedByRuntime: serverState.serverOwnedByRuntime,
      modelInstalled,
      ready,
      lastError: ''
    }),
    lastError: ''
  });
}

function subscribe(listener) {
  if (typeof listener !== 'function') {
    return () => {};
  }

  state.listeners.add(listener);
  if (state.snapshot) {
    listener(state.snapshot);
  }

  return () => {
    state.listeners.delete(listener);
  };
}

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

function createSha256HashForFile(filePath) {
  return new Promise((resolve, reject) => {
    const hash = crypto.createHash('sha256');
    const stream = fs.createReadStream(filePath);

    stream.on('error', reject);
    stream.on('data', (chunk) => {
      hash.update(chunk);
    });
    stream.on('end', () => {
      resolve(hash.digest('hex'));
    });
  });
}

async function verifyDownloadChecksum(filePath, expectedDigest, label) {
  const normalizedExpected = String(expectedDigest || '').trim().toLowerCase().replace(/^sha256:/, '');
  if (!normalizedExpected) {
    return;
  }

  const actualDigest = await createSha256HashForFile(filePath);
  if (actualDigest !== normalizedExpected) {
    fs.rmSync(filePath, { force: true });
    throw new Error(`${label} failed checksum verification.`);
  }
}

function escapePowerShellLiteral(value) {
  return String(value || '').replace(/'/g, "''");
}

function runPowerShell(command) {
  return new Promise((resolve, reject) => {
    const child = spawn(
      'powershell.exe',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', command],
      {
        windowsHide: true,
        stdio: ['ignore', 'pipe', 'pipe']
      }
    );

    let stdout = '';
    let stderr = '';

    child.stdout.on('data', (chunk) => {
      stdout += String(chunk || '');
    });

    child.stderr.on('data', (chunk) => {
      stderr += String(chunk || '');
    });

    child.on('error', reject);
    child.on('close', (code) => {
      if (code === 0) {
        resolve({ stdout, stderr });
        return;
      }
      reject(new Error(stderr.trim() || stdout.trim() || `PowerShell failed with exit code ${code}.`));
    });
  });
}

async function extractZipArchive(zipPath, destinationDir) {
  ensureManagedDirectories();
  fs.mkdirSync(destinationDir, { recursive: true });
  const escapedZipPath = escapePowerShellLiteral(zipPath);
  const escapedDestination = escapePowerShellLiteral(destinationDir);
  await runPowerShell(
    `Expand-Archive -LiteralPath '${escapedZipPath}' -DestinationPath '${escapedDestination}' -Force`
  );
}

function detectAmdGpu() {
  if (process.platform !== 'win32') {
    return false;
  }

  try {
    const result = spawnSync(
      'powershell.exe',
      [
        '-NoProfile',
        '-NonInteractive',
        '-ExecutionPolicy',
        'Bypass',
        '-Command',
        'Get-CimInstance Win32_VideoController | Select-Object -ExpandProperty Name'
      ],
      {
        encoding: 'utf8',
        windowsHide: true
      }
    );

    if (result.status !== 0) {
      return false;
    }

    const output = String(result.stdout || '').toLowerCase();
    return output.includes('amd') || output.includes('radeon');
  } catch {
    return false;
  }
}

async function downloadAndExtractManagedRuntime() {
  ensureManagedDirectories();
  const hasAmdGpu = detectAmdGpu();

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
    stage: 'verifying_runtime',
    statusText: `Verifying the local AI runtime (${OLLAMA_RUNTIME_VERSION})...`,
    progress: null
  });
  await verifyDownloadChecksum(runtimeZipPath, OLLAMA_WINDOWS_ZIP_SHA256, 'The local AI runtime archive');

  if (hasAmdGpu) {
    await downloadFile(OLLAMA_WINDOWS_ROCM_ZIP_URL, runtimeRocmZipPath, ({ progress }) => {
      setSnapshot({
        stage: 'downloading_runtime_gpu',
        statusText: 'Downloading AMD GPU support for the local AI runtime...',
        progress: typeof progress === 'number' ? progress : null
      });
    });
    setSnapshot({
      stage: 'verifying_runtime_gpu',
      statusText: 'Verifying AMD GPU support for the local AI runtime...',
      progress: null
    });
    await verifyDownloadChecksum(runtimeRocmZipPath, OLLAMA_WINDOWS_ROCM_ZIP_SHA256, 'The AMD GPU support archive');
  }

  setSnapshot({
    stage: 'extracting_runtime',
    statusText: hasAmdGpu
      ? 'Extracting the local AI runtime and AMD GPU support...'
      : 'Extracting the local AI runtime...',
    progress: null
  });

  fs.rmSync(runtimeDir, { recursive: true, force: true });
  fs.mkdirSync(runtimeDir, { recursive: true });
  await extractZipArchive(runtimeZipPath, runtimeDir);

  if (hasAmdGpu && fs.existsSync(runtimeRocmZipPath)) {
    await extractZipArchive(runtimeRocmZipPath, runtimeDir);
  }

  const runtimePath = findManagedRuntimePath();
  if (!runtimePath) {
    throw new Error('The local AI runtime was extracted, but MacroFlow could not find ollama.exe.');
  }

  return {
    installed: true,
    commandPath: runtimePath,
    source: 'managed'
  };
}

async function ensureServer(runtime) {
  const initialServerState = await getServerState(runtime);
  if (initialServerState.serverReachable && initialServerState.serverOwnedByRuntime) {
    return true;
  }

  if (initialServerState.serverReachable && !initialServerState.serverOwnedByRuntime) {
    throw new Error('Another Ollama instance is already using MacroFlow\'s local AI port. Close it or opt in to an external Ollama runtime.');
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
  } catch (error) {
    logger.warn('[LocalAI] failed to launch ollama serve', { error: error.message });
  }

  const startedAt = Date.now();
  while (Date.now() - startedAt < SETUP_SERVER_TIMEOUT_MS) {
    const serverState = await getServerState(runtime);
    if (serverState.serverReachable && serverState.serverOwnedByRuntime) {
      return true;
    }
    if (serverState.serverReachable && !serverState.serverOwnedByRuntime) {
      throw new Error('Another Ollama instance is already using MacroFlow\'s local AI port. Close it or opt in to an external Ollama runtime.');
    }
    await delay(1000);
  }

  throw new Error('The local AI runtime did not start in time.');
}

async function deleteInstalledModel(modelName) {
  const response = await requestJson({
    url: `${LOCAL_AI_BASE_URL.replace(/\/+$/, '')}/api/delete`,
    method: 'DELETE',
    bodyObject: {
      model: modelName
    },
    timeoutMs: SETUP_SERVER_TIMEOUT_MS
  });

  if (response.statusCode < 200 || response.statusCode >= 300) {
    const apiError = String(response.parsedBody?.error || response.bodyText || '').trim();
    throw new Error(apiError || `Model delete failed (HTTP ${response.statusCode}).`);
  }
}

function pullModel(modelName) {
  return new Promise((resolve, reject) => {
    const endpoint = `${LOCAL_AI_BASE_URL.replace(/\/+$/, '')}/api/pull`;
    const transport = String(endpoint).trim().toLowerCase().startsWith('https:') ? https : http;
    const request = transport.request(
      endpoint,
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        }
      },
      (response) => {
        if ((Number(response.statusCode) || 0) >= 400) {
          let errorText = '';
          response.setEncoding('utf8');
          response.on('data', (chunk) => {
            errorText += chunk;
          });
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
            const trimmed = line.trim();
            if (!trimmed) {
              continue;
            }

            try {
              const parsed = JSON.parse(trimmed);
              const total = Number(parsed?.total) || 0;
              const completed = Number(parsed?.completed) || 0;
              setSnapshot({
                stage: 'downloading_model',
                statusText: String(parsed?.status || `Downloading ${modelName}...`),
                progress: total > 0 ? completed / total : null
              });
            } catch {
              // Ignore malformed progress chunks.
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
      }
    );

    request.setTimeout(SETUP_SERVER_TIMEOUT_MS + 5 * 60 * 1000, () => {
      request.destroy(new Error('Model download timed out.'));
    });

    request.on('error', reject);
    request.write(JSON.stringify({ model: modelName, stream: true }));
    request.end();
  });
}

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
      runtime = await downloadAndExtractManagedRuntime();
    }

    setSnapshot({
      setupInProgress: true,
      runtimeInstalled: true,
      runtimeCommand: runtime.commandPath,
      runtimeSource: runtime.source,
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
        runtimeCommand: runtime.commandPath,
        runtimeSource: runtime.source,
        serverReachable: true,
        serverOwnedByRuntime: true,
        listenerProcessPath: runtime.commandPath,
        modelInstalled: false,
        stage: 'downloading_model',
        progress: null,
        statusText: `Downloading ${model}...`
      });
      await pullModel(model);
    }

    setSnapshot({
      ready: true,
      needsSetup: false,
      setupInProgress: false,
      runtimeInstalled: true,
      runtimeCommand: runtime.commandPath,
      runtimeSource: runtime.source,
      serverReachable: true,
      serverOwnedByRuntime: true,
      listenerProcessPath: runtime.commandPath,
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

async function runRemoveModelFlow() {
  const model = getSelectedModel();
  writeStore({ selectedModel: model, lastError: '' });
  setSnapshot({
    removeInProgress: true,
    progress: 0.08,
    stage: 'removing_model',
    statusText: 'Preparing removal...',
    lastError: ''
  });

  try {
    const runtime = findOllamaRuntime();
    if (runtime.installed) {
      setSnapshot({
        removeInProgress: true,
        progress: 0.24,
        stage: 'removing_model',
        statusText: 'Connecting to local AI...'
      });
      await ensureServer(runtime);
      setSnapshot({
        removeInProgress: true,
        progress: 0.52,
        stage: 'removing_model',
        statusText: 'Checking installed model...'
      });
      const serverState = await getServerState(runtime);
      const models = serverState.serverReachable && serverState.serverOwnedByRuntime
        ? await listInstalledModels()
        : [];
      if (isModelInstalled(models, model)) {
        setSnapshot({
          removeInProgress: true,
          progress: 0.82,
          stage: 'removing_model',
          statusText: 'Removing local model...'
        });
        await deleteInstalledModel(model);
      }
    }
    setSnapshot({
      removeInProgress: true,
      progress: 0.96,
      stage: 'removing_model',
      statusText: 'Finalizing...'
    });
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
  await getStatus();
}

async function setup() {
  if (state.setupPromise || state.removePromise) {
    return {
      success: true,
      started: false,
      status: buildSnapshot()
    };
  }

  const status = await getStatus();
  if (status.ready) {
    return {
      success: true,
      started: false,
      status
    };
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
    return {
      success: true,
      started: false,
      status: buildSnapshot()
    };
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

module.exports = {
  getStatus,
  setup,
  removeModel,
  subscribe
};
