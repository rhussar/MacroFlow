const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');

let worker = null;
let workerStartPromise = null;
let buffer = '';
let nextId = 1;
const pending = new Map();

function getWorkerPath() {
  const isDev = process.env.NODE_ENV === 'development' || Boolean(process.defaultApp);
  if (!isDev && process.resourcesPath) {
    return path.join(process.resourcesPath, 'llm', 'llm-worker.mjs');
  }
  return path.join(__dirname, 'llm-worker.mjs');
}

function getNodePath() {
  const exeName = process.platform === 'win32' ? 'node.exe' : 'node';
  const bundledPath = process.resourcesPath
    ? path.join(process.resourcesPath, 'llm', exeName)
    : null;

  if (bundledPath && fs.existsSync(bundledPath)) {
    return bundledPath;
  }

  return 'node';
}

function startWorker() {
  if (workerStartPromise) {
    return workerStartPromise;
  }

  workerStartPromise = new Promise((resolve, reject) => {
    const workerPath = getWorkerPath();
    if (!fs.existsSync(workerPath)) {
      workerStartPromise = null;
      reject(new Error(`LLM worker not found at ${workerPath}`));
      return;
    }

    const nodePath = getNodePath();

    const env = { ...process.env };
    if (process.resourcesPath && !env.NODE_PATH) {
      const unpackedModules = path.join(process.resourcesPath, 'app.asar.unpacked', 'node_modules');
      if (fs.existsSync(unpackedModules)) {
        env.NODE_PATH = unpackedModules;
      }
    } else if (process.resourcesPath && env.NODE_PATH) {
      const unpackedModules = path.join(process.resourcesPath, 'app.asar.unpacked', 'node_modules');
      if (fs.existsSync(unpackedModules)) {
        env.NODE_PATH = `${unpackedModules}${path.delimiter}${env.NODE_PATH}`;
      }
    }

    worker = spawn(nodePath, [workerPath], {
      stdio: ['pipe', 'pipe', 'pipe'],
      env,
    });

    worker.on('error', (err) => {
      worker = null;
      workerStartPromise = null;
      reject(err);
    });

    worker.on('exit', (code, signal) => {
      const reason = signal ? `signal ${signal}` : `code ${code}`;
      const error = new Error(`LLM worker exited (${reason})`);
      for (const { reject: rejectPending } of pending.values()) {
        rejectPending(error);
      }
      pending.clear();
      worker = null;
      workerStartPromise = null;
    });

    worker.stdout.setEncoding('utf8');
    worker.stdout.on('data', (chunk) => {
      buffer += chunk;
      let index = buffer.indexOf('\n');
      while (index !== -1) {
        const line = buffer.slice(0, index).trim();
        buffer = buffer.slice(index + 1);
        if (line) {
          handleWorkerMessage(line);
        }
        index = buffer.indexOf('\n');
      }
    });

    worker.stderr.setEncoding('utf8');
    worker.stderr.on('data', (chunk) => {
      const message = chunk.trim();
      if (message) {
        console.error('[LLM Worker]', message);
      }
    });

    resolve();
  });

  return workerStartPromise;
}

function handleWorkerMessage(line) {
  let message;
  try {
    message = JSON.parse(line);
  } catch (err) {
    console.error('[LLM Worker] Invalid JSON:', line);
    return;
  }

  const { id, success, data, error } = message;
  if (!id || !pending.has(id)) {
    return;
  }

  const { resolve, reject } = pending.get(id);
  pending.delete(id);

  if (success) {
    resolve({ success: true, data });
  } else {
    reject(new Error(error || 'Unknown LLM worker error'));
  }
}

async function sendCommand(type, payload = {}) {
  await startWorker();

  if (!worker || !worker.stdin) {
    throw new Error('LLM worker is not running.');
  }

  const id = nextId++;
  const message = { id, type, ...payload };

  return new Promise((resolve, reject) => {
    pending.set(id, { resolve, reject });
    worker.stdin.write(`${JSON.stringify(message)}\n`);
  });
}

/**
 * Initialize the Local LLM Engine
 * @param {string} modelPath - Absolute path to the .gguf model file
 */
async function initLLM(modelPath) {
  try {
    const result = await sendCommand('init', { modelPath });
    return { success: true, ...result };
  } catch (error) {
    console.error('LLM Init Failed:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Generate a response from the AI
 * @param {string} userPrompt - The user's question
 * @param {object} excelContext - JSON object containing active sheet/selection data
 */
async function chat(userPrompt, excelContext = '') {
  const result = await sendCommand('chat', { prompt: userPrompt, context: excelContext });
  return result.data;
}

module.exports = { initLLM, chat };
