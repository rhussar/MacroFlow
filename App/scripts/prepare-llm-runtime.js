const fs = require('fs');
const path = require('path');

const appRoot = path.resolve(__dirname, '..');
const resourcesDir = path.join(appRoot, 'Resources');
const llmDir = path.join(resourcesDir, 'llm');
const modelsSrcDir = path.join(appRoot, 'models');
const modelsDestDir = path.join(resourcesDir, 'models');
const workerSrc = path.join(appRoot, 'electron', 'llm-worker.mjs');
const workerDest = path.join(llmDir, 'llm-worker.mjs');
const llamaDep = path.join(appRoot, 'node_modules', 'node-llama-cpp');

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

function copyIfChanged(src, dest) {
  if (!fs.existsSync(src)) {
    throw new Error(`Missing file: ${src}`);
  }

  let shouldCopy = true;
  if (fs.existsSync(dest)) {
    const srcStat = fs.statSync(src);
    const destStat = fs.statSync(dest);
    shouldCopy = srcStat.size !== destStat.size || srcStat.mtimeMs > destStat.mtimeMs;
  }

  if (shouldCopy) {
    ensureDir(path.dirname(dest));
    fs.copyFileSync(src, dest);
    console.log(`[LLM] Copied ${src} -> ${dest}`);
  } else {
    console.log(`[LLM] Up-to-date: ${dest}`);
  }
}

function copyModels() {
  if (!fs.existsSync(modelsSrcDir)) {
    console.warn(`[LLM] Models directory not found: ${modelsSrcDir}`);
    return;
  }

  const files = fs.readdirSync(modelsSrcDir).filter((name) => name.endsWith('.gguf'));
  if (files.length === 0) {
    console.warn(`[LLM] No .gguf models found in ${modelsSrcDir}`);
    return;
  }

  ensureDir(modelsDestDir);
  for (const file of files) {
    const src = path.join(modelsSrcDir, file);
    const dest = path.join(modelsDestDir, file);
    copyIfChanged(src, dest);
  }
}

function copyNodeRuntime() {
  const nodeSrc = process.execPath;
  const nodeDest = path.join(llmDir, path.basename(nodeSrc));
  copyIfChanged(nodeSrc, nodeDest);
}

function main() {
  ensureDir(resourcesDir);
  ensureDir(llmDir);

  if (!fs.existsSync(llamaDep)) {
    throw new Error('node-llama-cpp is not installed in App/node_modules. Run npm install in App first.');
  }

  copyIfChanged(workerSrc, workerDest);
  copyNodeRuntime();
  copyModels();
}

try {
  main();
} catch (err) {
  console.error(`[LLM] Prepare failed: ${err.message}`);
  process.exit(1);
}
