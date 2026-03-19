#!/usr/bin/env node

/**
 * publish-update.js — Build MacroFlow and upload Squirrel artifacts to R2/S3.
 *
 * Usage:
 *   node scripts/publish-update.js           (build + upload)
 *   node scripts/publish-update.js --upload   (upload only, skip build — use if you already ran `npm run make`)
 *
 * Required environment variables (set in .env or your shell):
 *   AWS_ACCESS_KEY_ID        — R2 API token access key
 *   AWS_SECRET_ACCESS_KEY    — R2 API token secret key
 *   S3_ENDPOINT              — R2 endpoint, e.g. https://<account-id>.r2.cloudflarestorage.com
 *   S3_BUCKET                — Bucket name, e.g. macroflow-releases
 *
 * What gets uploaded:
 *   out/make/squirrel.windows/x64/RELEASES                → s3://{bucket}/updates/RELEASES
 *   out/make/squirrel.windows/x64/<name>-<ver>-full.nupkg  → s3://{bucket}/updates/<name>-<ver>-full.nupkg
 *   out/make/squirrel.windows/x64/<name>Setup.exe          → s3://{bucket}/installer/<name>Setup.exe
 */

const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

// ── Helpers ────────────────────────────────────────────────────────────────────

function run(cmd, opts = {}) {
  console.log(`\n> ${cmd}`);
  execSync(cmd, { stdio: 'inherit', cwd: path.resolve(__dirname, '..'), ...opts });
}

function requireEnv(name) {
  const val = process.env[name];
  if (!val) {
    console.error(`\n  Missing required environment variable: ${name}`);
    console.error('  See the setup guide in UPDATING.md or scripts/publish-update.js header.\n');
    process.exit(1);
  }
  return val;
}

// ── Main ───────────────────────────────────────────────────────────────────────

const skipBuild = process.argv.includes('--upload');
const appDir = path.resolve(__dirname, '..');
const squirrelOut = path.join(appDir, 'out', 'make', 'squirrel.windows', 'x64');

// 1. Build (unless --upload)
if (!skipBuild) {
  console.log('\n=== Building MacroFlow ===');
  run('npm run make');
}

// 2. Verify output exists
if (!fs.existsSync(path.join(squirrelOut, 'RELEASES'))) {
  console.error(`\n  Squirrel output not found at: ${squirrelOut}`);
  console.error('  Run "npm run make" first, or remove the --upload flag.\n');
  process.exit(1);
}

// 3. Validate env vars
const endpoint = requireEnv('S3_ENDPOINT');
const bucket = requireEnv('S3_BUCKET');
requireEnv('AWS_ACCESS_KEY_ID');
requireEnv('AWS_SECRET_ACCESS_KEY');

// 4. List what we're uploading
const files = fs.readdirSync(squirrelOut);
const releasesFile = files.find((f) => f === 'RELEASES');
const nupkgFile = files.find((f) => f.endsWith('-full.nupkg'));
const setupExe = files.find((f) => f.toLowerCase().endsWith('setup.exe'));

if (!releasesFile || !nupkgFile) {
  console.error('\n  Missing RELEASES or .nupkg in Squirrel output.');
  process.exit(1);
}

console.log('\n=== Uploading to R2 ===');
console.log(`  Endpoint: ${endpoint}`);
console.log(`  Bucket:   ${bucket}`);
console.log(`  Files:`);
console.log(`    RELEASES         → updates/RELEASES`);
console.log(`    ${nupkgFile}     → updates/${nupkgFile}`);
if (setupExe) {
  console.log(`    ${setupExe}      → installer/${setupExe}`);
}

// 5. Upload using AWS CLI (s3-compatible)
const s3Flags = `--endpoint-url ${endpoint}`;

run(`aws s3 cp "${path.join(squirrelOut, 'RELEASES')}" s3://${bucket}/updates/RELEASES ${s3Flags}`);
run(`aws s3 cp "${path.join(squirrelOut, nupkgFile)}" s3://${bucket}/updates/${nupkgFile} ${s3Flags}`);

if (setupExe) {
  run(`aws s3 cp "${path.join(squirrelOut, setupExe)}" s3://${bucket}/installer/${setupExe} ${s3Flags}`);
}

// 6. Read version from package.json for summary
const pkg = JSON.parse(fs.readFileSync(path.join(appDir, 'package.json'), 'utf8'));
console.log(`\n=== Done ===`);
console.log(`  Version ${pkg.version} published.`);
console.log(`  Users running MacroFlow will pick up this update automatically.\n`);
