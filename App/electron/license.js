/**
 * License validation via Keygen.sh
 *
 * Security: NO product/admin tokens are shipped in this code.
 * All API calls use `Authorization: License {key}` — the license key
 * authenticates itself. This is Keygen's recommended approach for
 * client-side desktop apps.
 *
 * Flow:
 *   1. On startup, check for a cached license key
 *   2. Validate it against Keygen's API using the key itself as auth
 *   3. Cache the validation result so the app works offline for a grace period
 *   4. If no key or invalid, the renderer shows the activation screen
 *
 * Keygen API docs: https://keygen.sh/docs/api/
 */

const { ipcMain, app } = require('electron');
const path = require('node:path');
const fs = require('node:fs');
const https = require('node:https');
const logger = require('./logger');

// ── HTTP helper (Electron v22 doesn't have global fetch) ────────────────────

function httpsRequest(url, options, body) {
  return new Promise((resolve, reject) => {
    const req = https.request(url, options, (res) => {
      const chunks = [];
      res.on('data', (chunk) => chunks.push(chunk));
      res.on('end', () => {
        const raw = Buffer.concat(chunks).toString('utf8');
        let json;
        try {
          json = JSON.parse(raw);
        } catch {
          json = null;
        }
        resolve({ ok: res.statusCode >= 200 && res.statusCode < 300, status: res.statusCode, json });
      });
    });
    req.on('error', reject);
    if (body) req.write(body);
    req.end();
  });
}

async function keygenFetch(urlPath, method, headers, bodyObj) {
  const fullUrl = `${KEYGEN_API_URL}${urlPath}`;
  const bodyStr = bodyObj ? JSON.stringify(bodyObj) : null;

  const res = await httpsRequest(fullUrl, {
    method,
    headers: {
      ...headers,
      ...(bodyStr ? { 'Content-Length': Buffer.byteLength(bodyStr) } : {}),
    },
  }, bodyStr);

  return res;
}

// ── Keygen configuration ────────────────────────────────────────────────────

const KEYGEN_ACCOUNT_ID = 'e5b35ff4-71fb-4f0e-8d21-dcdb733121d3';
const KEYGEN_API_URL = `https://api.keygen.sh/v1/accounts/${KEYGEN_ACCOUNT_ID}`;

// How long a cached validation stays trusted when offline (days).
const OFFLINE_GRACE_DAYS = 7;

// ── Persistent storage ──────────────────────────────────────────────────────

const storePath = path.join(app.getPath('userData'), 'license.json');

function readStore() {
  try {
    if (fs.existsSync(storePath)) {
      return JSON.parse(fs.readFileSync(storePath, 'utf8'));
    }
  } catch (err) {
    logger.warn('[License] failed to read store', { error: err.message });
  }
  return {};
}

function writeStore(data) {
  try {
    fs.writeFileSync(storePath, JSON.stringify(data, null, 2), 'utf8');
  } catch (err) {
    logger.warn('[License] failed to write store', { error: err.message });
  }
}

// ── State ───────────────────────────────────────────────────────────────────

let licenseValid = false;
let licenseData = null; // { key, status, expiry, name, email, ... }
let ipcRegistered = false;
let startupCheckPromise = null;

// ── Helpers ─────────────────────────────────────────────────────────────────

/**
 * Generate a stable machine fingerprint.
 */
function getFingerprint() {
  const os = require('os');
  const crypto = require('crypto');
  const raw = `${os.hostname()}:${os.platform()}:${os.arch()}:${os.userInfo().username}`;
  return crypto.createHash('sha256').update(raw).digest('hex').slice(0, 32);
}

/**
 * Standard headers for Keygen API calls using license key auth.
 */
function keygenHeaders(licenseKey) {
  return {
    'Content-Type': 'application/vnd.api+json',
    Accept: 'application/vnd.api+json',
    Authorization: `License ${licenseKey}`,
  };
}

// ── Validation ──────────────────────────────────────────────────────────────

/**
 * Validate a license key against Keygen's API.
 * Uses the license key itself as authentication (no product token needed).
 */
async function validateWithKeygen(licenseKey) {
  const fingerprint = getFingerprint();

  const res = await keygenFetch('/licenses/actions/validate-key', 'POST', keygenHeaders(licenseKey), {
    meta: {
      key: licenseKey,
      scope: { fingerprints: [fingerprint] },
    },
  });

  const json = res.json;
  const valid = json?.meta?.valid === true;
  const detail = json?.meta?.detail || '';
  const code = json?.meta?.code || '';
  const license = json?.data?.attributes || {};
  const licenseId = json?.data?.id || null;

  return { valid, detail, code, license, licenseId };
}

/**
 * Activate a machine fingerprint for a license key.
 */
async function activateMachine(licenseId, licenseKey) {
  const fingerprint = getFingerprint();

  const res = await keygenFetch('/machines', 'POST', keygenHeaders(licenseKey), {
    data: {
      type: 'machines',
      attributes: {
        fingerprint,
        name: `${require('os').hostname()} (MacroFlow)`,
        platform: process.platform,
      },
      relationships: {
        license: { data: { type: 'licenses', id: licenseId } },
      },
    },
  });

  const json = res.json;
  if (res.ok) {
    logger.info('[License] machine activated', { fingerprint });
    return { success: true };
  }

  // Already activated is fine.
  const errorCode = json.errors?.[0]?.code;
  if (errorCode === 'MACHINE_UNIQUENESS_VIOLATION') {
    logger.info('[License] machine already activated', { fingerprint });
    return { success: true };
  }

  const detail = json.errors?.[0]?.detail || 'Unknown error';
  logger.warn('[License] machine activation failed', { detail });
  return { success: false, detail };
}

// ── Public API ──────────────────────────────────────────────────────────────

/**
 * Try to restore + validate the cached license on startup.
 */
async function checkCachedLicense() {
  const store = readStore();
  if (!store.key) {
    logger.info('[License] no cached license key');
    return { valid: false, licenseData: null };
  }

  try {
    const result = await validateWithKeygen(store.key);

    if (result.valid) {
      licenseValid = true;
      licenseData = {
        key: store.key,
        status: result.license.status,
        expiry: result.license.expiry,
        name: result.license.metadata?.name || store.name || '',
        email: result.license.metadata?.email || store.email || '',
        lastValidated: new Date().toISOString(),
      };
      writeStore({ ...licenseData });
      logger.info('[License] cached license valid', { status: result.license.status });
      return { valid: true, licenseData };
    }

    // If the fingerprint isn't activated yet, try to activate automatically.
    if (result.code === 'FINGERPRINT_SCOPE_MISMATCH' || result.code === 'NO_MACHINES' || result.code === 'NO_MACHINE') {
      if (result.licenseId) {
        const activation = await activateMachine(result.licenseId, store.key);
        if (activation.success) {
          const recheck = await validateWithKeygen(store.key);
          if (recheck.valid) {
            licenseValid = true;
            licenseData = {
              key: store.key,
              status: recheck.license.status,
              expiry: recheck.license.expiry,
              name: recheck.license.metadata?.name || store.name || '',
              email: recheck.license.metadata?.email || store.email || '',
              lastValidated: new Date().toISOString(),
            };
            writeStore({ ...licenseData });
            logger.info('[License] auto-activated and valid');
            return { valid: true, licenseData };
          }
        }
      }
    }

    // Online validation failed — check offline grace period.
    if (store.lastValidated) {
      const lastValid = new Date(store.lastValidated);
      const graceMs = OFFLINE_GRACE_DAYS * 24 * 60 * 60 * 1000;
      if (Date.now() - lastValid.getTime() < graceMs) {
        licenseValid = true;
        licenseData = store;
        logger.info('[License] using cached validation (offline grace)', {
          detail: result.detail,
        });
        return { valid: true, licenseData };
      }
    }

    logger.info('[License] cached license invalid', { detail: result.detail });
    return { valid: false, licenseData: null, detail: result.detail };
  } catch (err) {
    // Network error — fall back to offline grace.
    if (store.lastValidated) {
      const lastValid = new Date(store.lastValidated);
      const graceMs = OFFLINE_GRACE_DAYS * 24 * 60 * 60 * 1000;
      if (Date.now() - lastValid.getTime() < graceMs) {
        licenseValid = true;
        licenseData = store;
        logger.info('[License] offline — using cached validation', { error: err.message });
        return { valid: true, licenseData };
      }
    }
    logger.warn('[License] validation failed (network)', { error: err.message });
    return { valid: false, licenseData: null, detail: 'Network error — could not validate license.' };
  }
}

/**
 * Activate a new license key (called from the renderer activation screen).
 */
async function activateLicense(licenseKey) {
  try {
    // Validate to get the license ID and status.
    const result = await validateWithKeygen(licenseKey);

    if (!result.licenseId) {
      return {
        success: false,
        message: result.detail || 'Invalid license key.',
      };
    }

    // Check if the license is in a usable state.
    const status = result.license.status;
    if (status === 'EXPIRED') {
      return { success: false, message: 'This license has expired.' };
    }
    if (status === 'SUSPENDED') {
      return { success: false, message: 'This license has been suspended.' };
    }
    if (status === 'BANNED') {
      return { success: false, message: 'This license has been revoked.' };
    }

    // If already valid (fingerprint matches), we're done.
    if (result.valid) {
      licenseValid = true;
      licenseData = {
        key: licenseKey,
        status: result.license.status,
        expiry: result.license.expiry,
        name: result.license.metadata?.name || '',
        email: result.license.metadata?.email || '',
        lastValidated: new Date().toISOString(),
      };
      writeStore({ ...licenseData });
      logger.info('[License] activated successfully (already valid)');
      return { success: true, licenseData };
    }

    // Activate this machine.
    const activation = await activateMachine(result.licenseId, licenseKey);
    if (!activation.success) {
      return { success: false, message: activation.detail || 'Could not activate this device.' };
    }

    // Re-validate with fingerprint scope.
    const recheck = await validateWithKeygen(licenseKey);
    if (!recheck.valid) {
      return {
        success: false,
        message: recheck.detail || 'License validation failed after activation.',
      };
    }

    licenseValid = true;
    licenseData = {
      key: licenseKey,
      status: recheck.license.status,
      expiry: recheck.license.expiry,
      name: recheck.license.metadata?.name || '',
      email: recheck.license.metadata?.email || '',
      lastValidated: new Date().toISOString(),
    };
    writeStore({ ...licenseData });
    logger.info('[License] activated successfully');

    return { success: true, licenseData };
  } catch (err) {
    logger.error('[License] activation error', { error: err.message });
    return { success: false, message: 'Network error — please check your connection.' };
  }
}

/**
 * Activate a license via Auth0 login.
 * After Auth0 authentication, the Auth0 Action stores the Keygen license key
 * in the user's ID token claims (app_metadata.license_key). This function
 * reads that key and activates it on this machine.
 */
async function activateFromAuth(authResult) {
  if (!authResult?.success) {
    return { success: false, message: authResult?.message || 'Authentication failed.' };
  }

  const profile = authResult.profile || {};
  // The Auth0 Action stores the license key in a custom claim namespace.
  const licenseKey =
    profile['https://macroflow.com/license_key'] ||
    profile.license_key ||
    null;

  if (!licenseKey) {
    logger.warn('[License] Auth0 profile has no license key', { sub: profile.sub });
    return {
      success: false,
      message: 'No license found for this account. Please contact support.',
    };
  }

  logger.info('[License] activating from Auth0', { sub: profile.sub });
  const result = await activateLicense(licenseKey);

  // Store the Auth0 user info alongside the license data.
  if (result.success && licenseData) {
    licenseData.auth0Sub = profile.sub || '';
    licenseData.email = profile.email || licenseData.email || '';
    licenseData.name = profile.name || licenseData.name || '';
    writeStore({ ...licenseData });
  }

  return result;
}

/**
 * Deactivate / remove the license from this machine.
 */
async function deactivateLicense() {
  licenseValid = false;
  licenseData = null;
  writeStore({});
  logger.info('[License] deactivated');
  return { success: true };
}

/**
 * Register IPC handlers for the renderer.
 */
function registerLicenseHandlers() {
  if (ipcRegistered) return;
  ipcRegistered = true;

  ipcMain.handle('license:status', async () => {
    // Wait for startup check to complete before responding
    if (startupCheckPromise) {
      await startupCheckPromise;
    }
    return {
      valid: licenseValid,
      licenseData,
    };
  });

  ipcMain.handle('license:activate', async (_event, licenseKey) => {
    return activateLicense(licenseKey);
  });

  ipcMain.handle('license:activate-from-auth', async (_event, authResult) => {
    return activateFromAuth(authResult);
  });

  ipcMain.handle('license:deactivate', async () => {
    return deactivateLicense();
  });

  ipcMain.handle('license:check', async () => {
    return checkCachedLicense();
  });
}

function isLicenseValid() {
  return licenseValid;
}

function setStartupCheckPromise(promise) {
  startupCheckPromise = promise;
}

module.exports = {
  checkCachedLicense,
  registerLicenseHandlers,
  isLicenseValid,
  setStartupCheckPromise,
};
