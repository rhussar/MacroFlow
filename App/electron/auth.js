/**
 * Auth0 authentication for MacroFlow.
 *
 * Uses Authorization Code + PKCE flow:
 *   1. Opens system browser to Auth0 login page
 *   2. Auth0 redirects to macroflow://auth/callback with an auth code
 *   3. Electron catches the deep link via custom protocol handler
 *   4. Exchanges the code for tokens
 *
 * No secrets are shipped in this code — PKCE eliminates the need for
 * a client secret in native/desktop apps.
 */

const { ipcMain, shell } = require('electron');
const https = require('node:https');
const crypto = require('node:crypto');
const logger = require('./logger');

// ── Auth0 configuration ─────────────────────────────────────────────────────

const AUTH0_DOMAIN = 'dev-viwrgtrnum1egndy.us.auth0.com';
const AUTH0_CLIENT_ID = 'nPfLjbCtAsObrLuHAxCOnNKXy3PgR973';
const AUTH0_REDIRECT_URI = 'macroflow://auth/callback';
const AUTH0_LOGOUT_REDIRECT_URI = 'macroflow://auth/logout';
const AUTH0_SCOPES = 'openid profile email offline_access';
const AUTH0_AUDIENCE = `https://${AUTH0_DOMAIN}/api/v2/`;

// ── PKCE helpers ────────────────────────────────────────────────────────────

function base64URLEncode(buffer) {
  return buffer.toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');
}

function generateCodeVerifier() {
  return base64URLEncode(crypto.randomBytes(32));
}

function generateCodeChallenge(verifier) {
  return base64URLEncode(crypto.createHash('sha256').update(verifier).digest());
}

// ── HTTP helper ─────────────────────────────────────────────────────────────

function httpsPost(url, body) {
  return new Promise((resolve, reject) => {
    const bodyStr = typeof body === 'string' ? body : new URLSearchParams(body).toString();
    const parsed = new URL(url);

    const req = https.request({
      hostname: parsed.hostname,
      path: parsed.pathname,
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(bodyStr),
      },
    }, (res) => {
      const chunks = [];
      res.on('data', (chunk) => chunks.push(chunk));
      res.on('end', () => {
        const raw = Buffer.concat(chunks).toString('utf8');
        try {
          resolve({ ok: res.statusCode >= 200 && res.statusCode < 300, json: JSON.parse(raw) });
        } catch {
          resolve({ ok: false, json: null });
        }
      });
    });
    req.on('error', reject);
    req.write(bodyStr);
    req.end();
  });
}

function httpsGet(url, headers = {}) {
  return new Promise((resolve, reject) => {
    const parsed = new URL(url);
    const req = https.request({
      hostname: parsed.hostname,
      path: parsed.pathname + parsed.search,
      method: 'GET',
      headers,
    }, (res) => {
      const chunks = [];
      res.on('data', (chunk) => chunks.push(chunk));
      res.on('end', () => {
        const raw = Buffer.concat(chunks).toString('utf8');
        try {
          resolve({ ok: res.statusCode >= 200 && res.statusCode < 300, json: JSON.parse(raw) });
        } catch {
          resolve({ ok: false, json: null });
        }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

// ── State ───────────────────────────────────────────────────────────────────

let codeVerifier = null;
let pendingResolve = null;
let ipcRegistered = false;

// ── Core ────────────────────────────────────────────────────────────────────

/**
 * Start the login flow — opens the system browser to Auth0.
 * Returns a Promise that resolves with { accessToken, idToken, profile }
 * when the callback is received.
 */
function startLogin() {
  return new Promise((resolve, reject) => {
    codeVerifier = generateCodeVerifier();
    const codeChallenge = generateCodeChallenge(codeVerifier);

    const params = new URLSearchParams({
      response_type: 'code',
      client_id: AUTH0_CLIENT_ID,
      redirect_uri: AUTH0_REDIRECT_URI,
      scope: AUTH0_SCOPES,
      audience: AUTH0_AUDIENCE,
      code_challenge: codeChallenge,
      code_challenge_method: 'S256',
    });

    const loginUrl = `https://${AUTH0_DOMAIN}/authorize?${params.toString()}`;

    pendingResolve = resolve;

    shell.openExternal(loginUrl).catch((err) => {
      pendingResolve = null;
      codeVerifier = null;
      logger.error('[Auth] failed to open browser', { error: err.message });
      reject(new Error('Could not open browser for sign-in.'));
    });

    logger.info('[Auth] login flow started');
  });
}

/**
 * Handle the macroflow://auth/callback deep link.
 * Called from the custom protocol handler in main.js.
 */
async function handleCallback(callbackUrl) {
  try {
    const url = new URL(callbackUrl);
    const code = url.searchParams.get('code');
    const error = url.searchParams.get('error');
    const errorDescription = url.searchParams.get('error_description');

    if (error) {
      logger.warn('[Auth] callback error', { error, errorDescription });
      if (pendingResolve) {
        pendingResolve({ success: false, message: errorDescription || error });
        pendingResolve = null;
      }
      return;
    }

    if (!code || !codeVerifier) {
      logger.warn('[Auth] callback missing code or verifier');
      if (pendingResolve) {
        pendingResolve({ success: false, message: 'Authentication failed — missing authorization code.' });
        pendingResolve = null;
      }
      return;
    }

    // Exchange authorization code for tokens.
    const tokenRes = await httpsPost(`https://${AUTH0_DOMAIN}/oauth/token`, {
      grant_type: 'authorization_code',
      client_id: AUTH0_CLIENT_ID,
      code_verifier: codeVerifier,
      code,
      redirect_uri: AUTH0_REDIRECT_URI,
    });

    codeVerifier = null;

    if (!tokenRes.ok || !tokenRes.json?.access_token) {
      logger.error('[Auth] token exchange failed', { response: tokenRes.json });
      if (pendingResolve) {
        pendingResolve({ success: false, message: 'Authentication failed — could not exchange code.' });
        pendingResolve = null;
      }
      return;
    }

    const { access_token, id_token } = tokenRes.json;

    // Fetch user profile.
    const profileRes = await httpsGet(`https://${AUTH0_DOMAIN}/userinfo`, {
      Authorization: `Bearer ${access_token}`,
    });

    const profile = profileRes.json || {};

    logger.info('[Auth] authenticated', {
      sub: profile.sub,
      email: profile.email,
    });

    if (pendingResolve) {
      pendingResolve({
        success: true,
        accessToken: access_token,
        idToken: id_token,
        profile,
      });
      pendingResolve = null;
    }
  } catch (err) {
    logger.error('[Auth] callback handling failed', { error: err.message });
    if (pendingResolve) {
      pendingResolve({ success: false, message: 'Authentication failed — ' + err.message });
      pendingResolve = null;
    }
  }
}

/**
 * Open the browser to Auth0's logout endpoint.
 */
function startLogout() {
  const params = new URLSearchParams({
    client_id: AUTH0_CLIENT_ID,
    returnTo: AUTH0_LOGOUT_REDIRECT_URI,
  });

  shell.openExternal(`https://${AUTH0_DOMAIN}/v2/logout?${params.toString()}`).catch((err) => {
    logger.error('[Auth] failed to open logout URL', { error: err.message });
  });

  logger.info('[Auth] logout initiated');
}

/**
 * Register IPC handlers for the renderer.
 */
function registerAuthHandlers() {
  if (ipcRegistered) return;
  ipcRegistered = true;

  ipcMain.handle('auth:login', async () => {
    try {
      const result = await startLogin();
      return result;
    } catch (err) {
      return { success: false, message: err.message };
    }
  });

  ipcMain.handle('auth:logout', async () => {
    startLogout();
    return { success: true };
  });
}

module.exports = {
  handleCallback,
  registerAuthHandlers,
  startLogin,
  startLogout,
};
